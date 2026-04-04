/* =========================
   0) Schema namespace
   ========================= */
IF SCHEMA_ID('acct') IS NULL
    EXEC('CREATE SCHEMA acct');
GO

/* =========================
   1) Lookup: account types
   ========================= */
IF OBJECT_ID('acct.lkp_account_type', 'U') IS NULL
BEGIN
    CREATE TABLE acct.lkp_account_type (
        account_type_id TINYINT NOT NULL PRIMARY KEY,   -- 1..5
        account_type_code VARCHAR(20) NOT NULL UNIQUE,  -- ASSET, LIABILITY, EQUITY, INCOME, EXPENSE
        normal_side CHAR(1) NOT NULL CHECK (normal_side IN ('D','C')) -- D=debit normal, C=credit normal
    );

    INSERT INTO acct.lkp_account_type(account_type_id, account_type_code, normal_side)
    VALUES
      (1, 'ASSET',     'D'),
      (2, 'LIABILITY', 'C'),
      (3, 'EQUITY',    'C'),
      (4, 'INCOME',    'C'),
      (5, 'EXPENSE',   'D');
END
GO

/* =========================
   2) Chart of accounts
   ========================= */
IF OBJECT_ID('acct.dim_account', 'U') IS NULL
BEGIN
    CREATE TABLE acct.dim_account (
        account_id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        account_code INT NOT NULL UNIQUE,                   -- e.g. 1010, 4030
        account_name NVARCHAR(200) NOT NULL,
        account_type_id TINYINT NOT NULL
            CONSTRAINT FK_dim_account_type REFERENCES acct.lkp_account_type(account_type_id),
        is_active BIT NOT NULL CONSTRAINT DF_dim_account_active DEFAULT (1),
        created_at DATETIME2(0) NOT NULL CONSTRAINT DF_dim_account_created DEFAULT (SYSUTCDATETIME()),
        updated_at DATETIME2(0) NULL
    );
END
GO

/* =========================
   3) Reporting periods
   ========================= */
IF OBJECT_ID('acct.dim_period', 'U') IS NULL
BEGIN
    CREATE TABLE acct.dim_period (
        period_id INT NOT NULL PRIMARY KEY,                 -- YYYYMM, e.g. 201601
        period_start_date DATE NOT NULL UNIQUE,             -- first day of month
        period_end_date DATE NOT NULL,                      -- last day of month
        [year] SMALLINT NOT NULL,
        [month] TINYINT NOT NULL CHECK ([month] BETWEEN 1 AND 12),
        is_closed BIT NOT NULL CONSTRAINT DF_dim_period_closed DEFAULT (0),
        CONSTRAINT CK_dim_period_id
            CHECK (period_id = ([year] * 100 + [month]))
    );
END
GO

/* =========================
   4) ETL batch metadata
   ========================= */
IF OBJECT_ID('acct.etl_batch', 'U') IS NULL
BEGIN
    CREATE TABLE acct.etl_batch (
        batch_id BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        source_type VARCHAR(30) NOT NULL,                   -- CHART_OF_ACCOUNTS / COLLECTION_PAYMENT
        file_name NVARCHAR(260) NOT NULL,
        file_checksum VARBINARY(32) NULL,                   -- SHA-256 from Python
        period_id INT NULL
            CONSTRAINT FK_etl_batch_period REFERENCES acct.dim_period(period_id),
        loaded_at DATETIME2(0) NOT NULL CONSTRAINT DF_etl_batch_loaded DEFAULT (SYSUTCDATETIME()),
        status VARCHAR(20) NOT NULL CONSTRAINT DF_etl_batch_status DEFAULT ('LOADED'),
        row_count INT NULL,
        error_message NVARCHAR(2000) NULL
    );

    CREATE INDEX IX_etl_batch_period ON acct.etl_batch(period_id);
END
GO

/* =========================
   5) Journal entry header
   ========================= */
IF OBJECT_ID('acct.journal_entry', 'U') IS NULL
BEGIN
    CREATE TABLE acct.journal_entry (
        journal_entry_id BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        period_id INT NOT NULL
            CONSTRAINT FK_journal_entry_period REFERENCES acct.dim_period(period_id),
        transaction_date DATE NOT NULL,
        description NVARCHAR(500) NULL,
        source_batch_id BIGINT NOT NULL
            CONSTRAINT FK_journal_entry_batch REFERENCES acct.etl_batch(batch_id),
        source_row_number INT NULL,                         -- row number in input sheet
        source_row_hash VARBINARY(32) NULL,                -- deterministic hash for idempotency
        created_at DATETIME2(0) NOT NULL CONSTRAINT DF_journal_entry_created DEFAULT (SYSUTCDATETIME())
    );

    CREATE UNIQUE INDEX UX_journal_entry_batch_row
        ON acct.journal_entry(source_batch_id, source_row_number)
        WHERE source_row_number IS NOT NULL;

    CREATE UNIQUE INDEX UX_journal_entry_row_hash
        ON acct.journal_entry(source_row_hash)
        WHERE source_row_hash IS NOT NULL;

    CREATE INDEX IX_journal_entry_period_date
        ON acct.journal_entry(period_id, transaction_date);
END
GO

/* =========================
   6) Journal entry lines
   ========================= */
IF OBJECT_ID('acct.journal_entry_line', 'U') IS NULL
BEGIN
    CREATE TABLE acct.journal_entry_line (
        journal_entry_line_id BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        journal_entry_id BIGINT NOT NULL
            CONSTRAINT FK_journal_entry_line_header
            REFERENCES acct.journal_entry(journal_entry_id) ON DELETE CASCADE,
        line_no TINYINT NOT NULL,                           -- 1..n within entry
        account_id INT NOT NULL
            CONSTRAINT FK_journal_entry_line_account REFERENCES acct.dim_account(account_id),
        dr_cr CHAR(1) NOT NULL CHECK (dr_cr IN ('D','C')),
        amount DECIMAL(19,2) NOT NULL CHECK (amount > 0),
        memo NVARCHAR(500) NULL
    );

    CREATE UNIQUE INDEX UX_journal_entry_line_no
        ON acct.journal_entry_line(journal_entry_id, line_no);

    CREATE INDEX IX_journal_entry_line_account
        ON acct.journal_entry_line(account_id);
END
GO

/* =========================
   7) Optional monthly snapshot
   (for fast reporting/caching)
   ========================= */
IF OBJECT_ID('acct.fact_trial_balance_monthly', 'U') IS NULL
BEGIN
    CREATE TABLE acct.fact_trial_balance_monthly (
        period_id INT NOT NULL
            CONSTRAINT FK_tb_period REFERENCES acct.dim_period(period_id),
        account_id INT NOT NULL
            CONSTRAINT FK_tb_account REFERENCES acct.dim_account(account_id),
        opening_balance DECIMAL(19,2) NOT NULL,             -- signed
        period_debits DECIMAL(19,2) NOT NULL,
        period_credits DECIMAL(19,2) NOT NULL,
        closing_balance DECIMAL(19,2) NOT NULL,             -- signed
        calculated_at DATETIME2(0) NOT NULL CONSTRAINT DF_tb_calc DEFAULT (SYSUTCDATETIME()),
        CONSTRAINT PK_fact_trial_balance_monthly PRIMARY KEY (period_id, account_id)
    );
END
GO

