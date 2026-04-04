from __future__ import annotations

import argparse
import hashlib
import os
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable

import pandas as pd
import pyodbc
from dotenv import load_dotenv


@dataclass
class EntryRow:
    excel_row_number: int
    transaction_date: date
    description: str
    amount: Decimal
    debit_account_code: int
    credit_account_code: int
    period_id: int


def parse_amount(value: object) -> Decimal:
    """Parse values like 100.00, ' 1,040.00 ', '(2,235.14)' into positive Decimal."""
    if value is None:
        raise ValueError("Amount is missing")

    text = str(value).strip().replace(",", "")
    if text == "":
        raise ValueError("Amount is blank")

    negative_parentheses = text.startswith("(") and text.endswith(")")
    text = text.replace("(", "").replace(")", "")

    try:
        amount = Decimal(text)
    except InvalidOperation as ex:
        raise ValueError(f"Invalid amount: {value!r}") from ex

    if negative_parentheses:
        amount = -amount

    return abs(amount)


def parse_date_yyyymmdd(value: object) -> date:
    if value is None:
        raise ValueError("Date is missing")

    text = str(value).strip()
    if not text:
        raise ValueError("Date is blank")

    # Support Excel numeric style and plain yyyymmdd.
    if "." in text:
        text = text.split(".")[0]

    if len(text) != 8 or not text.isdigit():
        raise ValueError(f"Invalid date format (expected YYYYMMDD): {value!r}")

    return datetime.strptime(text, "%Y%m%d").date()


def normalize_account_code(value: object, label: str) -> int:
    if value is None:
        raise ValueError(f"{label} is missing")

    text = str(value).strip()
    if "." in text:
        text = text.split(".")[0]

    if not text.isdigit():
        raise ValueError(f"Invalid {label}: {value!r}")

    return int(text)


def compute_file_checksum(path: Path) -> bytes:
    hasher = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.digest()


def compute_row_hash(row: EntryRow) -> bytes:
    payload = "|".join(
        [
            row.transaction_date.isoformat(),
            row.description.strip().lower(),
            f"{row.amount:.2f}",
            str(row.debit_account_code),
            str(row.credit_account_code),
            str(row.period_id),
        ]
    )
    return hashlib.sha256(payload.encode("utf-8")).digest()


def get_connection() -> pyodbc.Connection:
    load_dotenv()

    required = ["DRIVER", "SERVER", "DATABASE", "ENTRA_USER", "ENTRA_PASS"]
    missing = [k for k in required if not os.getenv(k)]
    if missing:
        raise RuntimeError(f"Missing environment variables: {', '.join(missing)}")

    conn_str = (
        f"Driver={{{os.getenv('DRIVER')}}};"
        f"Server=tcp:{os.getenv('SERVER')},1433;"
        f"Database={os.getenv('DATABASE')};"
        f"UID={os.getenv('ENTRA_USER')};"
        f"PWD={os.getenv('ENTRA_PASS')};"
        "Encrypt=yes;"
        "TrustServerCertificate=no;"
        "Connection Timeout=60;"
        "Authentication=ActiveDirectoryPassword;"
    )
    return pyodbc.connect(conn_str)


def read_chart_of_accounts(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    required = {"Code", "Account", "Type"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Chart_of_Accounts missing columns: {sorted(missing)}")

    out = df[["Code", "Account", "Type"]].copy()
    out["Code"] = out["Code"].apply(lambda v: normalize_account_code(v, "Code"))
    out["Account"] = out["Account"].astype(str).str.strip()
    out["Type"] = out["Type"].astype(str).str.strip().str.upper()
    return out


def read_collection_and_payment(path: Path) -> list[EntryRow]:
    df = pd.read_excel(path)

    required = {"Date", "Transaction Description", "Ammount", "Debit Account", "Credit Account"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Collection_and_Payment missing columns: {sorted(missing)}")

    rows: list[EntryRow] = []
    for idx, record in df.iterrows():
        excel_row_number = int(idx) + 2  # +1 for 1-based index and +1 for header row

        transaction_date = parse_date_yyyymmdd(record["Date"])
        description = str(record["Transaction Description"] or "").strip()
        if not description:
            description = "(no description)"

        amount = parse_amount(record["Ammount"])
        debit_account_code = normalize_account_code(record["Debit Account"], "Debit Account")
        credit_account_code = normalize_account_code(record["Credit Account"], "Credit Account")
        period_id = transaction_date.year * 100 + transaction_date.month

        rows.append(
            EntryRow(
                excel_row_number=excel_row_number,
                transaction_date=transaction_date,
                description=description,
                amount=amount,
                debit_account_code=debit_account_code,
                credit_account_code=credit_account_code,
                period_id=period_id,
            )
        )

    return rows


def ensure_periods(cursor: pyodbc.Cursor, periods: Iterable[int]) -> None:
    for period_id in sorted(set(periods)):
        year = period_id // 100
        month = period_id % 100
        start = date(year, month, 1)
        if month == 12:
            end = date(year + 1, 1, 1) - pd.Timedelta(days=1)
        else:
            end = date(year, month + 1, 1) - pd.Timedelta(days=1)

        cursor.execute(
            """
IF NOT EXISTS (SELECT 1 FROM acct.dim_period WHERE period_id = ?)
BEGIN
    INSERT INTO acct.dim_period (period_id, period_start_date, period_end_date, [year], [month], is_closed)
    VALUES (?, ?, ?, ?, ?, 0)
END
""",
            period_id,
            period_id,
            start,
            end,
            year,
            month,
        )


def load_chart_of_accounts(cursor: pyodbc.Cursor, chart_df: pd.DataFrame) -> None:
    account_type_map = {
        row.account_type_code: row.account_type_id
        for row in cursor.execute(
            "SELECT account_type_id, account_type_code FROM acct.lkp_account_type"
        ).fetchall()
    }

    for _, row in chart_df.iterrows():
        account_code = int(row["Code"])
        account_name = str(row["Account"]).strip()
        type_code = str(row["Type"]).strip().upper()

        if type_code not in account_type_map:
            raise ValueError(f"Unknown account type in chart: {type_code}")

        account_type_id = account_type_map[type_code]

        cursor.execute(
            """
IF EXISTS (SELECT 1 FROM acct.dim_account WHERE account_code = ?)
BEGIN
    UPDATE acct.dim_account
       SET account_name = ?,
           account_type_id = ?,
           is_active = 1,
           updated_at = SYSUTCDATETIME()
     WHERE account_code = ?
END
ELSE
BEGIN
    INSERT INTO acct.dim_account (account_code, account_name, account_type_id, is_active)
    VALUES (?, ?, ?, 1)
END
""",
            account_code,
            account_name,
            account_type_id,
            account_code,
            account_code,
            account_name,
            account_type_id,
        )


def create_etl_batch(
    cursor: pyodbc.Cursor,
    file_name: str,
    file_checksum: bytes,
    period_id: int | None,
) -> int:
    cursor.execute(
        """
INSERT INTO acct.etl_batch (source_type, file_name, file_checksum, period_id, status)
OUTPUT INSERTED.batch_id
VALUES ('COLLECTION_PAYMENT', ?, ?, ?, 'LOADED');
""",
        file_name,
        pyodbc.Binary(file_checksum),
        period_id,
    )
    return int(cursor.fetchone()[0])


def get_account_id_map(cursor: pyodbc.Cursor) -> dict[int, int]:
    return {int(r.account_code): int(r.account_id) for r in cursor.execute("SELECT account_id, account_code FROM acct.dim_account").fetchall()}


def insert_journal_entries(cursor: pyodbc.Cursor, batch_id: int, rows: list[EntryRow]) -> tuple[int, int]:
    account_map = get_account_id_map(cursor)
    inserted = 0
    skipped_duplicates = 0

    for row in rows:
        if row.debit_account_code not in account_map:
            raise ValueError(f"Debit account code not found in dim_account: {row.debit_account_code}")
        if row.credit_account_code not in account_map:
            raise ValueError(f"Credit account code not found in dim_account: {row.credit_account_code}")

        row_hash = compute_row_hash(row)

        cursor.execute(
            "SELECT 1 FROM acct.journal_entry WHERE source_row_hash = ?",
            pyodbc.Binary(row_hash),
        )
        if cursor.fetchone() is not None:
            skipped_duplicates += 1
            continue

        cursor.execute(
            """
INSERT INTO acct.journal_entry (
    period_id,
    transaction_date,
    description,
    source_batch_id,
    source_row_number,
    source_row_hash
)
OUTPUT INSERTED.journal_entry_id
VALUES (?, ?, ?, ?, ?, ?);
""",
            row.period_id,
            row.transaction_date,
            row.description,
            batch_id,
            row.excel_row_number,
            pyodbc.Binary(row_hash),
        )
        journal_entry_id = int(cursor.fetchone()[0])

        cursor.execute(
            """
INSERT INTO acct.journal_entry_line (journal_entry_id, line_no, account_id, dr_cr, amount, memo)
VALUES (?, 1, ?, 'D', ?, ?),
       (?, 2, ?, 'C', ?, ?)
""",
            journal_entry_id,
            account_map[row.debit_account_code],
            row.amount,
            row.description,
            journal_entry_id,
            account_map[row.credit_account_code],
            row.amount,
            row.description,
        )
        inserted += 1

    return inserted, skipped_duplicates


def set_batch_result(cursor: pyodbc.Cursor, batch_id: int, status: str, row_count: int | None, error_message: str | None = None) -> None:
    cursor.execute(
        """
UPDATE acct.etl_batch
   SET status = ?,
       row_count = ?,
       error_message = ?
 WHERE batch_id = ?
""",
        status,
        row_count,
        error_message,
        batch_id,
    )


def run_loader(chart_path: Path, collection_path: Path) -> None:
    if not chart_path.exists():
        raise FileNotFoundError(f"Chart file not found: {chart_path}")
    if not collection_path.exists():
        raise FileNotFoundError(f"Collection file not found: {collection_path}")

    chart_df = read_chart_of_accounts(chart_path)
    rows = read_collection_and_payment(collection_path)

    file_checksum = compute_file_checksum(collection_path)
    period_ids = sorted({r.period_id for r in rows})
    batch_period = period_ids[0] if len(period_ids) == 1 else None

    conn = get_connection()
    conn.autocommit = False
    cursor = conn.cursor()

    batch_id: int | None = None
    try:
        ensure_periods(cursor, period_ids)
        load_chart_of_accounts(cursor, chart_df)
        batch_id = create_etl_batch(
            cursor=cursor,
            file_name=collection_path.name,
            file_checksum=file_checksum,
            period_id=batch_period,
        )

        inserted, skipped = insert_journal_entries(cursor, batch_id, rows)
        set_batch_result(cursor, batch_id, status="LOADED", row_count=inserted)

        conn.commit()
        print(
            f"✅ Load complete. Batch {batch_id}. "
            f"Rows in file: {len(rows)} | Journal entries inserted: {inserted} | Duplicates skipped: {skipped}"
        )

    except Exception as ex:
        conn.rollback()
        if batch_id is not None:
            try:
                # record failure in separate transaction
                cursor.execute("BEGIN TRANSACTION")
                set_batch_result(cursor, batch_id, status="FAILED", row_count=None, error_message=str(ex)[:2000])
                cursor.execute("COMMIT TRANSACTION")
            except Exception:
                cursor.execute("ROLLBACK TRANSACTION")
        raise
    finally:
        cursor.close()
        conn.close()


def build_default_paths() -> tuple[Path, Path]:
    # Script lives in ROC/. Input files live one folder up.
    root = Path(__file__).resolve().parent.parent
    return root / "Chart_of_Accounts.xlsx", root / "Collection_and_Payment.xlsx"


def main() -> None:
    default_chart, default_collection = build_default_paths()

    parser = argparse.ArgumentParser(
        description="Load Chart_of_Accounts + monthly Collection_and_Payment XLSX into SQL ledger tables."
    )
    parser.add_argument("--chart", type=Path, default=default_chart, help="Path to Chart_of_Accounts.xlsx")
    parser.add_argument(
        "--collection",
        type=Path,
        default=default_collection,
        help="Path to Collection_and_Payment.xlsx",
    )
    args = parser.parse_args()

    run_loader(args.chart, args.collection)


if __name__ == "__main__":
    main()
