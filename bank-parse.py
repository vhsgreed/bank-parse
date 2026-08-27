#!/usr/bin/env python3
"""Parse Handelsbanken xlsx exports -> unified transactions CSV.

Usage: python3 scripts/bank-parse.py finance/data/*.xlsx --out finance/data/transactions.csv
"""
import argparse, csv, re
from datetime import date
from pathlib import Path

def parse_xlsx(path: Path) -> list[dict]:
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    out = []
    for sn in wb.sheetnames:
        ws = wb[sn]
        rows = list(ws.iter_rows(values_only=True))
        # find header row: Reskontradatum
        hdr_i = next((i for i, r in enumerate(rows) if r and r[0] == "Reskontradatum"), None)
        if hdr_i is None:
            continue
        # account name: scan above header for the line that is the account title (not Saldo/Period/Kontoform)
        account = None
        for r in rows[:hdr_i]:
            for c in r:
                if c and isinstance(c, str) and re.search(r"\d{3} \d{3} \d{3}", c):
                    account = c.strip()
                    break
            if account: break
        for r in rows[hdr_i + 1:]:
            if not r or not r[0]: continue
            out.append({
                "account": account or path.stem,
                "reskontra": str(r[0]) if r[0] else "",
                "trans_date": str(r[1]) if r[1] else "",
                "text": str(r[2]).strip() if r[2] else "",
                "amount": float(r[3]) if r[3] is not None else 0.0,
                "balance": float(r[4]) if r[4] is not None else None,
            })
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--out", default="finance/data/transactions.csv")
    args = ap.parse_args()
    rows = []
    for f in args.files:
        rows += parse_xlsx(Path(f))
    # dedupe identical (account, trans_date, text, amount)
    seen, uniq = set(), []
    for r in rows:
        k = (r["account"], r["trans_date"], r["text"], r["amount"])
        if k in seen: continue
        seen.add(k); uniq.append(r)
    uniq.sort(key=lambda r: (r["trans_date"], r["account"]))
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=["account", "reskontra", "trans_date", "text", "amount", "balance"])
        w.writeheader(); w.writerows(uniq)
    print(f"{len(uniq)} transactions ({len(rows)} raw) -> {out}")

if __name__ == "__main__":
    main()
