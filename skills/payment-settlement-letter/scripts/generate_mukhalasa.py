#!/usr/bin/env python3
"""Generate a payment settlement letter (خطاب مخالصة مالية) for a contractor
working on a cost-plus-fee contract (تكلفة + نسبة), from a JSON spec.

Usage:
    python generate_mukhalasa.py spec.json -o "settlement-4.docx" \
        [--previous previous_settlement.docx] [--company company.json] \
        [--template template.docx]

Without --template a neutral template is generated from company.json.

Spec JSON:
{
  "date": "2026-07-26",              // optional, defaults to today
  "payment_number": 4,
  "payment_amount": 30000.0,
  "fee_percent": 15,                  // optional, default 15
  "project": "اسم العميل – المدينة - الحي",
  "client_name": "اسم العميل",
  "expenses": [{"desc": "...", "amount": 8500.0}, ...],
  "previous_payments": [              // optional if --previous is given
    {"num": 1, "amount": 30000, "fee": 3913.04, "expenses": 27631.0}
  ]
}

The fee is computed as: fee = payment_amount * P / (100 + P)
(the payment includes cost + P%, so the fee is P% of the cost portion).
"""
import argparse
import copy
import datetime
import json
import re
import sys

from docx import Document

RIYAL = "ر.س"
LRM = "‎"


def fmt(x, dec=2):
    """Format a number Arabic-accounting style: 1,234.56 / (1,234.56) for negatives."""
    s = f"{abs(x):,.{dec}f}"
    return f"({s})" if x < -1e-9 else s


def money(x, dec=2):
    return f"{fmt(x, dec)} {RIYAL}"


def smart_dec(x):
    """No decimals for whole numbers (matches the template's cumulative table)."""
    return 0 if abs(x - round(x)) < 0.005 else 2


def parse_money(s):
    """Parse '3,913.04 ر.س' / '(1,545.00) ر.س' back into a float."""
    s = s.replace(RIYAL, "").replace(LRM, "").replace(",", "").strip()
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").strip()
    v = float(s)
    return -v if neg else v


def set_para_text(para, text):
    """Replace paragraph text while keeping the first run's formatting."""
    if para.runs:
        para.runs[0].text = text
        for r in para.runs[1:]:
            r.text = ""
    else:
        para.add_run(text)


def set_cell_text(cell, text):
    set_para_text(cell.paragraphs[0], text)
    for p in cell.paragraphs[1:]:
        set_para_text(p, "")


def clone_row_after(table, proto_row, before_tr=None):
    """Deep-copy a row. Insert before before_tr if given, else append."""
    new_tr = copy.deepcopy(proto_row._tr)
    if before_tr is not None:
        before_tr.addprevious(new_tr)
    else:
        proto_row._tr.addnext(new_tr)
    from docx.table import _Row
    return _Row(new_tr, table)


def extract_previous_payments(path):
    """Read the cumulative table from a previous mukhalasa docx."""
    doc = Document(path)
    for t in doc.tables:
        header = [c.text for c in t.rows[0].cells]
        if any("الرصيد التراكمي" in h for h in header):
            rows = []
            for r in t.rows[1:]:
                cells = [c.text.strip() for c in r.cells]
                if not cells[0].strip():
                    continue
                rows.append({
                    "num": int(re.sub(r"\D", "", cells[0]) or 0),
                    "amount": parse_money(cells[1]),
                    "fee": parse_money(cells[2]),
                    "expenses": parse_money(cells[3]),
                    # keep the previously published figures verbatim so the new
                    # letter never contradicts what the client already received
                    "net": parse_money(cells[4]),
                    "balance": parse_money(cells[5]),
                })
            return rows
    raise SystemExit("لم يتم العثور على جدول الكشف التراكمي في الملف السابق")


def pct_label(p):
    p_str = str(int(p)) if float(p).is_integer() else str(p)
    return f"{LRM}{p_str}%{LRM}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("--previous", help="previous mukhalasa .docx to pull the cumulative table from")
    ap.add_argument("--template", default=None)
    ap.add_argument("--company", default=None,
                    help="company.json: name / signer / signer_title / tagline / footer_lines")
    args = ap.parse_args()
    company = {"name": "اسم شركتك", "signer": "اسم المدير", "signer_title": "المدير العام"}
    if args.company:
        with open(args.company, encoding="utf-8") as f:
            company.update(json.load(f))

    with open(args.spec, encoding="utf-8") as f:
        spec = json.load(f)

    import os
    import tempfile
    template = args.template
    if not template:
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import make_template
        template = os.path.join(tempfile.mkdtemp(), "template.docx")
        make_template.build(company, template)

    amount = float(spec["payment_amount"])
    p = float(spec.get("fee_percent", 15))
    num = int(spec["payment_number"])
    fee = round(amount * p / (100 + p), 2)
    available = round(amount - fee, 2)
    expenses = spec["expenses"]
    total_exp = round(sum(float(e["amount"]) for e in expenses), 2)
    net = round(available - total_exp, 2)
    date = spec.get("date") or datetime.date.today().isoformat()

    prev = spec.get("previous_payments") or []
    if args.previous:
        prev = extract_previous_payments(args.previous)
    prev = sorted(prev, key=lambda r: r["num"])

    doc = Document(template)
    paras = doc.paragraphs

    def find_para(substr):
        for para in paras:
            if substr in para.text:
                return para
        raise SystemExit(f"لم يتم العثور على الفقرة: {substr}")

    set_para_text(find_para("التاريخ:"), f"التاريخ: {date}")
    set_para_text(find_para("الموضوع:"),
                  f"الموضوع: خطاب مخالصة مالية – الدفعة رقم ({num})")
    set_para_text(find_para("المشروع:"), f"المشروع: {spec['project']}")

    # The closing paragraph (رابعاً) also hardcodes the payment number
    closing = find_para("الدفعة الحالية رقم")
    set_para_text(closing, re.sub(r"رقم \(\d+\)", f"رقم ({num})", closing.text))

    # Client line: keep 'الأستاذ' + tabs + 'المحترم' runs, swap only the name run
    client_para = find_para("الأستاذ")
    for r in client_para.runs:
        if "/" in r.text:
            r.text = f" / {spec['client_name']}  "
            break

    # --- Table 1: financial summary ---
    t0 = doc.tables[0]
    set_cell_text(t0.rows[0].cells[1], money(amount))
    set_cell_text(t0.rows[1].cells[0], f"أتعاب إدارة المشروع ({pct_label(p)})")
    set_cell_text(t0.rows[1].cells[1], money(fee))
    set_cell_text(t0.rows[2].cells[1], money(available))
    set_cell_text(t0.rows[3].cells[1], money(total_exp))
    set_cell_text(t0.rows[4].cells[1], money(net))

    # --- Table 2: expense details (header / N data rows / total) ---
    t1 = doc.tables[1]
    proto = t1.rows[1]
    total_row = t1.rows[-1]
    for r in list(t1.rows[2:-1]):
        r._tr.getparent().remove(r._tr)
    # fill the prototype with the first expense, clone for the rest
    set_cell_text(proto.cells[0], expenses[0]["desc"])
    set_cell_text(proto.cells[1], money(float(expenses[0]["amount"])))
    for e in expenses[1:]:
        new = clone_row_after(t1, proto, before_tr=total_row._tr)
        set_cell_text(new.cells[0], e["desc"])
        set_cell_text(new.cells[1], money(float(e["amount"])))
    set_cell_text(total_row.cells[0], "الإجمالي")
    set_cell_text(total_row.cells[1], money(total_exp))

    # --- Table 3: cumulative statement ---
    t2 = doc.tables[2]
    set_cell_text(t2.rows[0].cells[2], f"أتعاب الشركة ({pct_label(p)})")
    proto2 = t2.rows[1]
    for r in list(t2.rows[2:]):
        r._tr.getparent().remove(r._tr)

    all_rows = prev + [{"num": num, "amount": amount, "fee": fee, "expenses": total_exp}]
    balance = 0.0
    last_row = proto2
    for i, row in enumerate(all_rows):
        r_fee = round(float(row["fee"]), 2)
        if "net" in row:  # previously published row — reproduce it verbatim
            r_net = round(float(row["net"]), 2)
        else:
            r_net = round(float(row["amount"]) - r_fee - float(row["expenses"]), 2)
        if "balance" in row:
            balance = round(float(row["balance"]), 2)
        else:
            balance = round(balance + r_net, 2)
        target = proto2 if i == 0 else clone_row_after(t2, last_row)
        last_row = target
        set_cell_text(target.cells[0], str(row["num"]))
        set_cell_text(target.cells[1], money(row["amount"], smart_dec(row["amount"])))
        set_cell_text(target.cells[2], money(r_fee, smart_dec(r_fee)))
        set_cell_text(target.cells[3], money(row["expenses"], smart_dec(row["expenses"])))
        set_cell_text(target.cells[4], money(r_net))
        set_cell_text(target.cells[5], money(balance))

    set_para_text(find_para("الرصيد التراكمي المتبقي"),
                  f"الرصيد التراكمي المتبقي حتى تاريخه: {money(balance)}")

    tokens = {"{{COMPANY}}": company["name"], "{{SIGNER}}": company.get("signer", ""),
              "{{SIGNER_TITLE}}": company.get("signer_title", "")}
    for para in doc.paragraphs:
        for r in para.runs:
            for k, v in tokens.items():
                if k in r.text:
                    r.text = r.text.replace(k, v)
    doc.save(args.output)

    # Verification summary for the caller to sanity-check
    print(json.dumps({
        "payment_number": num,
        "payment_amount": amount,
        "fee": fee,
        "available": available,
        "total_expenses": total_exp,
        "net_remaining": net,
        "cumulative_balance": balance,
        "expense_count": len(expenses),
        "output": args.output,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
