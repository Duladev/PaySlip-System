"""
One-click salary-slip mailer over HTTPS (Brevo free plan: 300 emails/day).
Uses normal web traffic (port 443), so it works where SMTP ports are blocked.
No extra packages except openpyxl:   pip install openpyxl

Excel columns:  email | Employee_No | PDF | send_date
  PDF       optional. Blank = auto-find the slip by Employee_No in the PDF folder
            (sub-folder named <Employee_No>, or a PDF whose name contains it).
            Or put a full path / file name here.
  send_date optional. Today or earlier (or blank) = send now; future = skipped.
Results are written back into the Excel file: Status, Sent_At.
"""
import base64
import json
import re
import threading
import urllib.error
import urllib.request
from datetime import datetime, date
from pathlib import Path

import openpyxl

API_BASE = "https://api.brevo.com/v3"
REQUIRED = ["email", "employee_no"]
MAX_PDF_BYTES = 14 * 1024 * 1024          # Brevo allows 20 MB per email incl. base64 growth
SETTINGS = Path(__file__).with_name("send_pdfs_settings.json")


# ----------------------------------------------------------------- helpers
def to_date(v):
    if v in (None, ""):
        return date.today()
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(str(v).strip(), fmt).date()
        except ValueError:
            pass
    raise ValueError(f"bad send_date: {v!r}")


def norm_emp(v):
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip() if v is not None else ""


def find_pdf(base, emp):
    """Find a salary slip for an employee number inside base.
    1) a sub-folder named <emp>  -> newest PDF inside it
    2) a PDF whose file name contains <emp> as a whole token
       (E001.pdf, E001_Nov2026.pdf, Slip-E001.pdf; but not E0010.pdf)"""
    base = Path(base)
    if not base.is_dir():
        return None
    sub = next((d for d in base.iterdir()
                if d.is_dir() and d.name.lower() == emp.lower()), None)
    if sub:
        pdfs = sorted(sub.rglob("*.pdf"), key=lambda f: f.stat().st_mtime, reverse=True)
        if pdfs:
            return pdfs[0]
    pat = re.compile(r"(?<![A-Za-z0-9])" + re.escape(emp) + r"(?![A-Za-z0-9])", re.I)
    hits = [f for f in base.rglob("*.pdf") if pat.search(f.stem)]
    hits.sort(key=lambda f: (f.stem.lower() != emp.lower(), -f.stat().st_mtime))
    return hits[0] if hits else None


# ----------------------------------------------------------------- Brevo API
def api_call(method, path, api_key, payload=None):
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(
        API_BASE + path, data=data, method=method,
        headers={"api-key": api_key, "accept": "application/json",
                 "content-type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            raw = r.read()
            return json.loads(raw) if raw else {}
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "replace")[:300]
        hint = ""
        if e.code == 401:
            hint = (" -> API key is wrong, or your IP is not authorised "
                    "(Brevo > Security > Authorised IPs: deactivate the restriction).")
        raise RuntimeError(f"Brevo error {e.code}: {detail}{hint}")
    except urllib.error.URLError as e:
        raise ConnectionError(f"Cannot reach api.brevo.com over HTTPS: {e.reason}")


def send_email(api_key, sender_email, sender_name, to, subject, body, pdf):
    data = pdf.read_bytes()
    if len(data) > MAX_PDF_BYTES:
        raise ValueError(f"{pdf.name} is too large ({len(data) // 1024 // 1024} MB)")
    sender = {"email": sender_email}
    if sender_name:
        sender["name"] = sender_name[:70]
    payload = {
        "sender": sender,
        "to": [{"email": to}],
        "subject": subject,
        "textContent": body,
        "attachment": [{"name": pdf.name,
                        "content": base64.b64encode(data).decode("ascii")}],
    }
    return api_call("POST", "/smtp/email", api_key, payload).get("messageId", "")


# ----------------------------------------------------------------- main job
def run(xlsx, pdf_dir, api_key, sender_email, sender_name, subject, body,
        log, dry_run=False):
    api_key = api_key.strip()
    sender_email = sender_email.strip()
    if not dry_run:
        if not api_key or not sender_email:
            raise ValueError("Enter the Brevo API key and the sender email.")
        try:                      # fail early if Excel still has the file open
            open(xlsx, "ab").close()
        except PermissionError:
            raise PermissionError("Close the Excel file first, then click SEND again.")
        acct = api_call("GET", "/account", api_key)
        log(f"Connected to Brevo ({acct.get('email', 'ok')})")

    wb = openpyxl.load_workbook(xlsx)
    ws = wb.active
    head = {str(c.value).strip().lower(): c.column for c in ws[1] if c.value}
    missing = [h for h in REQUIRED if h not in head]
    if missing:
        raise ValueError(f"Missing column(s): {', '.join(missing)}")
    for name in ("status", "sent_at"):
        if name not in head:
            head[name] = ws.max_column + 1
            ws.cell(1, head[name], "Status" if name == "status" else "Sent_At")

    def save():
        try:
            wb.save(xlsx)
        except PermissionError:
            log("WARNING: could not save Status to Excel (file is open). "
                "Do NOT re-run until you check which rows were sent.")

    sent = skipped = failed = 0
    for r in range(2, ws.max_row + 1):
        g = lambda k: ws.cell(r, head[k]).value if k in head else None
        email = str(g("email") or "").strip()
        if not email:
            continue
        if str(g("status") or "").upper().startswith("SENT"):
            skipped += 1
            continue
        try:
            if "@" not in email:
                raise ValueError("invalid email address")
            if to_date(g("send_date")) > date.today():
                log(f"Row {r}: {email} scheduled later - skipped")
                skipped += 1
                continue
            emp = norm_emp(g("employee_no"))
            base = Path(pdf_dir) if pdf_dir else Path(xlsx).parent
            raw = str(g("pdf") or "").strip().strip('"').strip("'")
            if raw:
                p = Path(raw)
                if not p.is_absolute():
                    p = base / p
            else:
                p = find_pdf(base, emp)
                if p is None:
                    raise FileNotFoundError(f"No PDF for employee {emp} in {base}")
            if not p.is_file():
                raise FileNotFoundError(f"PDF not found: {p}")
            if dry_run:
                log(f"Row {r}: [DRY RUN] would send {p.name} to {email}")
            else:
                send_email(api_key, sender_email, sender_name, email,
                           subject.replace("{emp}", emp),
                           body.replace("{emp}", emp), p)
                ws.cell(r, head["status"], "SENT")
                ws.cell(r, head["sent_at"], datetime.now().strftime("%Y-%m-%d %H:%M"))
                save()
                log(f"Row {r}: sent {p.name} -> {email}")
            sent += 1
        except Exception as e:
            failed += 1
            if not dry_run:
                ws.cell(r, head["status"], f"FAILED: {e}")
            log(f"Row {r}: FAILED ({email}): {e}")
    if not dry_run:
        save()
    log(f"\nDone. {'Would send' if dry_run else 'Sent'}: {sent}  "
        f"Skipped: {skipped}  Failed: {failed}")


# ----------------------------------------------------------------- GUI
def load_settings():
    try:
        return json.loads(SETTINGS.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_settings(d):
    try:
        SETTINGS.write_text(json.dumps(d, indent=2), encoding="utf-8")
    except Exception:
        pass


def main():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("Send Salary Slips (HTTPS)")
    root.geometry("660x620")
    keys = ("xlsx", "dir", "key", "sender", "name", "subject")
    v = {k: tk.StringVar() for k in keys}
    v["subject"].set("Your salary slip - Employee {emp}")
    saved = load_settings()
    for k in keys:                       # the API key is never saved to disk
        if k != "key" and saved.get(k):
            v[k].set(saved[k])
    dry = tk.BooleanVar(value=True)

    def row(label, key, browse=None, show=None):
        f = tk.Frame(root)
        f.pack(fill="x", padx=10, pady=3)
        tk.Label(f, text=label, width=20, anchor="w").pack(side="left")
        tk.Entry(f, textvariable=v[key], show=show).pack(side="left", fill="x", expand=True)
        if browse:
            tk.Button(f, text="Browse", command=browse).pack(side="left", padx=4)

    row("Excel file", "xlsx", lambda: v["xlsx"].set(
        filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])))
    row("PDF folder (optional)", "dir", lambda: v["dir"].set(filedialog.askdirectory()))
    row("Brevo API key", "key", show="*")
    row("Sender email", "sender")
    row("Sender name (optional)", "name")
    row("Subject", "subject")
    tk.Label(root, text="Email body ({emp} = employee no.)", anchor="w").pack(fill="x", padx=10)
    body = tk.Text(root, height=5)
    body.pack(fill="x", padx=10)
    body.insert("1.0", saved.get("body") or
                "Hello,\n\nPlease find your salary slip attached.\n\nRegards,\nHR")
    tk.Checkbutton(root, text="Dry run (test only, nothing is sent)",
                   variable=dry).pack(anchor="w", padx=10)
    out = scrolledtext.ScrolledText(root, height=12)
    out.pack(fill="both", expand=True, padx=10, pady=4)

    def log(m):
        root.after(0, lambda: (out.insert("end", m + "\n"), out.see("end")))

    def go():
        if not v["xlsx"].get():
            return messagebox.showwarning("Missing", "Choose the Excel file.")
        if not dry.get() and not messagebox.askyesno("Confirm", "Send real emails now?"):
            return
        save_settings({**{k: v[k].get() for k in keys if k != "key"},
                       "body": body.get("1.0", "end").strip()})
        btn.config(state="disabled")

        def work():
            try:
                run(v["xlsx"].get(), v["dir"].get(), v["key"].get(), v["sender"].get(),
                    v["name"].get(), v["subject"].get(), body.get("1.0", "end").strip(),
                    log, dry.get())
            except Exception as e:
                log(f"ERROR: {e}")
            root.after(0, lambda: btn.config(state="normal"))

        threading.Thread(target=work, daemon=True).start()

    btn = tk.Button(root, text="SEND", font=("Segoe UI", 12, "bold"),
                    bg="#2e7d32", fg="white", command=go)
    btn.pack(pady=6, ipadx=30)
    root.mainloop()


if __name__ == "__main__":
    main()