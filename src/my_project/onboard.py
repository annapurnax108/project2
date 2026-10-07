"""Submit a Contractor Access & Maintenance (Onboarding) request in ServiceNow.

Usage:
    python -m my_project.onboard --csv inputs/contractors.csv            # dry run (default)
    python -m my_project.onboard --csv inputs/contractors.csv --submit   # really submit

Dry run fills the whole form and stops before the submit button.
"""
import argparse
import csv
import datetime as dt
import re
import sys
from pathlib import Path

import yaml

from my_project.excel_fill import fill_template

ROOT = Path(__file__).resolve().parents[2]

# Template columns that change per contractor.
PERSON_COLUMNS = {
    "first_name": "Z",
    "last_name": "AA",
    "employment_status": "AB",
    "rehire_details": "AC",
    "email": "AF",
    "phone": "AG",
}
START_COL, END_COL = "C", "D"


def load_config(path: Path) -> dict:
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def read_people(args) -> list[dict]:
    if args.csv:
        with open(args.csv, newline="", encoding="utf-8-sig") as f:
            people = list(csv.DictReader(f))
    else:
        people = [{k: getattr(args, k) or "" for k in PERSON_COLUMNS}]
    for i, p in enumerate(people, 1):
        missing = [k for k in ("first_name", "last_name", "employment_status", "email", "phone") if not p.get(k)]
        if missing:
            sys.exit(f"Contractor {i}: missing {', '.join(missing)}")
        if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", p["email"]):
            sys.exit(f"Contractor {i}: invalid email {p['email']!r}")
        if len(re.sub(r"\D", "", p["phone"])) < 10:
            sys.exit(f"Contractor {i}: phone needs at least 10 digits: {p['phone']!r}")
        if p["employment_status"] not in ("New Hire", "Re-Hire"):
            sys.exit(f"Contractor {i}: employment_status must be 'New Hire' or 'Re-Hire'")
        if p["employment_status"] == "Re-Hire" and not p.get("rehire_details"):
            sys.exit(f"Contractor {i}: Re-Hire needs rehire_details (SAP ID / email)")
    return people


def build_records(people: list[dict], fixed: dict, start: dt.date) -> list[dict]:
    end = start + dt.timedelta(days=364)
    records = []
    for p in people:
        record = dict(fixed)
        record[START_COL], record[END_COL] = start, end
        for key, col in PERSON_COLUMNS.items():
            record[col] = p.get(key, "")
        records.append(record)
    return records


# --- browser helpers -------------------------------------------------------
# Locators are label-based and have NOT been verified against the live form.
# Run with the default dry run and watch the browser; fix any step that fails.

def pick(page, label: str, value: str) -> None:
    """Choose a dropdown option by field label (native <select> or portal widget)."""
    field = page.get_by_label(label, exact=False).first
    try:
        field.select_option(label=value, timeout=3000)
    except Exception:
        field.click()
        page.get_by_role("option", name=re.compile(re.escape(value), re.I)).first.click()


def upload(page, label: str, path: Path) -> None:
    button = page.get_by_text(label, exact=False).first.locator(
        "xpath=following::button[normalize-space()='Upload'][1]"
    )
    with page.expect_file_chooser() as chooser:
        button.click()
    chooser.value.set_files(str(path))


def run_browser(cfg: dict, records: list[dict], ppmd: Path, tprm: Path, out_dir: Path, submit: bool) -> None:
    from playwright.sync_api import sync_playwright

    out_dir.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as pw:
        # Persistent profile: sign in through SSO once; later runs reuse the session.
        ctx = pw.chromium.launch_persistent_context(
            str(ROOT / ".browser_profile"), channel="msedge", headless=False, accept_downloads=True
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(cfg["url"])
        input("Sign in if prompted. When the request form is showing, press Enter... ")

        pick(page, "Contractor Location", cfg["contractor_location"])
        # Type of Request and Requestor are locked/prefilled on this form; verify, don't set.
        if cfg["request_type"].lower() not in page.inner_text("body").lower():
            sys.exit(f"Form does not show request type {cfg['request_type']!r}; wrong catalog item?")
        pick(page, "Memberfirm (MF) employee", cfg["memberfirm_employee"])

        print("\nACKNOWLEDGMENT: you affirm that TPRM clearance, a valid background check, an executed")
        print("contractual agreement and (if required) an NDA are complete for these contractors.")
        if input("Type 'yes' to select the YES acknowledgment: ").strip().lower() != "yes":
            sys.exit("Acknowledgment not confirmed; nothing submitted.")
        pick(page, "Acknowledgement", cfg["acknowledgement_prefix"])
        pick(page, "multiple contractors", "Yes")

        with page.expect_download() as dl:
            page.get_by_role("link", name=re.compile("click here", re.I)).first.click()
        template = out_dir / "template_downloaded.xlsx"
        dl.value.save_as(str(template))
        filled = out_dir / dl.value.suggested_filename
        fill_template(template, filled, records)
        print(f"Filled template: {filled}")

        upload(page, "Upload Template", filled)
        upload(page, "upload the PPMD approval", ppmd)
        pick(page, "Is this a GPS project", cfg["gps_project"])
        pick(page, "What type of GPS project", cfg["gps_type"])
        upload(page, "upload the TPRM approval", tprm)

        shot = out_dir / "form_before_submit.png"
        page.screenshot(path=str(shot), full_page=True)
        print(f"Screenshot: {shot}")

        if not submit:
            input("DRY RUN - nothing submitted. Review the browser, then press Enter to close... ")
        elif input("Type 'submit' to submit this ticket: ").strip().lower() == "submit":
            page.get_by_role("button", name=re.compile(r"submit|order now", re.I)).first.click()
            banner = page.get_by_text(re.compile(r"HRC\d+ created")).first
            banner.wait_for(timeout=90000)
            ticket = re.search(r"HRC\d+", banner.inner_text()).group()
            page.screenshot(path=str(out_dir / "confirmation.png"), full_page=True)
            (out_dir / "ticket.txt").write_text(ticket + "\n")
            print(f"Submitted: {ticket}. Confirmation screenshot saved.")
        else:
            print("Not submitted.")
        ctx.close()


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", default=str(ROOT / "config.yaml"))
    ap.add_argument("--csv", help="CSV with columns: " + ",".join(PERSON_COLUMNS))
    for key in PERSON_COLUMNS:
        ap.add_argument(f"--{key.replace('_', '-')}", dest=key)
    ap.add_argument("--ppmd", required=True, type=Path, help="PPMD approval file for this request")
    ap.add_argument("--start-date", type=dt.date.fromisoformat, default=dt.date.today(), help="YYYY-MM-DD (default today)")
    ap.add_argument("--out", type=Path, default=ROOT / "output" / dt.datetime.now().strftime("%Y%m%d-%H%M%S"))
    ap.add_argument("--submit", action="store_true", help="actually submit (default is dry run)")
    args = ap.parse_args()

    cfg = load_config(Path(args.config))
    tprm = ROOT / cfg["tprm_pdf"]
    for label, p in (("PPMD approval", args.ppmd), ("TPRM approval", tprm)):
        if not Path(p).is_file():
            sys.exit(f"{label} not found: {p}")
    people = read_people(args)
    records = build_records(people, cfg["fixed_columns"], args.start_date)
    print(f"{len(people)} contractor(s); start {args.start_date}, end {args.start_date + dt.timedelta(days=364)}")
    run_browser(cfg, records, args.ppmd, tprm, args.out, args.submit)


if __name__ == "__main__":
    main()
