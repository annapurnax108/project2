# Running the contractor onboarding helper

This is plain Python. It does not use Claude. The script fills the bulk onboarding
Excel and saves it in the contractor's folder; you do the form in Edge yourself.

> Browser automation is blocked by Deloitte's Edge policy ("DevTools remote
> debugging is disallowed by the system admin"), so the script does not drive the
> ServiceNow form. Only the `--prepare-only` mode is supported on this machine.

## Each terminal session

```powershell
cd C:\Users\sakhadka\projects\my-project
.\.venv\Scripts\Activate.ps1
```

## For each new contractor

1. **Folder:** create the contractor's folder under `Badge_Request` and put their
   `PPMD-Approval.pdf` in it.
2. **CSV:** create `inputs\<name>.csv` (one row per contractor):

   ```
   first_name,last_name,employment_status,rehire_details,email,phone
   Jane,Doe,New Hire,,jane.doe@example.com,703-555-0100
   ```

   - `employment_status` is `New Hire` or `Re-Hire`.
   - A `Re-Hire` needs `rehire_details` (SAP ID / Deloitte email).
   - Email and phone are the contractor's personal ones (mandatory in the template).
   - Several contractors on one request: add one row each.
3. **Fill the Excel** into their folder:

   ```powershell
   python -m my_project.onboard --prepare-only --csv inputs\jane_doe.csv --folder "C:\Users\sakhadka\OneDrive - Deloitte (O365D)\DAF_ IMBSS - FMSS-Nautical Project Management - Documents\Badge_Request\Jane Doe"
   ```

   - Start date defaults to today and end date is start + 364 days.
     Use `--start-date YYYY-MM-DD` to change the start.
   - An existing file is never overwritten; a second run saves `... (1).xlsx`.
4. **Submit the form in Edge** (https://deloitteus.service-now.com/tod?id=sc_cat_item&sys_id=21177d2ddb694550ce9bc170ba96197e):
   - Contractor Location: US. MF employee: No. Multiple contractors: Yes.
   - Upload the filled Excel to **Upload Template**.
   - Upload `PPMD-Approval.pdf` to the PPMD approval field.
   - GPS project: Yes. Type: Federal.
   - Upload `inputs\tprm_approval.pdf` (copy of `TPRM_Approval.pdf` in `Badge_Request`).
   - Read the acknowledgment, select the YES option yourself, and submit.
5. **Screenshot:** save the confirmation (it must show the `HRC...` ticket number)
   into the contractor's folder, e.g. `HRC1234567_confirmation.png`.

## Check before every submission

The fixed template values in `config.yaml` (OGC contact, engaging manager, sponsoring
PPMD, cost center, office, **vendor name / phone / email**) were copied from an earlier
request. Confirm they apply to this contractor, especially the vendor columns (R-T).

## Files that are not in git

`config.yaml`, `inputs\` (blank template, TPRM PDF, CSVs), `output\` and
`.browser_profile\` are git-ignored because they hold real names and approvals.

## Setting up on a new machine

```powershell
git clone https://github.com/annapurnax108/project2.git
cd project2
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -e ".[dev]"
copy config.example.yaml config.yaml    # then fill in the real values
```

Then recreate `inputs\blank_template.xlsx` (the form's template with row 10 cleared)
and `inputs\tprm_approval.pdf`.

## When the template changes

Download the new template from the form's "click here" link, delete the data in
row 10 (keep rows 1-9), and save it as `inputs\blank_template.xlsx`. Keep the
sheet named `Onboarding` and the data row at row 10, or update `first_row` in
`src/my_project/excel_fill.py`.

## Tests

```powershell
python -m pytest
```
