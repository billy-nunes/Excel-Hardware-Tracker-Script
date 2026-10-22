# Excel Hardware Tracker Script

A Python script that generates a ready-to-use Excel workbook for tracking IT hardware requests, from the original ticket through approval, ordering, delivery and deployment.

Built for a retail IT support environment (head office + stores, mixed Mac/Windows/iPad estate), but the columns and dropdowns are easy to change.

## What it builds

| Sheet | Purpose |
|---|---|
| **Requests** | The main log: ticket ref, requester, store/dept, item, approval, PO, costs, delivery dates, asset tag, serial, MDM, status |
| **Summary** | Open / overdue / awaiting approval / late delivery counts, average age, total spend, breakdowns by status and category |
| **Lists** | Dropdown options and the overdue threshold (default 14 days) |
| **How to use** | Short guide for whoever fills it in |

### Features
- **Dropdowns** for request type, item category, priority, approval status, MDM and status
- **Auto-calculated** Total Cost (Qty × Unit Cost) and Days Open (today or close date − date raised)
- **Conditional formatting:** overdue requests in red, late deliveries in orange, closed rows greyed out, colour-coded status
- **Date validation** on all date columns
- Formatted as an **Excel table** with filters and frozen headers

## Usage

```bash
pip install -r requirements.txt
python build_tracker.py                      # -> Hardware_Request_Tracker.xlsx
python build_tracker.py -o my_tracker.xlsx   # custom file name
python build_tracker.py --rows 1000          # pre-format more rows
```

Open the output in Excel. Formulas are calculated on open.

## Customising

Everything configurable is at the top of `build_tracker.py`:
- `COLUMNS`: headers and widths
- `DROPDOWNS`: options for each dropdown column
- `STATUS_COLOURS`: status highlight colours
- `DEFAULT_OVERDUE_DAYS`: overdue threshold

## Notes
- `openpyxl` writes formulas but doesn't calculate them, so reading the file back in Python before it has been opened in Excel returns empty formula cells.
- Formulas use widely supported functions (`COUNTIF`, `SUMIFS`, `SUMPRODUCT`) so the workbook works in Excel 2010+ and LibreOffice.

## Ideas / roadmap
- [ ] Import requests from a Freshservice CSV export
- [ ] Pull ticket data via the Freshservice API
- [ ] Monthly spend chart on the Summary sheet
