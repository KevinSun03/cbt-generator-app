# Daily CBT Generator

This app turns 3 attendance spreadsheets into one daily CBT workbook.

## What it does

Upload:

1. NOVA spreadsheet
2. Newstart spreadsheet
3. HRN spreadsheet

Then it creates one Excel file with three formatted sheets:

- `NOVA_MMDD`
- `Newstart_MMDD`
- `HRN_MMDD`

The output follows the CBT style: title row, date/company row, header row, alternating blue body rows, and columns for 编号、姓名、劳务公司、上班时间、下班时间、备注.

Newstart automatically uses **TT only** by default.

## How to run

Open Terminal in this folder, then run:

```bash
python3 -m pip install -r requirements.txt
streamlit run app.py
```

Your browser will open the app.

## Daily workflow

1. Pick the CBT date.
2. Upload the 3 spreadsheets.
3. Click **Generate CBT Excel**.
4. Download `cbt_MMDD.xlsx`.

## Notes

- This version is for spreadsheets, not handwritten photos.
- If HRN is a weekly SWX sheet, the app uses the selected date if that date exists in the sheet. If not, it uses the latest non-empty IN/OUT pair for each employee.
- Newstart TT notes include built-in defaults for names like 林虹、孙亦可、祁奕帆、王竹龙、杨先忠、彭宇、李少龙.
