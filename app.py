from __future__ import annotations

from datetime import date
from pathlib import Path
import tempfile

import streamlit as st

from cbt_generator import generate_cbt_file


st.set_page_config(page_title="Daily CBT Generator", page_icon="📋", layout="centered")

st.title("Daily CBT Generator")
st.caption(
    "Upload the company spreadsheets you have, then export one CBT workbook. "
    "Newstart uses TT only; HRN uses the selected date's weekday column."
)

with st.form("cbt_form"):
    work_date = st.date_input("CBT date", value=date.today())

    st.subheader("Companies to include")
    st.caption("Uncheck a company when that company did not work or you do not have that spreadsheet.")

    include_nova = st.checkbox("Include NOVA", value=True)
    nova_file = None
    if include_nova:
        nova_file = st.file_uploader("Upload NOVA spreadsheet", type=["xlsx"], key="nova")

    include_newstart = st.checkbox("Include Newstart", value=True)
    newstart_file = None
    if include_newstart:
        newstart_file = st.file_uploader("Upload Newstart spreadsheet", type=["xlsx"], key="newstart")

    include_hrn = st.checkbox("Include HRN", value=True)
    hrn_file = None
    if include_hrn:
        hrn_file = st.file_uploader("Upload HRN spreadsheet", type=["xlsx"], key="hrn")

    submitted = st.form_submit_button("Generate CBT Excel", type="primary")

if submitted:
    selected_files = {
        "NOVA": nova_file if include_nova else None,
        "Newstart": newstart_file if include_newstart else None,
        "HRN": hrn_file if include_hrn else None,
    }

    selected_companies = [company for company, file in selected_files.items() if file is not None]
    missing_selected = [
        company
        for company, file in selected_files.items()
        if ((company == "NOVA" and include_nova) or (company == "Newstart" and include_newstart) or (company == "HRN" and include_hrn))
        and file is None
    ]

    if not (include_nova or include_newstart or include_hrn):
        st.error("Please select at least one company to include.")
        st.stop()

    if missing_selected:
        st.error("Please upload the selected spreadsheet(s): " + ", ".join(missing_selected))
        st.stop()

    if not selected_companies:
        st.error("Please upload at least one spreadsheet.")
        st.stop()

    mmdd = f"{work_date.month:02d}{work_date.day:02d}"
    out_name = f"cbt_{mmdd}.xlsx"

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        output_path = tmp / out_name

        nova_path = None
        newstart_path = None
        hrn_path = None

        if nova_file is not None:
            nova_path = tmp / "nova.xlsx"
            nova_path.write_bytes(nova_file.getvalue())

        if newstart_file is not None:
            newstart_path = tmp / "newstart.xlsx"
            newstart_path.write_bytes(newstart_file.getvalue())

        if hrn_file is not None:
            hrn_path = tmp / "hrn.xlsx"
            hrn_path.write_bytes(hrn_file.getvalue())

        try:
            result = generate_cbt_file(
                nova_path=nova_path,
                newstart_path=newstart_path,
                hrn_path=hrn_path,
                output_path=output_path,
                work_date=work_date,
                newstart_tt_only=True,
            )
        except Exception as e:
            st.exception(e)
            st.stop()

        st.success("CBT workbook generated.")
        st.write({f"{company} rows": count for company, count in result["counts"].items()})

        st.download_button(
            label=f"Download {out_name}",
            data=output_path.read_bytes(),
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
