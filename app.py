from __future__ import annotations

from datetime import date
from pathlib import Path
import tempfile

import streamlit as st

from cbt_generator import generate_cbt_file


st.set_page_config(page_title="LAXCBT Attendance Generator", page_icon="📋", layout="centered")

st.title("LAXCBT Attendance Generator")
st.caption("Upload NOVA, Newstart, and HRN spreadsheets, then export one CBT workbook with consistent formatting.")

with st.form("cbt_form"):
    work_date = st.date_input("CBT date", value=date.today())

    nova_file = st.file_uploader("Upload NOVA spreadsheet", type=["xlsx"], key="nova")
    newstart_file = st.file_uploader("Upload Newstart spreadsheet", type=["xlsx"], key="newstart")
    hrn_file = st.file_uploader("Upload HRN spreadsheet", type=["xlsx"], key="hrn")

    submitted = st.form_submit_button("Generate CBT Excel", type="primary")

if submitted:
    if not (nova_file and newstart_file and hrn_file):
        st.error("Please upload all 3 spreadsheets.")
        st.stop()

    mmdd = f"{work_date.month:02d}{work_date.day:02d}"
    out_name = f"cbt_{mmdd}.xlsx"

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        nova_path = tmp / "nova.xlsx"
        newstart_path = tmp / "newstart.xlsx"
        hrn_path = tmp / "hrn.xlsx"
        output_path = tmp / out_name

        nova_path.write_bytes(nova_file.getvalue())
        newstart_path.write_bytes(newstart_file.getvalue())
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
        st.write(
            {
                "NOVA rows": result["counts"]["NOVA"],
                "Newstart rows": result["counts"]["Newstart"],
                "HRN rows": result["counts"]["HRN"],
            }
        )

        st.download_button(
            label=f"Download {out_name}",
            data=output_path.read_bytes(),
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
