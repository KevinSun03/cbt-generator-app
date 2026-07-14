from __future__ import annotations

from datetime import datetime, time
import os
from pathlib import Path
import tempfile
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from attendance_rate import (
    build_attendance_rate_workbook,
    calculate_attendance_rate,
    extract_attendance_rows_from_images,
)
from cbt_generator import generate_cbt_file


# ---------------------------------------------------------------------------
# General configuration
# ---------------------------------------------------------------------------

LA_TZ = ZoneInfo("America/Los_Angeles")

COMPANIES = (
    "NOVA",
    "Newstart",
    "HRN",
)

REVIEW_COLUMNS = [
    "Name",
    "Check In",
    "Check Out",
    "Confidence",
    "OCR Notes",
]


def today_la():
    """Return today's date in the Los Angeles time zone."""
    return datetime.now(LA_TZ).date()


def get_secret(name: str, default: str | None = None) -> str | None:
    """Read a setting from Streamlit Secrets, then environment variables."""
    try:
        if name in st.secrets:
            return str(st.secrets[name])
    except Exception:
        pass

    return os.getenv(name, default)


def get_openai_api_key() -> str | None:
    return get_secret("OPENAI_API_KEY")


def get_vision_model() -> str:
    return get_secret(
        "OPENAI_VISION_MODEL",
        "gpt-4.1-mini",
    ) or "gpt-4.1-mini"


def initialize_session_state() -> None:
    """Create session-state containers used by both app sections."""
    defaults = {
        "cbt_output_bytes": None,
        "cbt_output_name": None,
        "cbt_generation_result": None,
        "photo_attendance_rows": {},
        "photo_attendance_results": {},
        "photo_attendance_result_date": None,
        "photo_editor_versions": {},
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


initialize_session_state()


st.set_page_config(
    page_title="Daily CBT Generator",
    page_icon="📋",
    layout="centered",
)

st.title("Daily CBT Generator")

st.caption(
    "Generate the daily CBT workbook from company spreadsheets, "
    "or calculate end-of-day attendance rates from completed attendance photos."
)


spreadsheet_tab, photo_tab = st.tabs(
    [
        "Daily CBT Excel",
        "End-of-Day Photo Attendance",
    ]
)


# ===========================================================================
# TAB 1 — DAILY CBT EXCEL
# ===========================================================================

with spreadsheet_tab:
    st.header("Daily CBT Excel")

    st.caption(
        "Upload the company spreadsheets you have and export one CBT workbook. "
        "Newstart uses TT only. HRN uses Second Shift only and selects attendance "
        "using the chosen date's weekday column."
    )

    with st.form("cbt_form"):
        work_date = st.date_input(
            "CBT date",
            value=today_la(),
            key="cbt_work_date",
        )

        st.subheader("Companies to include")

        st.caption(
            "Uncheck a company when that company did not work "
            "or you do not have that spreadsheet."
        )

        include_nova = st.checkbox(
            "Include NOVA",
            value=True,
            key="include_nova",
        )

        nova_file = None

        if include_nova:
            nova_file = st.file_uploader(
                "Upload NOVA spreadsheet",
                type=["xlsx"],
                key="nova_spreadsheet",
            )

        include_newstart = st.checkbox(
            "Include Newstart",
            value=True,
            key="include_newstart",
        )

        newstart_file = None

        if include_newstart:
            newstart_file = st.file_uploader(
                "Upload Newstart spreadsheet",
                type=["xlsx"],
                key="newstart_spreadsheet",
            )

        include_hrn = st.checkbox(
            "Include HRN",
            value=True,
            key="include_hrn",
        )

        hrn_file = None

        if include_hrn:
            hrn_file = st.file_uploader(
                "Upload HRN spreadsheet",
                type=["xlsx"],
                key="hrn_spreadsheet",
            )

        submitted = st.form_submit_button(
            "Generate CBT Excel",
            type="primary",
        )

    if submitted:
        selected_files = {
            "NOVA": {
                "enabled": include_nova,
                "file": nova_file,
            },
            "Newstart": {
                "enabled": include_newstart,
                "file": newstart_file,
            },
            "HRN": {
                "enabled": include_hrn,
                "file": hrn_file,
            },
        }

        enabled_companies = [
            company
            for company, settings in selected_files.items()
            if settings["enabled"]
        ]

        missing_files = [
            company
            for company, settings in selected_files.items()
            if settings["enabled"] and settings["file"] is None
        ]

        if not enabled_companies:
            st.error(
                "Please select at least one company to include."
            )

        elif missing_files:
            st.error(
                "Please upload the selected spreadsheet(s): "
                + ", ".join(missing_files)
            )

        else:
            mmdd = (
                f"{work_date.month:02d}"
                f"{work_date.day:02d}"
            )

            output_name = f"cbt_{mmdd}.xlsx"

            with tempfile.TemporaryDirectory() as temp_directory:
                temp_path = Path(temp_directory)
                output_path = temp_path / output_name

                nova_path = None
                newstart_path = None
                hrn_path = None

                if nova_file is not None:
                    nova_path = temp_path / "nova.xlsx"
                    nova_path.write_bytes(
                        nova_file.getvalue()
                    )

                if newstart_file is not None:
                    newstart_path = (
                        temp_path / "newstart.xlsx"
                    )
                    newstart_path.write_bytes(
                        newstart_file.getvalue()
                    )

                if hrn_file is not None:
                    hrn_path = temp_path / "hrn.xlsx"
                    hrn_path.write_bytes(
                        hrn_file.getvalue()
                    )

                try:
                    generation_result = generate_cbt_file(
                        nova_path=nova_path,
                        newstart_path=newstart_path,
                        hrn_path=hrn_path,
                        output_path=output_path,
                        work_date=work_date,
                        newstart_tt_only=True,
                    )

                    st.session_state.cbt_output_bytes = (
                        output_path.read_bytes()
                    )

                    st.session_state.cbt_output_name = (
                        output_name
                    )

                    st.session_state.cbt_generation_result = (
                        generation_result
                    )

                except Exception as exc:
                    st.session_state.cbt_output_bytes = None
                    st.session_state.cbt_output_name = None
                    st.session_state.cbt_generation_result = None

                    st.error(
                        "The CBT workbook could not be generated."
                    )
                    st.exception(exc)

    if (
        st.session_state.cbt_output_bytes is not None
        and st.session_state.cbt_generation_result is not None
    ):
        result = st.session_state.cbt_generation_result

        st.success("CBT workbook generated.")

        count_columns = st.columns(
            max(len(result["counts"]), 1)
        )

        for column, (company, count) in zip(
            count_columns,
            result["counts"].items(),
        ):
            column.metric(
                f"{company} rows",
                count,
            )

        st.caption(
            "Created sheets: "
            + ", ".join(result.get("sheets", []))
        )

        st.download_button(
            label=(
                f"Download "
                f"{st.session_state.cbt_output_name}"
            ),
            data=st.session_state.cbt_output_bytes,
            file_name=st.session_state.cbt_output_name,
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="download_cbt_workbook",
        )


# ===========================================================================
# TAB 2 — END-OF-DAY PHOTO ATTENDANCE
# ===========================================================================

with photo_tab:
    st.header("End-of-Day Photo Attendance")

    st.caption(
        "Upload completed attendance-sheet photos, review the extracted names "
        "and times, and calculate the strict attendance rate. An employee must "
        "arrive on time and complete the entire scheduled shift. One minute late "
        "or one minute early is an attendance exception."
    )

    photo_date = st.date_input(
        "Attendance date",
        value=today_la(),
        key="photo_attendance_date",
    )

    st.info(
        "Attendance Rate = Qualified employees ÷ Expected scheduled employees × 100%"
    )

    company_inputs: dict[str, dict] = {}

    default_schedules = {
        "NOVA": {
            "in": time(19, 0),
            "out": time(2, 0),
        },
        "Newstart": {
            "in": time(16, 30),
            "out": time(23, 0),
        },
        "HRN": {
            "in": time(17, 0),
            "out": time(23, 0),
        },
    }

    for company in COMPANIES:
        with st.expander(
            company,
            expanded=(company == "NOVA"),
        ):
            enabled = st.checkbox(
                f"Calculate {company}",
                value=True,
                key=f"photo_enabled_{company}",
            )

            if not enabled:
                continue

            schedule_column_1, schedule_column_2 = (
                st.columns(2)
            )

            with schedule_column_1:
                scheduled_in = st.time_input(
                    "Scheduled check-in",
                    value=default_schedules[company]["in"],
                    step=60,
                    key=f"photo_scheduled_in_{company}",
                )

            with schedule_column_2:
                scheduled_out = st.time_input(
                    "Scheduled check-out",
                    value=default_schedules[company]["out"],
                    step=60,
                    key=f"photo_scheduled_out_{company}",
                )

            expected_headcount = st.number_input(
                "Expected scheduled employee headcount",
                min_value=1,
                max_value=10000,
                value=1,
                step=1,
                key=f"photo_expected_{company}",
            )

            photos = st.file_uploader(
                f"Upload completed {company} attendance photo(s)",
                type=[
                    "png",
                    "jpg",
                    "jpeg",
                    "webp",
                ],
                accept_multiple_files=True,
                key=f"photo_uploads_{company}",
            )

            company_inputs[company] = {
                "scheduled_in": scheduled_in,
                "scheduled_out": scheduled_out,
                "expected_headcount": int(
                    expected_headcount
                ),
                "photos": photos,
            }

    # Remove saved rows/results when a company is currently disabled.
    enabled_photo_companies = set(company_inputs)

    for company in COMPANIES:
        if company not in enabled_photo_companies:
            st.session_state.photo_attendance_rows.pop(
                company,
                None,
            )

            st.session_state.photo_attendance_results.pop(
                company,
                None,
            )

    extract_clicked = st.button(
        "Extract Attendance from Photos",
        type="primary",
        key="extract_photo_attendance",
    )

    if extract_clicked:
        api_key = get_openai_api_key()

        if not api_key:
            st.error(
                "OPENAI_API_KEY is not configured. Add it in "
                "Streamlit Community Cloud under App settings → Secrets."
            )

        else:
            companies_with_photos = [
                company
                for company, settings in company_inputs.items()
                if settings["photos"]
            ]

            if not companies_with_photos:
                st.error(
                    "Upload at least one attendance photo."
                )

            else:
                successful_extractions = 0

                for company in companies_with_photos:
                    settings = company_inputs[company]

                    with st.spinner(
                        f"Reading {company} attendance photo(s)..."
                    ):
                        try:
                            extracted_rows = (
                                extract_attendance_rows_from_images(
                                    image_files=settings["photos"],
                                    company=company,
                                    scheduled_in=settings[
                                        "scheduled_in"
                                    ],
                                    scheduled_out=settings[
                                        "scheduled_out"
                                    ],
                                    api_key=api_key,
                                    model=get_vision_model(),
                                )
                            )

                        except Exception as exc:
                            st.error(
                                f"{company} photo extraction failed."
                            )
                            st.exception(exc)
                            continue

                    st.session_state.photo_attendance_rows[
                        company
                    ] = extracted_rows

                    st.session_state.photo_attendance_results.pop(
                        company,
                        None,
                    )

                    current_version = (
                        st.session_state.photo_editor_versions.get(
                            company,
                            0,
                        )
                    )

                    st.session_state.photo_editor_versions[
                        company
                    ] = current_version + 1

                    successful_extractions += 1

                    if extracted_rows:
                        st.success(
                            f"{company}: extracted "
                            f"{len(extracted_rows)} attendance rows."
                        )
                    else:
                        st.warning(
                            f"{company}: no rows were confidently extracted. "
                            "You can add the rows manually in the review table."
                        )

                if successful_extractions:
                    st.success(
                        "Photo extraction finished. Review and correct "
                        "all names and times before calculating."
                    )

    edited_frames: dict[str, pd.DataFrame] = {}

    for company, settings in company_inputs.items():
        rows = st.session_state.photo_attendance_rows.get(
            company
        )

        if rows is None:
            continue

        st.divider()
        st.subheader(f"{company} Review")

        st.caption(
            "Correct any misread names or handwritten times. "
            "Use 24-hour time such as 19:00 and 02:00, "
            "or AM/PM time such as 7:00 PM and 2:00 AM."
        )

        review_dataframe = pd.DataFrame(
            rows,
            columns=REVIEW_COLUMNS,
        )

        editor_version = (
            st.session_state.photo_editor_versions.get(
                company,
                0,
            )
        )

        edited_dataframe = st.data_editor(
            review_dataframe,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Name": st.column_config.TextColumn(
                    "Employee Name",
                    required=True,
                ),
                "Check In": st.column_config.TextColumn(
                    "Check In",
                    help=(
                        "Examples: 19:00, 7:00 PM, 18:59."
                    ),
                ),
                "Check Out": st.column_config.TextColumn(
                    "Check Out",
                    help=(
                        "Examples: 02:00, 2:00 AM, 23:00."
                    ),
                ),
                "Confidence": (
                    st.column_config.SelectboxColumn(
                        "OCR Confidence",
                        options=[
                            "high",
                            "medium",
                            "low",
                        ],
                    )
                ),
                "OCR Notes": st.column_config.TextColumn(
                    "OCR Notes"
                ),
            },
            key=(
                f"photo_editor_{company}_"
                f"{editor_version}"
            ),
        )

        edited_frames[company] = edited_dataframe

    calculate_clicked = False

    if edited_frames:
        st.divider()

        calculate_clicked = st.button(
            "Calculate Attendance Rate",
            type="primary",
            key="calculate_photo_attendance",
        )

    if calculate_clicked:
        calculated_results: dict[str, dict] = {}

        for company, edited_dataframe in (
            edited_frames.items()
        ):
            settings = company_inputs[company]

            records = (
                edited_dataframe
                .fillna("")
                .to_dict(orient="records")
            )

            # Save corrections so they persist in session state.
            st.session_state.photo_attendance_rows[
                company
            ] = records

            try:
                calculated_results[company] = (
                    calculate_attendance_rate(
                        company=company,
                        records=records,
                        work_date=photo_date,
                        scheduled_in=settings[
                            "scheduled_in"
                        ],
                        scheduled_out=settings[
                            "scheduled_out"
                        ],
                        expected_headcount=settings[
                            "expected_headcount"
                        ],
                    )
                )

            except Exception as exc:
                st.error(
                    f"{company} attendance rate could not be calculated."
                )
                st.exception(exc)

        if calculated_results:
            st.session_state.photo_attendance_results = (
                calculated_results
            )

            st.session_state.photo_attendance_result_date = (
                photo_date
            )

            st.success(
                "Attendance rates calculated."
            )

    if st.session_state.photo_attendance_results:
        st.divider()
        st.header("Attendance Rate Results")

        result_date = (
            st.session_state.photo_attendance_result_date
            or photo_date
        )

        st.caption(
            "Calculated for "
            f"{result_date.month}/"
            f"{result_date.day}/"
            f"{result_date.year}"
        )

        for company in COMPANIES:
            result = (
                st.session_state
                .photo_attendance_results
                .get(company)
            )

            if not result:
                continue

            summary = result["summary"]

            st.subheader(company)

            (
                metric_expected,
                metric_found,
                metric_qualified,
                metric_exceptions,
                metric_rate,
            ) = st.columns(5)

            metric_expected.metric(
                "Expected",
                summary["Expected"],
            )

            metric_found.metric(
                "Records found",
                summary["Records Found"],
            )

            metric_qualified.metric(
                "Qualified",
                summary["Qualified"],
            )

            metric_exceptions.metric(
                "Exceptions",
                summary["Exceptions"],
            )

            metric_rate.metric(
                "Attendance rate",
                f'{summary["Attendance Rate"]:.2f}%',
            )

            if (
                summary["Records Found"]
                < summary["Expected"]
            ):
                missing_count = (
                    summary["Expected"]
                    - summary["Records Found"]
                )

                st.warning(
                    f"{company}: {missing_count} expected employee(s) "
                    "are missing from the uploaded attendance records."
                )

            elif (
                summary["Records Found"]
                > summary["Expected"]
            ):
                extra_count = (
                    summary["Records Found"]
                    - summary["Expected"]
                )

                st.warning(
                    f"{company}: {extra_count} more attendance record(s) "
                    "were found than the expected headcount."
                )

            details_dataframe = pd.DataFrame(
                result["details"]
            )

            if details_dataframe.empty:
                st.info(
                    f"No attendance detail rows are available for {company}."
                )
            else:
                st.dataframe(
                    details_dataframe,
                    use_container_width=True,
                    hide_index=True,
                )

        attendance_workbook_bytes = (
            build_attendance_rate_workbook(
                st.session_state.photo_attendance_results,
                result_date,
            )
        )

        result_mmdd = (
            f"{result_date.month:02d}"
            f"{result_date.day:02d}"
        )

        result_filename = (
            f"attendance_rate_{result_mmdd}.xlsx"
        )

        st.download_button(
            label=f"Download {result_filename}",
            data=attendance_workbook_bytes,
            file_name=result_filename,
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="download_photo_attendance",
        )
