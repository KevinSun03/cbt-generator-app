from __future__ import annotations

from datetime import datetime, time
import os
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from attendance_rate import (
    build_attendance_rate_workbook,
    calculate_attendance_rate,
    extract_attendance_rows_from_images,
)


# ============================================================================
# Configuration
# ============================================================================

LA_TZ = ZoneInfo("America/Los_Angeles")

GROUP_ORDER = (
    "NOVA 1",
    "NOVA 2",
    "Newstart",
    "HRN",
)

DEFAULT_SETTINGS = {
    "NOVA 1": {
        "enabled": True,
        "scheduled_in": time(19, 0),
        "scheduled_out": time(2, 0),
    },
    "NOVA 2": {
        "enabled": False,
        "scheduled_in": time(20, 0),
        "scheduled_out": time(3, 0),
    },
    "Newstart": {
        "enabled": True,
        "scheduled_in": time(16, 30),
        "scheduled_out": time(23, 0),
    },
    "HRN": {
        "enabled": True,
        "scheduled_in": time(17, 0),
        "scheduled_out": time(23, 0),
    },
}

REVIEW_COLUMNS = [
    "Name",
    "Check In",
    "Check Out",
    "Confidence",
    "OCR Notes",
]


# ============================================================================
# Helpers
# ============================================================================

def today_la():
    """Return today's date in the Los Angeles time zone."""
    return datetime.now(LA_TZ).date()


def key_slug(value: str) -> str:
    """Create a safe Streamlit widget-key suffix."""
    return (
        value.lower()
        .replace(" ", "_")
        .replace("-", "_")
    )


def get_secret(
    name: str,
    default: str | None = None,
) -> str | None:
    """
    Read a value from Streamlit Secrets first, then from environment
    variables when running locally.
    """
    try:
        if name in st.secrets:
            return str(st.secrets[name])
    except Exception:
        pass

    return os.getenv(name, default)


def get_openai_api_key() -> str | None:
    return get_secret("OPENAI_API_KEY")


def get_vision_model() -> str:
    return (
        get_secret(
            "OPENAI_VISION_MODEL",
            "gpt-5-mini",
        )
        or "gpt-5-mini"
    )


def initialize_session_state() -> None:
    defaults = {
        "attendance_rows": {},
        "attendance_results": {},
        "attendance_result_date": None,
        "editor_versions": {},
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def reset_group(group_name: str) -> None:
    """Create a blank editable table for a group."""
    st.session_state.attendance_rows[group_name] = []
    st.session_state.attendance_results.pop(
        group_name,
        None,
    )

    current_version = (
        st.session_state.editor_versions.get(
            group_name,
            0,
        )
    )

    st.session_state.editor_versions[group_name] = (
        current_version + 1
    )


def clear_all_attendance_data() -> None:
    st.session_state.attendance_rows = {}
    st.session_state.attendance_results = {}
    st.session_state.attendance_result_date = None
    st.session_state.editor_versions = {}


def render_group_input(
    group_name: str,
    expanded: bool = False,
) -> dict | None:
    """
    Render one attendance group.

    Each group has:
    - Its own scheduled check-in
    - Its own scheduled check-out
    - Its own expected headcount
    - Its own photo uploader
    """
    slug = key_slug(group_name)
    defaults = DEFAULT_SETTINGS[group_name]

    with st.expander(
        group_name,
        expanded=expanded,
    ):
        enabled = st.checkbox(
            f"Calculate {group_name}",
            value=defaults["enabled"],
            key=f"enabled_{slug}",
        )

        if not enabled:
            return None

        schedule_col_1, schedule_col_2, count_col = (
            st.columns(3)
        )

        with schedule_col_1:
            scheduled_in = st.time_input(
                "Scheduled check-in",
                value=defaults["scheduled_in"],
                step=60,
                key=f"scheduled_in_{slug}",
            )

        with schedule_col_2:
            scheduled_out = st.time_input(
                "Scheduled check-out",
                value=defaults["scheduled_out"],
                step=60,
                key=f"scheduled_out_{slug}",
            )

        with count_col:
            expected_headcount = st.number_input(
                "Expected headcount",
                min_value=1,
                max_value=10000,
                value=1,
                step=1,
                key=f"expected_headcount_{slug}",
            )

        photos = st.file_uploader(
            f"Upload completed {group_name} attendance photo(s)",
            type=[
                "png",
                "jpg",
                "jpeg",
                "webp",
            ],
            accept_multiple_files=True,
            key=f"photos_{slug}",
            help=(
                "Upload multiple photos when the attendance sheet "
                "has more than one page. Overlapping employee rows "
                "will be merged by name."
            ),
        )

        if st.button(
            f"Create blank {group_name} review table",
            key=f"blank_table_{slug}",
        ):
            reset_group(group_name)

        return {
            "scheduled_in": scheduled_in,
            "scheduled_out": scheduled_out,
            "expected_headcount": int(
                expected_headcount
            ),
            "photos": photos,
        }


# ============================================================================
# App setup
# ============================================================================

initialize_session_state()

st.set_page_config(
    page_title="End-of-Day Attendance Rate",
    page_icon="📋",
    layout="wide",
)

st.title("End-of-Day Attendance Rate")

st.caption(
    "Upload completed attendance-sheet photos, review the extracted "
    "employee times, and calculate strict attendance rates."
)

st.info(
    "Attendance Rate = employees who arrive on time and complete the "
    "full scheduled shift ÷ expected scheduled employees × 100%."
)

api_key = get_openai_api_key()

if api_key:
    st.success(
        f"OpenAI API is configured. Vision model: {get_vision_model()}."
    )
else:
    st.warning(
        "OpenAI API is not configured. Add OPENAI_API_KEY in "
        "Streamlit Secrets before using photo extraction. You can "
        "still create blank review tables and enter attendance manually."
    )

top_col_1, top_col_2 = st.columns(
    [3, 1]
)

with top_col_1:
    attendance_date = st.date_input(
        "Attendance date",
        value=today_la(),
        key="attendance_date",
    )

with top_col_2:
    st.write("")
    st.write("")

    if st.button(
        "Clear all data",
        key="clear_all_data",
    ):
        clear_all_attendance_data()
        st.rerun()


# ============================================================================
# Daily group settings
# ============================================================================

st.divider()
st.header("Daily Schedule and Photos")

st.caption(
    "Enter the schedule and expected headcount separately for every "
    "active group. NOVA 1 and NOVA 2 are handled independently."
)

group_inputs: dict[str, dict] = {}

st.subheader("NOVA Lines")

nova_1_settings = render_group_input(
    "NOVA 1",
    expanded=True,
)

if nova_1_settings is not None:
    group_inputs["NOVA 1"] = nova_1_settings

nova_2_settings = render_group_input(
    "NOVA 2",
    expanded=False,
)

if nova_2_settings is not None:
    group_inputs["NOVA 2"] = nova_2_settings


st.subheader("Other Companies")

newstart_settings = render_group_input(
    "Newstart",
    expanded=False,
)

if newstart_settings is not None:
    group_inputs["Newstart"] = (
        newstart_settings
    )

hrn_settings = render_group_input(
    "HRN",
    expanded=False,
)

if hrn_settings is not None:
    group_inputs["HRN"] = hrn_settings


# ============================================================================
# Photo extraction
# ============================================================================

st.divider()

extract_clicked = st.button(
    "Extract Attendance from Uploaded Photos",
    type="primary",
    key="extract_attendance",
)

if extract_clicked:
    if not group_inputs:
        st.error(
            "Enable at least one attendance group."
        )

    elif not api_key:
        st.error(
            "OPENAI_API_KEY is not configured. Add it under "
            "Streamlit App settings → Secrets."
        )

    else:
        groups_with_photos = [
            group_name
            for group_name, settings
            in group_inputs.items()
            if settings["photos"]
        ]

        groups_without_photos = [
            group_name
            for group_name, settings
            in group_inputs.items()
            if not settings["photos"]
        ]

        if not groups_with_photos:
            st.error(
                "Upload at least one attendance photo."
            )

        else:
            for group_name in groups_without_photos:
                st.warning(
                    f"{group_name}: no photos were uploaded, "
                    "so this group was skipped."
                )

            successful_extractions = 0

            for group_name in groups_with_photos:
                settings = group_inputs[group_name]

                with st.spinner(
                    f"Reading {group_name} attendance photo(s)..."
                ):
                    try:
                        extracted_rows = (
                            extract_attendance_rows_from_images(
                                image_files=settings["photos"],
                                group_name=group_name,
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
                            f"{group_name}: photo extraction failed."
                        )
                        st.exception(exc)
                        continue

                st.session_state.attendance_rows[
                    group_name
                ] = extracted_rows

                st.session_state.attendance_results.pop(
                    group_name,
                    None,
                )

                current_version = (
                    st.session_state.editor_versions.get(
                        group_name,
                        0,
                    )
                )

                st.session_state.editor_versions[
                    group_name
                ] = current_version + 1

                successful_extractions += 1

                if extracted_rows:
                    st.success(
                        f"{group_name}: extracted "
                        f"{len(extracted_rows)} employee row(s)."
                    )
                else:
                    st.warning(
                        f"{group_name}: no rows were extracted. "
                        "A blank review table is available for "
                        "manual entry."
                    )

            if successful_extractions:
                st.success(
                    "Extraction finished. Review every row below "
                    "before calculating."
                )


# ============================================================================
# Editable review tables
# ============================================================================

edited_frames: dict[str, pd.DataFrame] = {}

groups_ready_for_review = [
    group_name
    for group_name in GROUP_ORDER
    if (
        group_name in group_inputs
        and group_name
        in st.session_state.attendance_rows
    )
]

if groups_ready_for_review:
    st.divider()
    st.header("Review Extracted Attendance")

    st.caption(
        "Correct all names and times. Use 24-hour format such as "
        "19:00 and 02:00, or include AM/PM such as 7:00 PM."
    )

for group_name in groups_ready_for_review:
    rows = st.session_state.attendance_rows.get(
        group_name,
        [],
    )

    slug = key_slug(group_name)

    st.subheader(f"{group_name} Review")

    review_dataframe = pd.DataFrame(
        rows,
        columns=REVIEW_COLUMNS,
    )

    low_confidence_count = 0

    if not review_dataframe.empty:
        low_confidence_count = int(
            review_dataframe["Confidence"]
            .astype(str)
            .str.lower()
            .isin(["low", "medium"])
            .sum()
        )

    if low_confidence_count:
        st.warning(
            f"{group_name}: {low_confidence_count} row(s) have "
            "medium or low OCR confidence."
        )

    editor_version = (
        st.session_state.editor_versions.get(
            group_name,
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
                        "manual",
                    ],
                )
            ),
            "OCR Notes": st.column_config.TextColumn(
                "OCR Notes"
            ),
        },
        key=(
            f"attendance_editor_{slug}_"
            f"{editor_version}"
        ),
    )

    edited_frames[group_name] = (
        edited_dataframe
    )


# ============================================================================
# Attendance calculation
# ============================================================================

if edited_frames:
    st.divider()

    calculate_clicked = st.button(
        "Calculate Attendance Rates",
        type="primary",
        key="calculate_attendance",
    )

    if calculate_clicked:
        calculated_results: dict[str, dict] = {}

        for group_name in GROUP_ORDER:
            if group_name not in edited_frames:
                continue

            if group_name not in group_inputs:
                continue

            settings = group_inputs[group_name]
            edited_dataframe = (
                edited_frames[group_name]
            )

            records = (
                edited_dataframe
                .fillna("")
                .to_dict(orient="records")
            )

            # Remove completely blank rows.
            records = [
                record
                for record in records
                if str(
                    record.get("Name", "")
                ).strip()
            ]

            st.session_state.attendance_rows[
                group_name
            ] = records

            try:
                result = calculate_attendance_rate(
                    group_name=group_name,
                    records=records,
                    work_date=attendance_date,
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

            except Exception as exc:
                st.error(
                    f"{group_name}: attendance rate "
                    "could not be calculated."
                )
                st.exception(exc)
                continue

            calculated_results[group_name] = (
                result
            )

        if calculated_results:
            st.session_state.attendance_results = (
                calculated_results
            )

            st.session_state.attendance_result_date = (
                attendance_date
            )

            st.success(
                "Attendance rates calculated."
            )


# ============================================================================
# Results
# ============================================================================

if st.session_state.attendance_results:
    st.divider()
    st.header("Attendance Rate Results")

    result_date = (
        st.session_state.attendance_result_date
        or attendance_date
    )

    st.caption(
        "Calculated for "
        f"{result_date.month}/"
        f"{result_date.day}/"
        f"{result_date.year}"
    )

    for group_name in GROUP_ORDER:
        result = (
            st.session_state
            .attendance_results
            .get(group_name)
        )

        if not result:
            continue

        summary = result["summary"]

        st.subheader(group_name)

        (
            expected_col,
            found_col,
            qualified_col,
            exception_col,
            rate_col,
        ) = st.columns(5)

        expected_col.metric(
            "Expected",
            summary["Expected"],
        )

        found_col.metric(
            "Records found",
            summary["Records Found"],
        )

        qualified_col.metric(
            "Qualified",
            summary["Qualified"],
        )

        exception_col.metric(
            "Exceptions",
            summary["Scheduled Exceptions"],
        )

        rate_col.metric(
            "Attendance rate",
            f'{summary["Attendance Rate"]:.2f}%',
        )

        st.caption(
            f'Scheduled shift: {summary["Scheduled Shift"]}'
        )

        if summary["Missing Records"] > 0:
            st.warning(
                f'{group_name}: '
                f'{summary["Missing Records"]} expected employee(s) '
                "had no attendance record."
            )

        if summary["Extra Records"] > 0:
            st.warning(
                f'{group_name}: '
                f'{summary["Extra Records"]} more attendance '
                "record(s) were found than the expected headcount. "
                "Verify the expected headcount or remove unscheduled rows."
            )

        details_dataframe = pd.DataFrame(
            result["details"]
        )

        if details_dataframe.empty:
            st.info(
                f"No detail rows are available for {group_name}."
            )

        else:
            show_exceptions_only = st.checkbox(
                "Show exceptions only",
                value=False,
                key=(
                    f"show_exceptions_"
                    f"{key_slug(group_name)}"
                ),
            )

            displayed_dataframe = (
                details_dataframe
            )

            if show_exceptions_only:
                displayed_dataframe = (
                    displayed_dataframe[
                        displayed_dataframe["Status"]
                        == "Exception"
                    ]
                )

            st.dataframe(
                displayed_dataframe,
                use_container_width=True,
                hide_index=True,
            )

    attendance_workbook_bytes = (
        build_attendance_rate_workbook(
            st.session_state.attendance_results,
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
        key="download_attendance_results",
    )
