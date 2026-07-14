from __future__ import annotations

import base64
from datetime import date, datetime, time, timedelta
from io import BytesIO
import json
import re
from typing import Iterable, Sequence

from openai import OpenAI
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


def _clean(value) -> str:
    if value is None:
        return ""

    return re.sub(r"\s+", " ", str(value).strip())


def _image_data_url(uploaded_file) -> str:
    mime = getattr(uploaded_file, "type", None) or "image/jpeg"
    encoded = base64.b64encode(
        uploaded_file.getvalue()
    ).decode("utf-8")

    return f"data:{mime};base64,{encoded}"


def _name_key(name: str) -> str:
    return re.sub(
        r"[^a-z0-9\u4e00-\u9fff]+",
        "",
        name.lower(),
    )


def merge_duplicate_rows(rows: Iterable[dict]) -> list[dict]:
    """
    Merge duplicate employee rows when multiple uploaded photos overlap.
    """

    merged: dict[str, dict] = {}
    order: list[str] = []

    confidence_rank = {
        "low": 0,
        "medium": 1,
        "high": 2,
    }

    for row in rows:
        name = _clean(row.get("Name"))
        key = _name_key(name)

        if not key:
            continue

        candidate = {
            "Name": name,
            "Check In": _clean(row.get("Check In")),
            "Check Out": _clean(row.get("Check Out")),
            "Confidence": (
                _clean(row.get("Confidence")) or "low"
            ),
            "OCR Notes": _clean(row.get("OCR Notes")),
        }

        if key not in merged:
            merged[key] = candidate
            order.append(key)
            continue

        current = merged[key]

        if (
            not current["Check In"]
            and candidate["Check In"]
        ):
            current["Check In"] = candidate["Check In"]

        if (
            not current["Check Out"]
            and candidate["Check Out"]
        ):
            current["Check Out"] = candidate["Check Out"]

        current_confidence = confidence_rank.get(
            current["Confidence"],
            0,
        )

        candidate_confidence = confidence_rank.get(
            candidate["Confidence"],
            0,
        )

        if candidate_confidence > current_confidence:
            current["Confidence"] = candidate["Confidence"]

        notes = [
            current["OCR Notes"],
            candidate["OCR Notes"],
        ]

        current["OCR Notes"] = "; ".join(
            note
            for note in dict.fromkeys(notes)
            if note
        )

    return [merged[key] for key in order]


def extract_attendance_rows_from_images(
    image_files: Sequence,
    company: str,
    scheduled_in: time,
    scheduled_out: time,
    api_key: str,
    model: str = "gpt-5.6",
) -> list[dict]:
    """
    Extract employee names and handwritten check-in/check-out times
    from one or more attendance photos.
    """

    if not image_files:
        return []

    client = OpenAI(api_key=api_key)

    scheduled_in_text = scheduled_in.strftime("%H:%M")
    scheduled_out_text = scheduled_out.strftime("%H:%M")

    prompt = f"""
Read the attached completed attendance-sheet photo(s) for {company}.

The scheduled shift is {scheduled_in_text} to {scheduled_out_text}.
The shift may cross midnight.

Extract every employee row that is not clearly crossed out.

Include printed or handwritten employee names even when check-in or
check-out is blank.

Do not invent missing names or times.

Return all times in 24-hour HH:MM format.

Use the scheduled shift to infer AM versus PM:
- Check-in should normally be near {scheduled_in_text}.
- Check-out should normally be near {scheduled_out_text}.
- For an overnight shift, checkout occurs on the following calendar day.

If a handwritten time is unreadable:
- Return an empty string for that time.
- Explain the uncertainty in OCR Notes.
- Do not silently guess an unreadable digit.

Confidence must be exactly one of:
high, medium, low.
""".strip()

    content = [
        {
            "type": "input_text",
            "text": prompt,
        }
    ]

    for image in image_files:
        content.append(
            {
                "type": "input_image",
                "image_url": _image_data_url(image),
                "detail": "high",
            }
        )

    schema = {
        "type": "object",
        "properties": {
            "rows": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "Name": {
                            "type": "string",
                        },
                        "Check In": {
                            "type": "string",
                        },
                        "Check Out": {
                            "type": "string",
                        },
                        "Confidence": {
                            "type": "string",
                            "enum": [
                                "high",
                                "medium",
                                "low",
                            ],
                        },
                        "OCR Notes": {
                            "type": "string",
                        },
                    },
                    "required": [
                        "Name",
                        "Check In",
                        "Check Out",
                        "Confidence",
                        "OCR Notes",
                    ],
                    "additionalProperties": False,
                },
            }
        },
        "required": ["rows"],
        "additionalProperties": False,
    }

    response = client.responses.create(
        model=model,
        input=[
            {
                "role": "user",
                "content": content,
            }
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "attendance_photo_rows",
                "strict": True,
                "schema": schema,
            }
        },
    )

    payload = json.loads(response.output_text)

    return merge_duplicate_rows(
        payload.get("rows", [])
    )


def _parse_clock(value: str) -> time | None:
    """
    Accept values including:
    19:00
    1900
    7:00 PM
    7 PM
    """

    text = _clean(value)

    if not text:
        return None

    text = (
        text.replace("：", ":")
        .replace("；", ":")
        .replace(".", ":")
        .upper()
    )

    text = re.sub(r"\s+", " ", text).strip()

    formats = (
        "%H:%M",
        "%H%M",
        "%I:%M %p",
        "%I:%M%p",
        "%I %p",
        "%I%p",
    )

    for fmt in formats:
        try:
            return datetime.strptime(
                text,
                fmt,
            ).time()
        except ValueError:
            pass

    match = re.fullmatch(
        r"(\d{1,2})(?::(\d{1,2}))?\s*(AM|PM)?",
        text,
    )

    if not match:
        return None

    hour = int(match.group(1))
    minute = int(match.group(2) or 0)
    am_pm = match.group(3)

    if minute > 59:
        return None

    if am_pm:
        if not 1 <= hour <= 12:
            return None

        if am_pm == "AM":
            hour = 0 if hour == 12 else hour
        else:
            hour = 12 if hour == 12 else hour + 12

    elif not 0 <= hour <= 23:
        return None

    return time(hour, minute)


def _nearest_datetime(
    work_date: date,
    clock: time,
    target: datetime,
) -> datetime:
    """
    Attach an actual clock time to the calendar date closest to the
    scheduled time. This handles shifts that cross midnight.
    """

    candidates = [
        datetime.combine(
            work_date + timedelta(days=offset),
            clock,
        )
        for offset in (-1, 0, 1, 2)
    ]

    return min(
        candidates,
        key=lambda candidate: abs(candidate - target),
    )


def _minutes_text(minutes: int) -> str:
    if minutes == 1:
        return "1 minute"

    return f"{minutes} minutes"


def calculate_attendance_rate(
    company: str,
    records: Sequence[dict],
    work_date: date,
    scheduled_in: time,
    scheduled_out: time,
    expected_headcount: int,
) -> dict:
    """
    Qualified employee requirements:

    Actual check-in <= scheduled check-in
    Actual check-out >= scheduled check-out

    One minute late or one minute early is an exception.
    """

    if expected_headcount <= 0:
        raise ValueError(
            "Expected headcount must be greater than zero."
        )

    scheduled_start = datetime.combine(
        work_date,
        scheduled_in,
    )

    scheduled_end = datetime.combine(
        work_date,
        scheduled_out,
    )

    if scheduled_end <= scheduled_start:
        scheduled_end += timedelta(days=1)

    details: list[dict] = []
    qualified_observed = 0

    for record in merge_duplicate_rows(records):
        name = _clean(record.get("Name"))

        if not name:
            continue

        check_in_text = _clean(
            record.get("Check In")
        )

        check_out_text = _clean(
            record.get("Check Out")
        )

        check_in_clock = _parse_clock(
            check_in_text
        )

        check_out_clock = _parse_clock(
            check_out_text
        )

        actual_in = None
        actual_out = None

        if check_in_clock:
            actual_in = _nearest_datetime(
                work_date,
                check_in_clock,
                scheduled_start,
            )

        if check_out_clock:
            actual_out = _nearest_datetime(
                work_date,
                check_out_clock,
                scheduled_end,
            )

        reasons: list[str] = []

        if actual_in is None:
            reasons.append(
                "Missing or invalid check-in"
            )

        if actual_out is None:
            reasons.append(
                "Missing or invalid check-out"
            )

        if (
            actual_in is not None
            and actual_in > scheduled_start
        ):
            late_minutes = int(
                (
                    actual_in
                    - scheduled_start
                ).total_seconds()
                // 60
            )

            reasons.append(
                f"Late by "
                f"{_minutes_text(late_minutes)}"
            )

        if (
            actual_out is not None
            and actual_out < scheduled_end
        ):
            early_minutes = int(
                (
                    scheduled_end
                    - actual_out
                ).total_seconds()
                // 60
            )

            reasons.append(
                f"Left early by "
                f"{_minutes_text(early_minutes)}"
            )

        qualified = (
            actual_in is not None
            and actual_out is not None
            and actual_in <= scheduled_start
            and actual_out >= scheduled_end
        )

        if qualified:
            qualified_observed += 1

        details.append(
            {
                "Company": company,
                "Employee": name,
                "Scheduled In": (
                    scheduled_in.strftime("%H:%M")
                ),
                "Scheduled Out": (
                    scheduled_out.strftime("%H:%M")
                ),
                "Actual In": check_in_text,
                "Actual Out": check_out_text,
                "Status": (
                    "Qualified"
                    if qualified
                    else "Exception"
                ),
                "Exception": (
                    ""
                    if qualified
                    else "; ".join(reasons)
                ),
                "OCR Confidence": _clean(
                    record.get("Confidence")
                ),
                "OCR Notes": _clean(
                    record.get("OCR Notes")
                ),
            }
        )

    records_found = len(details)

    # Prevent attendance rate from exceeding 100% if extra names
    # appear in the uploaded photos.
    qualified_for_rate = min(
        qualified_observed,
        expected_headcount,
    )

    exceptions = (
        expected_headcount
        - qualified_for_rate
    )

    attendance_rate = (
        qualified_for_rate
        / expected_headcount
        * 100
    )

    return {
        "summary": {
            "Company": company,
            "Scheduled Shift": (
                f"{scheduled_in:%H:%M}"
                f"–"
                f"{scheduled_out:%H:%M}"
            ),
            "Expected": expected_headcount,
            "Records Found": records_found,
            "Qualified": qualified_for_rate,
            "Exceptions": exceptions,
            "Attendance Rate": attendance_rate,
        },
        "details": details,
    }


def build_attendance_rate_workbook(
    results: dict[str, dict],
    work_date: date,
) -> bytes:
    """
    Create an Excel workbook containing:

    Attendance_Summary
    NOVA_Details
    Newstart_Details
    HRN_Details
    """

    wb = Workbook()

    summary_ws = wb.active
    summary_ws.title = "Attendance_Summary"

    blue = "2F75B5"
    light_blue = "BDD7EE"
    light_fill = "F3F6FC"
    exception_fill = "FCE4D6"
    qualified_fill = "E2F0D9"

    thin = Side(
        style="thin",
        color="A6A6A6",
    )

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin,
    )

    summary_ws.merge_cells("A1:G1")

    summary_ws["A1"] = (
        "End-of-Day Attendance Rate — "
        f"{work_date.month}/"
        f"{work_date.day}/"
        f"{work_date.year}"
    )

    summary_ws["A1"].fill = PatternFill(
        "solid",
        fgColor=blue,
    )

    summary_ws["A1"].font = Font(
        name="Arial",
        size=16,
        bold=True,
        color="FFFFFF",
    )

    summary_ws["A1"].alignment = Alignment(
        horizontal="center",
    )

    summary_headers = [
        "Company",
        "Scheduled Shift",
        "Expected",
        "Records Found",
        "Qualified",
        "Exceptions",
        "Attendance Rate",
    ]

    for col, header in enumerate(
        summary_headers,
        1,
    ):
        cell = summary_ws.cell(
            3,
            col,
            header,
        )

        cell.fill = PatternFill(
            "solid",
            fgColor=light_blue,
        )

        cell.font = Font(
            name="Arial",
            bold=True,
        )

        cell.alignment = Alignment(
            horizontal="center",
        )

        cell.border = border

    for row_number, result in enumerate(
        results.values(),
        4,
    ):
        summary = result["summary"]

        values = [
            summary["Company"],
            summary["Scheduled Shift"],
            summary["Expected"],
            summary["Records Found"],
            summary["Qualified"],
            summary["Exceptions"],
            summary["Attendance Rate"] / 100,
        ]

        for col, value in enumerate(
            values,
            1,
        ):
            cell = summary_ws.cell(
                row_number,
                col,
                value,
            )

            cell.fill = PatternFill(
                "solid",
                fgColor=light_fill,
            )

            cell.border = border

            cell.alignment = Alignment(
                horizontal="center",
            )

        summary_ws.cell(
            row_number,
            7,
        ).number_format = "0.00%"

    summary_widths = (
        16,
        20,
        12,
        15,
        12,
        12,
        18,
    )

    for col, width in enumerate(
        summary_widths,
        1,
    ):
        summary_ws.column_dimensions[
            chr(64 + col)
        ].width = width

    summary_ws.freeze_panes = "A4"
    summary_ws.sheet_view.showGridLines = False

    detail_headers = [
        "Company",
        "Employee",
        "Scheduled In",
        "Scheduled Out",
        "Actual In",
        "Actual Out",
        "Status",
        "Exception",
        "OCR Confidence",
        "OCR Notes",
    ]

    for company, result in results.items():
        safe_company = re.sub(
            r"[\[\]:*?/\\]",
            "_",
            company,
        )[:20]

        ws = wb.create_sheet(
            f"{safe_company}_Details"
        )

        for col, header in enumerate(
            detail_headers,
            1,
        ):
            cell = ws.cell(
                1,
                col,
                header,
            )

            cell.fill = PatternFill(
                "solid",
                fgColor=light_blue,
            )

            cell.font = Font(
                name="Arial",
                bold=True,
            )

            cell.alignment = Alignment(
                horizontal="center",
            )

            cell.border = border

        for row_number, detail in enumerate(
            result["details"],
            2,
        ):
            if detail["Status"] == "Qualified":
                row_fill = qualified_fill
            else:
                row_fill = exception_fill

            for col, header in enumerate(
                detail_headers,
                1,
            ):
                cell = ws.cell(
                    row_number,
                    col,
                    detail.get(header, ""),
                )

                cell.fill = PatternFill(
                    "solid",
                    fgColor=row_fill,
                )

                cell.border = border

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )

        detail_widths = (
            14,
            26,
            14,
            14,
            14,
            14,
            14,
            34,
            16,
            34,
        )

        for col, width in enumerate(
            detail_widths,
            1,
        ):
            ws.column_dimensions[
                chr(64 + col)
            ].width = width

        ws.freeze_panes = "A2"
        ws.sheet_view.showGridLines = False

        ws.auto_filter.ref = (
            f"A1:J{max(ws.max_row, 1)}"
        )

    output = BytesIO()
    wb.save(output)

    return output.getvalue()
