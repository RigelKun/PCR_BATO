from datetime import datetime
import csv
import io
import json
import os
import sqlite3
from pathlib import Path

from flask import Flask, flash, g, redirect, render_template, request, url_for, Response, send_file


BASE_DIR = Path(__file__).resolve().parent
DATABASE = Path(os.environ.get("PCR_DB_PATH", str(BASE_DIR / "pcr.db"))).resolve()
BACKUP_DIR = BASE_DIR / "backups"

app = Flask(__name__)
app.config["SECRET_KEY"] = "change-this-secret-key"


def get_db() -> sqlite3.Connection:
    if "db" not in g:
        g.db = sqlite3.connect(DATABASE)
        g.db.row_factory = sqlite3.Row
        # Durability-oriented defaults for local SQLite usage.
        g.db.execute("PRAGMA foreign_keys = ON")
        g.db.execute("PRAGMA journal_mode = WAL")
        g.db.execute("PRAGMA synchronous = FULL")
    return g.db


def init_db() -> None:
    DATABASE.parent.mkdir(parents=True, exist_ok=True)
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(DATABASE)
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS pcr_reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            patient_name TEXT NOT NULL,
            nature_of_call TEXT,
            call_date TEXT,
            time_of_call TEXT,
            status TEXT NOT NULL DEFAULT 'SUBMITTED',
            form_data TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
        """
    )
    # Backward-compatible migration from the earlier simple schema.
    cols = {row[1] for row in db.execute("PRAGMA table_info(pcr_reports)").fetchall()}
    if "status" not in cols:
        db.execute("ALTER TABLE pcr_reports ADD COLUMN status TEXT NOT NULL DEFAULT 'SUBMITTED'")
    if "form_data" not in cols:
        db.execute("ALTER TABLE pcr_reports ADD COLUMN form_data TEXT NOT NULL DEFAULT '{}' ")
    db.commit()
    db.close()


def _create_db_backup_file() -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = BACKUP_DIR / f"pcr_backup_{timestamp}.db"

    src = get_db()
    dest = sqlite3.connect(backup_path)
    try:
        src.backup(dest)
    finally:
        dest.close()

    return backup_path


# Ensure schema exists whether app is started via `python app.py` or `flask run`.
init_db()


@app.teardown_appcontext
def close_db(_error: BaseException | None) -> None:
    db = g.pop("db", None)
    if db is not None:
        db.close()


@app.route("/")
def home():
    return redirect(url_for("records"))


def _collect_multi(field: str) -> list[str]:
    return [v for v in request.form.getlist(field) if v]


def _collect_kin_entries() -> list[dict[str, str]]:
    """Collect dynamically added next of kin entries."""
    kins: list[dict[str, str]] = []
    indexes = sorted(
        {
            int(key.rsplit("_", 1)[1])
            for key in request.form.keys()
            if key.startswith("kin_name_") and key.rsplit("_", 1)[1].isdigit()
        }
        | {
            int(key.rsplit("_", 1)[1])
            for key in request.form.keys()
            if key.startswith("kin_contact_") and key.rsplit("_", 1)[1].isdigit()
        }
    )
    for idx in indexes:
        name = request.form.get(f"kin_name_{idx}", "").strip()
        contact = request.form.get(f"kin_contact_{idx}", "").strip()
        if name or contact:
            kins.append({"name": name, "contact": contact})
    return kins


def _collect_vital_rows() -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for idx in range(1, 6):
        row = {
            "time": request.form.get(f"vital_time_{idx}", "").strip(),
            "loc": request.form.get(f"vital_loc_{idx}", "").strip(),
            "bp": request.form.get(f"vital_bp_{idx}", "").strip(),
            "rr": request.form.get(f"vital_rr_{idx}", "").strip(),
            "pr": request.form.get(f"vital_pr_{idx}", "").strip(),
            "temp": request.form.get(f"vital_temp_{idx}", "").strip(),
            "spo2": request.form.get(f"vital_spo2_{idx}", "").strip(),
            "rbs": request.form.get(f"vital_rbs_{idx}", "").strip(),
            "pain_scale": request.form.get(f"vital_pain_scale_{idx}", "").strip(),
            "gcs_eye": request.form.get(f"vital_gcs_eye_{idx}", "").strip(),
            "gcs_verbal": request.form.get(f"vital_gcs_verbal_{idx}", "").strip(),
            "gcs_motor": request.form.get(f"vital_gcs_motor_{idx}", "").strip(),
            "gcs_total": request.form.get(f"vital_gcs_total_{idx}", "").strip(),
        }
        if any(row.values()):
            rows.append(row)
    return rows


def _split_csv_values(value: str) -> list[str]:
    return [item.strip() for item in value.split(",") if item.strip()]


def _load_form_data(row: sqlite3.Row) -> dict:
    try:
        return json.loads(row["form_data"] or "{}")
    except json.JSONDecodeError:
        return {}


def _build_prefill(form_data: dict | None = None, row: sqlite3.Row | None = None) -> dict:
    form_data = form_data or {}

    sample_history = form_data.get("sample_history", {})
    patient_assessment = form_data.get("patient_assessment", {})
    physical_exam = form_data.get("physical_exam", {})
    gcs = form_data.get("gcs", {})
    obstetrics = form_data.get("obstetrics_data", {})
    contact_dispatch = form_data.get("contact_dispatch", {})
    incident_details = form_data.get("incident_details", {})
    team_destination = form_data.get("team_destination", {})
    mvc = form_data.get("mvc", {})

    if row is None:
        patient_name = ""
        nature_of_call = ""
        call_date = ""
        time_of_call = ""
        status = "SUBMITTED"
    else:
        patient_name = row["patient_name"] or ""
        nature_of_call = row["nature_of_call"] or ""
        call_date = row["call_date"] or ""
        time_of_call = row["time_of_call"] or ""
        status = row["status"] or "SUBMITTED"

    return {
        "simple": {
            "patient_name": patient_name,
            "age": form_data.get("patient_info", {}).get("age", ""),
            "gender": form_data.get("patient_info", {}).get("gender", ""),
            "nationality": form_data.get("patient_info", {}).get("nationality", ""),
            "call_date": call_date,
            "time_of_call": time_of_call,
            "status": status,
            "permanent_address": contact_dispatch.get("permanent_address", ""),
            "contact_number": contact_dispatch.get("contact_number", ""),
            "etd_base": contact_dispatch.get("etd_base", ""),
            "eta_scene": contact_dispatch.get("eta_scene", ""),
            "etd_scene": incident_details.get("etd_scene", ""),
            "eta_hospital": incident_details.get("eta_hospital", ""),
            "etd_hospital": incident_details.get("etd_hospital", ""),
            "eta_base": incident_details.get("eta_base", ""),
            "location_of_incident": incident_details.get("location_of_incident", ""),
            "nature_of_illness": incident_details.get("nature_of_illness", ""),
            "mechanism_of_injury": incident_details.get("mechanism_of_injury", ""),
            "chief_complaint": patient_assessment.get("chief_complaint", ""),
            "body_diagram_drawing": physical_exam.get("body_diagram_drawing", ""),
            "deformity": physical_exam.get("deformity", ""),
            "bleeding": physical_exam.get("bleeding", ""),
            "contusion": physical_exam.get("contusion", ""),
            "tenderness": physical_exam.get("tenderness", ""),
            "abrasion": physical_exam.get("abrasion", ""),
            "laceration": physical_exam.get("laceration", ""),
            "punctured": physical_exam.get("punctured", ""),
            "swelling": physical_exam.get("swelling", ""),
            "pupils": patient_assessment.get("pupils", ""),
            "eye_opening": gcs.get("eye_opening", ""),
            "verbal_response": gcs.get("verbal_response", ""),
            "motor_response": gcs.get("motor_response", ""),
            "gcs_total": gcs.get("gcs_total", ""),
            "last_menstrual_period": obstetrics.get("last_menstrual_period", ""),
            "gravida": obstetrics.get("gravida", ""),
            "pre_term": obstetrics.get("pre_term", ""),
            "age_of_gestation": obstetrics.get("age_of_gestation", ""),
            "para": obstetrics.get("para", ""),
            "abortion": obstetrics.get("abortion", ""),
            "expected_date_of_delivery": obstetrics.get("expected_date_of_delivery", ""),
            "term": obstetrics.get("term", ""),
            "living": obstetrics.get("living", ""),
            "onset": form_data.get("opqrst", {}).get("onset", ""),
            "provoking_factors": form_data.get("opqrst", {}).get("provoking_factors", ""),
            "quality": form_data.get("opqrst", {}).get("quality", ""),
            "radiation": form_data.get("opqrst", {}).get("radiation", ""),
            "severity": form_data.get("opqrst", {}).get("severity", ""),
            "timing": form_data.get("opqrst", {}).get("timing", ""),
            "narrative": form_data.get("narrative_report", ""),
            "ambulance_driver": team_destination.get("ambulance_driver", ""),
            "plate_no": team_destination.get("plate_no", ""),
            "license_no": team_destination.get("license_no", ""),
            "responders_tl": team_destination.get("responders_tl", ""),
            "crew": team_destination.get("crew", ""),
            "destination_determination": team_destination.get("destination_determination", ""),
            "receiving_facility": team_destination.get("receiving_facility", ""),
            "receiving_personnel": team_destination.get("receiving_personnel", ""),
            "mvc_type_of_vehicles": mvc.get("type_of_vehicles", ""),
            "mvc_plate_numbers": mvc.get("plate_numbers", ""),
            "mvc_type_of_accident": mvc.get("type_of_accident", ""),
            "law_enforcer_name": mvc.get("law_enforcer_name", ""),
            "law_enforcer_contact": mvc.get("law_enforcer_contact", ""),
        },
        "radio": {
            "nature_of_call": nature_of_call,
            "c_spine": patient_assessment.get("c_spine", ""),
            "airway": patient_assessment.get("airway", ""),
            "loc_level": patient_assessment.get("loc_level", ""),
            "capillary_refill": patient_assessment.get("capillary_refill", ""),
        },
        "multi": {
            "types_of_emergencies": contact_dispatch.get("types_of_emergencies", []),
            "breathing": patient_assessment.get("breathing", []),
            "circulation": patient_assessment.get("circulation", []),
            "skin": form_data.get("physical_exam", {}).get("skin", []),
            "burns": form_data.get("physical_exam", {}).get("burns", []),
            "findings": form_data.get("physical_exam", {}).get("findings", []),
            "care_airway": form_data.get("care_management", {}).get("airway", []),
            "care_breathing": form_data.get("care_management", {}).get("breathing", []),
            "care_circulation": form_data.get("care_management", {}).get("circulation", []),
            "care_immobilization": form_data.get("care_management", {}).get("immobilization", []),
            "care_wound": form_data.get("care_management", {}).get("wound_care", []),
        },
        "symptoms": _split_csv_values(sample_history.get("symptoms", "")),
        "next_of_kin": form_data.get("next_of_kin", {}).get("kins", []),
        "vital_signs": form_data.get("vital_signs", []),
    }


def _flatten_for_csv(value, prefix: str, output: dict[str, str]) -> None:
    if isinstance(value, dict):
        for key, nested_value in value.items():
            next_prefix = f"{prefix}_{key}" if prefix else key
            _flatten_for_csv(nested_value, next_prefix, output)
        return

    if isinstance(value, list):
        if not value:
            output[prefix] = "N/A"
            return

        if all(isinstance(item, dict) for item in value):
            for index, item in enumerate(value, start=1):
                _flatten_for_csv(item, f"{prefix}_{index}", output)
        else:
            output[prefix] = " | ".join(str(item) for item in value)
        return

    text = "" if value is None else str(value).strip()
    output[prefix] = text if text else "N/A"


def _collect_form_data() -> dict:
    return {
        "patient_info": {
            "age": request.form.get("age", "").strip(),
            "gender": request.form.get("gender", "").strip(),
            "nationality": request.form.get("nationality", "").strip(),
            "date": request.form.get("call_date", "").strip(),
            "time_of_call": request.form.get("time_of_call", "").strip(),
        },
        "contact_dispatch": {
            "permanent_address": request.form.get("permanent_address", "").strip(),
            "contact_number": request.form.get("contact_number", "").strip(),
            "etd_base": request.form.get("etd_base", "").strip(),
            "types_of_emergencies": _collect_multi("types_of_emergencies"),
            "eta_scene": request.form.get("eta_scene", "").strip(),
        },
        "next_of_kin": {
            "kins": _collect_kin_entries(),
        },
        "incident_details": {
            "location_of_incident": request.form.get("location_of_incident", "").strip(),
            "nature_of_illness": request.form.get("nature_of_illness", "").strip(),
            "mechanism_of_injury": request.form.get("mechanism_of_injury", "").strip(),
            "etd_scene": request.form.get("etd_scene", "").strip(),
            "eta_hospital": request.form.get("eta_hospital", "").strip(),
            "etd_hospital": request.form.get("etd_hospital", "").strip(),
            "eta_base": request.form.get("eta_base", "").strip(),
        },
        "obstetrics_data": {
            "last_menstrual_period": request.form.get("last_menstrual_period", "").strip(),
            "gravida": request.form.get("gravida", "").strip(),
            "pre_term": request.form.get("pre_term", "").strip(),
            "age_of_gestation": request.form.get("age_of_gestation", "").strip(),
            "para": request.form.get("para", "").strip(),
            "abortion": request.form.get("abortion", "").strip(),
            "expected_date_of_delivery": request.form.get("expected_date_of_delivery", "").strip(),
            "term": request.form.get("term", "").strip(),
            "living": request.form.get("living", "").strip(),
        },
        "patient_assessment": {
            "chief_complaint": request.form.get("chief_complaint", "").strip(),
            "c_spine": request.form.get("c_spine", "").strip(),
            "airway": request.form.get("airway", "").strip(),
            "breathing": _collect_multi("breathing"),
            "circulation": _collect_multi("circulation"),
            "pupils": request.form.get("pupils", "").strip(),
            "loc_level": request.form.get("loc_level", "").strip(),
            "capillary_refill": request.form.get("capillary_refill", "").strip(),
        },
        "gcs": {
            "eye_opening": request.form.get("eye_opening", "").strip(),
            "verbal_response": request.form.get("verbal_response", "").strip(),
            "motor_response": request.form.get("motor_response", "").strip(),
            "gcs_total": request.form.get("gcs_total", "").strip(),
        },
        "vital_signs": _collect_vital_rows(),
        "sample_history": {
            "symptoms": request.form.get("symptoms", "").strip(),
            "allergies": request.form.get("allergies", "").strip(),
            "medications": request.form.get("medications", "").strip(),
            "past_medical_history": request.form.get("past_medical_history", "").strip(),
            "last_oral_intake": request.form.get("last_oral_intake", "").strip(),
            "events_prior": request.form.get("events_prior", "").strip(),
        },
        "opqrst": {
            "onset": request.form.get("onset", "").strip(),
            "provoking_factors": request.form.get("provoking_factors", "").strip(),
            "quality": request.form.get("quality", "").strip(),
            "radiation": request.form.get("radiation", "").strip(),
            "severity": request.form.get("severity", "").strip(),
            "timing": request.form.get("timing", "").strip(),
        },
        "physical_exam": {
            "skin": _collect_multi("skin"),
            "burns": _collect_multi("burns"),
            "findings": _collect_multi("findings"),
            "body_diagram_drawing": request.form.get("body_diagram_drawing", "").strip(),
            "deformity": request.form.get("deformity", "").strip(),
            "bleeding": request.form.get("bleeding", "").strip(),
            "contusion": request.form.get("contusion", "").strip(),
            "tenderness": request.form.get("tenderness", "").strip(),
            "abrasion": request.form.get("abrasion", "").strip(),
            "laceration": request.form.get("laceration", "").strip(),
            "punctured": request.form.get("punctured", "").strip(),
            "swelling": request.form.get("swelling", "").strip(),
        },
        "narrative_report": request.form.get("narrative", "").strip(),
        "team_destination": {
            "ambulance_driver": request.form.get("ambulance_driver", "").strip(),
            "plate_no": request.form.get("plate_no", "").strip(),
            "license_no": request.form.get("license_no", "").strip(),
            "responders_tl": request.form.get("responders_tl", "").strip(),
            "crew": request.form.get("crew", "").strip(),
            "destination_determination": request.form.get("destination_determination", "").strip(),
            "receiving_facility": request.form.get("receiving_facility", "").strip(),
            "receiving_personnel": request.form.get("receiving_personnel", "").strip(),
        },
        "care_management": {
            "airway": _collect_multi("care_airway"),
            "breathing": _collect_multi("care_breathing"),
            "circulation": _collect_multi("care_circulation"),
            "immobilization": _collect_multi("care_immobilization"),
            "wound_care": _collect_multi("care_wound"),
        },
        "mvc": {
            "type_of_vehicles": request.form.get("mvc_type_of_vehicles", "").strip(),
            "plate_numbers": request.form.get("mvc_plate_numbers", "").strip(),
            "type_of_accident": request.form.get("mvc_type_of_accident", "").strip(),
            "law_enforcer_name": request.form.get("law_enforcer_name", "").strip(),
            "law_enforcer_contact": request.form.get("law_enforcer_contact", "").strip(),
        },
    }


@app.route("/new", methods=["GET", "POST"])
def new_pcr():
    if request.method == "POST":
        patient_name = request.form.get("patient_name", "").strip()
        nature_of_call = request.form.get("nature_of_call", "").strip()
        call_date = request.form.get("call_date", "").strip()
        time_of_call = request.form.get("time_of_call", "").strip()
        status = request.form.get("status", "SUBMITTED").strip() or "SUBMITTED"

        if not patient_name:
            flash("Patient Name is required.", "error")
            return render_template("new.html")

        form_data = _collect_form_data()

        db = get_db()
        try:
            db.execute(
                """
                INSERT INTO pcr_reports (
                    patient_name, nature_of_call, call_date, time_of_call,
                    status, form_data, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    patient_name,
                    nature_of_call,
                    call_date,
                    time_of_call,
                    status,
                    json.dumps(form_data, ensure_ascii=True),
                    datetime.now().isoformat(timespec="seconds"),
                ),
            )
            db.commit()
        except sqlite3.Error:
            db.rollback()
            flash("Failed to save PCR report due to a database error.", "error")
            return render_template(
                "new.html",
                form_action=url_for("new_pcr"),
                page_title="PATIENT CARE REPORT",
                submit_label="Save PCR",
                prefill=_build_prefill(form_data),
            )

        flash("PCR report saved.", "success")
        return redirect(url_for("records"))

    return render_template(
        "new.html",
        form_action=url_for("new_pcr"),
        page_title="PATIENT CARE REPORT",
        submit_label="Save PCR",
        prefill=_build_prefill(),
    )


@app.route("/records")
def records():
    search_name = request.args.get("q", "").strip()
    search_date = request.args.get("date", "").strip()

    db = get_db()
    query = "SELECT * FROM pcr_reports WHERE 1=1"
    params: list[str] = []
    if search_name:
        query += " AND patient_name LIKE ?"
        params.append(f"%{search_name}%")
    if search_date:
        query += " AND call_date = ?"
        params.append(search_date)
    query += " ORDER BY id DESC"

    rows = db.execute(query, params).fetchall()
    return render_template(
        "records.html",
        records=rows,
        search_name=search_name,
        search_date=search_date,
    )


@app.route("/records/<int:report_id>")
def view_record(report_id: int):
    db = get_db()
    row = db.execute(
        "SELECT * FROM pcr_reports WHERE id = ?",
        (report_id,),
    ).fetchone()
    if row is None:
        flash("PCR record not found.", "error")
        return redirect(url_for("records"))

    form_data = {}
    form_data = _load_form_data(row)

    return render_template("view.html", record=row, form_data=form_data)


@app.route("/records/<int:report_id>/edit", methods=["GET", "POST"])
def edit_record(report_id: int):
    db = get_db()
    row = db.execute(
        "SELECT * FROM pcr_reports WHERE id = ?",
        (report_id,),
    ).fetchone()
    if row is None:
        flash("PCR record not found.", "error")
        return redirect(url_for("records"))

    form_data = _load_form_data(row)
    prefill = _build_prefill(form_data, row)

    if request.method == "POST":
        patient_name = request.form.get("patient_name", "").strip()
        nature_of_call = request.form.get("nature_of_call", "").strip()
        call_date = request.form.get("call_date", "").strip()
        time_of_call = request.form.get("time_of_call", "").strip()
        status = request.form.get("status", "SUBMITTED").strip() or "SUBMITTED"

        if not patient_name:
            flash("Patient Name is required.", "error")
        else:
            form_data = _collect_form_data()

            try:
                db.execute(
                    """
                    UPDATE pcr_reports
                    SET patient_name = ?, nature_of_call = ?, call_date = ?, time_of_call = ?,
                        status = ?, form_data = ?
                    WHERE id = ?
                    """,
                    (
                        patient_name,
                        nature_of_call,
                        call_date,
                        time_of_call,
                        status,
                        json.dumps(form_data, ensure_ascii=True),
                        report_id,
                    ),
                )
                db.commit()
            except sqlite3.Error:
                db.rollback()
                flash("Failed to update PCR report due to a database error.", "error")
                prefill = _build_prefill(form_data, row)
                return render_template(
                    "new.html",
                    record=row,
                    form_data=form_data,
                    prefill=prefill,
                    form_action=url_for("edit_record", report_id=report_id),
                    page_title=f"EDIT PCR RECORD #{report_id}",
                    submit_label="Update PCR",
                )

            flash("PCR record updated.", "success")
            return redirect(url_for("view_record", report_id=report_id))

    return render_template(
        "new.html",
        record=row,
        form_data=form_data,
        prefill=prefill,
        form_action=url_for("edit_record", report_id=report_id),
        page_title=f"EDIT PCR RECORD #{report_id}",
        submit_label="Update PCR",
    )


@app.route("/records/<int:report_id>/delete", methods=["POST"])
def delete_record(report_id: int):
    db = get_db()
    try:
        db.execute("DELETE FROM pcr_reports WHERE id = ?", (report_id,))
        remaining = db.execute("SELECT COUNT(*) AS total FROM pcr_reports").fetchone()["total"]
        if remaining == 0:
            # Reset autoincrement so a fresh dataset starts again at ID 1.
            db.execute("DELETE FROM sqlite_sequence WHERE name = 'pcr_reports'")
        db.commit()
    except sqlite3.Error:
        db.rollback()
        flash("Failed to delete PCR record due to a database error.", "error")
        return redirect(url_for("records"))
    flash("PCR record deleted.", "success")
    return redirect(url_for("records"))


@app.route("/records/delete-all", methods=["POST"])
def delete_all_records():
    db = get_db()
    try:
        db.execute("DELETE FROM pcr_reports")
        # Reset autoincrement so next record starts at ID 1.
        db.execute("DELETE FROM sqlite_sequence WHERE name = 'pcr_reports'")
        db.commit()
    except sqlite3.Error:
        db.rollback()
        flash("Failed to delete all records due to a database error.", "error")
        return redirect(url_for("records"))

    flash("All PCR records have been deleted.", "success")
    return redirect(url_for("records"))


@app.route("/records/export.csv")
def export_records_csv():
    db = get_db()
    rows = db.execute("SELECT * FROM pcr_reports ORDER BY id DESC").fetchall()

    export_rows: list[dict[str, str]] = []
    fieldnames: list[str] = [
        "id",
        "patient_name",
        "nature_of_call",
        "call_date",
        "time_of_call",
        "status",
        "created_at",
        "form_data_json",
    ]
    extra_fields: set[str] = set()

    for row in rows:
        row_data: dict[str, str] = {
            "id": str(row["id"]),
            "patient_name": row["patient_name"] or "",
            "nature_of_call": row["nature_of_call"] or "",
            "call_date": row["call_date"] or "",
            "time_of_call": row["time_of_call"] or "",
            "status": row["status"] or "",
            "created_at": row["created_at"] or "",
            "form_data_json": row["form_data"] or "{}",
        }

        form_data = _load_form_data(row)
        flattened: dict[str, str] = {}
        _flatten_for_csv(form_data, "form_data", flattened)
        row_data.update(flattened)
        extra_fields.update(flattened.keys())
        export_rows.append(row_data)

    fieldnames.extend(sorted(extra_fields))

    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=fieldnames)
    writer.writeheader()
    for row_data in export_rows:
        writer.writerow(row_data)

    csv_data = output.getvalue()
    output.close()

    return Response(
        csv_data,
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=pcr_records.csv"},
    )


@app.route("/records/export.db")
def export_records_db():
    backup_path = _create_db_backup_file()
    return send_file(
        backup_path,
        as_attachment=True,
        download_name=backup_path.name,
        mimetype="application/octet-stream",
    )


if __name__ == "__main__":
    init_db()
    app.run(debug=True)
