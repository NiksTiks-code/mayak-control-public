from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from flask import send_file
from datetime import datetime
from io import BytesIO
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer
)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.colors import HexColor
from reportlab.platypus import Image
from reportlab.lib.styles import ParagraphStyle
from flask import Flask, render_template, request, redirect, session, abort, url_for, jsonify
from werkzeug.security import check_password_hash
from werkzeug.utils import secure_filename
import json
import os
import re
import uuid
import fcntl
from contextlib import contextmanager
from excel_importer import (
    apply_import_to_data,
    normalize_chessboard,
    parse_excel_chessboard,
)

app = Flask(__name__)
secret_key = os.environ.get("MAYAK_SECRET_KEY")
if not secret_key:
    raise RuntimeError("MAYAK_SECRET_KEY is required")
app.secret_key = secret_key
del secret_key


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "data.json")
UPLOAD_FOLDER = os.path.join(BASE_DIR, "static", "uploads")
USERS_FILE = os.path.join(BASE_DIR, "users.json")
CONTRACTORS_FILE = os.path.join(BASE_DIR, "contractors.json")
IMPORT_PREVIEW_FOLDER = os.path.join(BASE_DIR, "instance", "import_previews")
DATA_LOCK_FILE = f"{DATA_FILE}.lock"
CONTRACTORS_LOCK_FILE = f"{CONTRACTORS_FILE}.lock"
USERS_LOCK_FILE = f"{USERS_FILE}.lock"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(IMPORT_PREVIEW_FOLDER, exist_ok=True)

objects = {
    "Demo Street 15": {
        "systems": ["ВО", "ХВС", "ТС"],
        "flats": 12,
        "floors": 3
    }
}

DEFAULT_CONTRACTORS = {
    "Demo Contractor": [
        "Demo Street 15"
    ]
}


def normalize_contractor_name(value):
    """Return a stable display name without changing the user's letter case."""
    return " ".join(str(value or "").split())


def contractor_lookup_key(value):
    return normalize_contractor_name(value).casefold()


@contextmanager
def interprocess_lock(lock_file):
    with open(lock_file, "a+") as handle:
        fcntl.flock(handle.fileno(), fcntl.LOCK_EX)
        try:
            yield
        finally:
            fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


def atomic_write_json(path, payload, mode=None):
    temporary_file = f"{path}.{os.getpid()}.{uuid.uuid4().hex}.tmp"
    try:
        with open(temporary_file, "w", encoding="utf-8") as file:
            json.dump(payload, file, ensure_ascii=False, indent=4)
            file.flush()
            os.fsync(file.fileno())
        if mode is not None:
            os.chmod(temporary_file, mode)
        os.replace(temporary_file, path)
    finally:
        if os.path.exists(temporary_file):
            os.remove(temporary_file)


def normalize_contractors(contractor_data):
    normalized = {}
    for raw_name, raw_addresses in contractor_data.items():
        name = normalize_contractor_name(raw_name)
        if not name or contractor_lookup_key(name) in {
            contractor_lookup_key(item) for item in normalized
        }:
            continue
        addresses = []
        for raw_address in raw_addresses if isinstance(raw_addresses, list) else []:
            address = str(raw_address or "").strip()
            if address and address not in addresses:
                addresses.append(address)
        normalized[name] = addresses

    return normalized


def save_contractors(contractor_data):
    with interprocess_lock(CONTRACTORS_LOCK_FILE):
        atomic_write_json(CONTRACTORS_FILE, normalize_contractors(contractor_data), 0o640)


def update_contractors(mutator):
    global contractors
    with interprocess_lock(CONTRACTORS_LOCK_FILE):
        latest = load_contractors()
        result = mutator(latest)
        latest = normalize_contractors(latest)
        atomic_write_json(CONTRACTORS_FILE, latest, 0o640)
        contractors.clear()
        contractors.update(latest)
        return result


def load_contractors():
    if not os.path.exists(CONTRACTORS_FILE):
        initial = {
            name: list(addresses)
            for name, addresses in DEFAULT_CONTRACTORS.items()
        }
        save_contractors(initial)
        return initial

    try:
        with open(CONTRACTORS_FILE, "r", encoding="utf-8") as file:
            stored = json.load(file)
    except (OSError, json.JSONDecodeError):
        stored = {}

    if not isinstance(stored, dict):
        stored = {}

    normalized = {}
    for raw_name, raw_addresses in stored.items():
        name = normalize_contractor_name(raw_name)
        if not name:
            continue
        if contractor_lookup_key(name) in {
            contractor_lookup_key(item) for item in normalized
        }:
            continue
        normalized[name] = list(dict.fromkeys(
            str(address).strip()
            for address in raw_addresses if str(address).strip()
        )) if isinstance(raw_addresses, list) else []

    for name, addresses in DEFAULT_CONTRACTORS.items():
        existing = next(
            (item for item in normalized if contractor_lookup_key(item) == contractor_lookup_key(name)),
            None,
        )
        if existing is None:
            normalized[name] = list(addresses)
        else:
            for address in addresses:
                if address not in normalized[existing]:
                    normalized[existing].append(address)

    return normalized


contractors = load_contractors()


def refresh_contractors():
    contractors.clear()
    contractors.update(load_contractors())

IMPORT_SYSTEMS = ("ВО", "ХВС", "ГВС", "ТС")
IMPORT_EXTENSIONS = {".xlsx", ".xlsm"}


def _preview_path(token):
    if not re.fullmatch(r"[a-f0-9]{32}", token or ""):
        abort(404)
    return os.path.join(IMPORT_PREVIEW_FOLDER, f"{token}.json")


def save_import_preview(token, payload):
    preview_path = _preview_path(token)
    temporary_file = f"{preview_path}.{os.getpid()}.tmp"

    try:
        with open(temporary_file, "w", encoding="utf-8") as file:
            json.dump(payload, file, ensure_ascii=False, indent=2)
            file.flush()
            os.fsync(file.fileno())
        os.replace(temporary_file, preview_path)
    finally:
        if os.path.exists(temporary_file):
            os.remove(temporary_file)


def load_import_preview(token):
    preview_path = _preview_path(token)
    if not os.path.exists(preview_path):
        abort(404)
    with open(preview_path, "r", encoding="utf-8") as file:
        payload = json.load(file)
    if not isinstance(payload, dict) or "chessboard" not in payload:
        abort(400)
    return payload


def delete_import_preview(token):
    preview_path = _preview_path(token)
    if os.path.exists(preview_path):
        os.remove(preview_path)


def prepare_chessboard_for_display(chessboard, system_data, contractor, address, system):
    board = normalize_chessboard(chessboard)

    for cell in board.get("cells", []):
        if cell.get("cell_type") not in {"apartment", "non_residential"}:
            continue

        room_key = str(cell.get("room_key") or cell.get("room_number") or "").strip()
        saved = system_data.get(room_key, {}) if room_key else {}
        cell["display_status"] = saved.get("status") or "white"

        if cell.get("cell_type") == "apartment" and room_key.isdigit():
            cell["href"] = url_for(
                "flat",
                flat_num=int(room_key),
                contractor=contractor,
                object=address,
                system=system,
            )
        else:
            cell["href"] = url_for(
                "imported_room",
                room_key=room_key,
                contractor=contractor,
                object=address,
                system=system,
            )

    return board


def infer_imported_object_size(chessboard):
    flat_rooms = [
        room for room in chessboard.get("rooms", [])
        if room.get("type") == "flat"
    ]
    flat_numbers = [
        int(room["number"])
        for room in flat_rooms
        if str(room.get("number", "")).isdigit()
    ]

    flats = max(flat_numbers, default=max(len(flat_rooms), 1))
    detected_floors = max(
        len({room.get("row") for room in flat_rooms}),
        1
    )
    floors = min(detected_floors, flats)

    while floors > 1 and flats % floors != 0:
        floors -= 1

    return flats, floors


def sync_imported_objects(data):
    imported_objects = data.get("imported_objects", {})

    if not isinstance(imported_objects, dict):
        return

    contractors_changed = False

    for address, imported_object in imported_objects.items():
        if not address or not isinstance(imported_object, dict):
            continue

        imported_systems = [
            system for system in imported_object.get("systems", [])
            if system in IMPORT_SYSTEMS
        ]
        existing_object = objects.get(address)

        if existing_object is None:
            objects[address] = {
                "systems": list(imported_systems),
                "flats": max(int(imported_object.get("flats", 1)), 1),
                "floors": max(int(imported_object.get("floors", 1)), 1)
            }
        else:
            for system in imported_systems:
                if system not in existing_object["systems"]:
                    existing_object["systems"].append(system)
            existing_object["flats"] = max(int(imported_object.get("flats", 1)), 1)
            existing_object["floors"] = max(int(imported_object.get("floors", 1)), 1)

        for raw_contractor in imported_object.get("contractors", []):
            contractor = normalize_contractor_name(raw_contractor)
            if not contractor:
                continue
            existing = next(
                (
                    item for item in contractors
                    if contractor_lookup_key(item) == contractor_lookup_key(contractor)
                ),
                None,
            )
            if existing is None:
                contractors[contractor] = []
                existing = contractor
                contractors_changed = True
            if address not in contractors[existing]:
                contractors[existing].append(address)
                contractors_changed = True

    if contractors_changed:
        desired = {
            name: list(addresses)
            for name, addresses in contractors.items()
        }
        def merge_imported_contractors(latest):
            for name, addresses in desired.items():
                target = next(
                    (item for item in latest if contractor_lookup_key(item) == contractor_lookup_key(name)),
                    name,
                )
                latest.setdefault(target, [])
                for address in addresses:
                    if address not in latest[target]:
                        latest[target].append(address)
        update_contractors(merge_imported_contractors)

users = {}


SUPER_ADMIN_USERNAME = "demo_admin"


def load_stored_users():
    if not os.path.exists(USERS_FILE):
        return {}

    with open(USERS_FILE, "r", encoding="utf-8") as file:
        stored_users = json.load(file)

    if not isinstance(stored_users, dict):
        raise RuntimeError("Некорректный формат users.json")

    return stored_users


def update_users(mutator):
    with interprocess_lock(USERS_LOCK_FILE):
        latest_users = load_stored_users()
        mutator(latest_users)
        atomic_write_json(USERS_FILE, latest_users, 0o600)
        return latest_users



def initialize_users():
    stored_users = load_stored_users()

    if not stored_users:
        raise RuntimeError("users.json must contain at least one user")

    for username, user in stored_users.items():
        if not isinstance(username, str) or not username:
            raise RuntimeError("Invalid username in users.json")
        if not isinstance(user, dict):
            raise RuntimeError(f"Invalid user record: {username}")
        if "password" in user:
            raise RuntimeError(f"Plaintext password is forbidden: {username}")
        if not isinstance(user.get("password_hash"), str) or not user["password_hash"]:
            raise RuntimeError(f"Missing password hash: {username}")
        if not isinstance(user.get("role"), str) or not user["role"]:
            raise RuntimeError(f"Missing role: {username}")

    super_admin = stored_users.get(SUPER_ADMIN_USERNAME)
    if not super_admin or super_admin.get("role") != "super_admin":
        raise RuntimeError("Super Admin is missing or invalid in users.json")

    os.chmod(USERS_FILE, 0o600)
    users.clear()
    users.update(stored_users)



def password_is_valid(user, password):
    password_hash = user.get("password_hash")
    if not isinstance(password_hash, str) or not password_hash:
        return False
    try:
        return check_password_hash(password_hash, password)
    except (TypeError, ValueError):
        return False



ADMIN_PERMISSIONS = frozenset({
    "view_admin_dashboard",
    "view_object_summary",
    "select_contractor",
    "export_pdf",
    "complete_work",
    "edit_completion_date",
})

SUPER_ADMIN_PERMISSIONS = ADMIN_PERMISSIONS | frozenset({
    "import_excel",
    "manage_contractors",
    "manage_system",
})

ROLE_PERMISSIONS = {
    "contractor": frozenset(),
    # "customer" is the existing application role of the ordinary administrator.
    "customer": ADMIN_PERMISSIONS,
    # Keep the canonical name available for current and future stored users.
    "admin": ADMIN_PERMISSIONS,
    "super_admin": SUPER_ADMIN_PERMISSIONS,
}


def current_user_role():
    username = session.get("user")
    return users.get(username, {}).get("role", "")


def current_user_can(permission):
    return permission in ROLE_PERMISSIONS.get(
        current_user_role(),
        frozenset()
    )


def current_user_is_admin():
    return current_user_can("view_admin_dashboard")


def current_user_is_super_admin():
    return current_user_can("manage_system")


initialize_users()


@app.context_processor
def inject_current_user_access():
    return {
        "current_user_role": current_user_role(),
        "is_admin": current_user_is_admin(),
        "is_super_admin": current_user_is_super_admin(),
        "can_manage_contractors": current_user_can("manage_contractors"),
        "can_view_object_summary": current_user_can("view_object_summary"),
        "can_complete_work": current_user_can("complete_work"),
        "can_edit_completion_date": current_user_can("edit_completion_date"),
    }

PDF_FONT_NAME = "MayakSans"
PDF_FONT_CANDIDATES = (
    # Ubuntu / production
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    # macOS / local development
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/Library/Fonts/Arial.ttf",
    # Windows
    "C:/Windows/Fonts/arial.ttf",
)


def register_pdf_font():
    """Register a Unicode font available on the current operating system."""
    if PDF_FONT_NAME in pdfmetrics.getRegisteredFontNames():
        return PDF_FONT_NAME

    for font_path in PDF_FONT_CANDIDATES:
        if os.path.isfile(font_path):
            pdfmetrics.registerFont(TTFont(PDF_FONT_NAME, font_path))
            return PDF_FONT_NAME

    raise RuntimeError(
        "Не найден шрифт для PDF. Установите DejaVu Sans "
        "(пакет fonts-dejavu-core на Ubuntu)."
    )

def load_data():
    if not os.path.exists(DATA_FILE):
        return {}

    with open(DATA_FILE, "r", encoding="utf-8") as file:
        return json.load(file)


def save_data(data=None):
    """Atomically persist data so readers never see a partial JSON file."""
    global flat_data

    if data is not None:
        flat_data = data

    temporary_file = f"{DATA_FILE}.{os.getpid()}.{uuid.uuid4().hex}.tmp"

    try:
        with open(temporary_file, "w", encoding="utf-8") as file:
            json.dump(
                flat_data,
                file,
                ensure_ascii=False,
                indent=4
            )
            file.flush()
            os.fsync(file.fileno())

        os.replace(temporary_file, DATA_FILE)
    finally:
        if os.path.exists(temporary_file):
            os.remove(temporary_file)


def update_data(mutator):
    """Serialize the complete read-modify-write transaction across workers."""
    global flat_data
    with interprocess_lock(DATA_LOCK_FILE):
        latest = load_data()
        result = mutator(latest)
        atomic_write_json(DATA_FILE, latest)
        flat_data = latest
        return result


def normalize_work_status(status):
    """Return MAYAK's canonical colour status for current and legacy values."""
    return {
        "green": "green",
        "done": "green",
        "yellow": "yellow",
        "check": "yellow",
        "red": "red",
        "denied": "red",
        "white": "white",
        "notdone": "white",
    }.get(str(status or "").strip().lower(), "white")


def get_selection_flats(data, contractor, address, system):
    """Use the same working board as Dashboard to collect room statuses."""
    system_data = data.get(address, {}).get(system, {})
    excel_key = f"{contractor}|{address}|{system}"
    excel_board = data.get("excel_chessboards", {}).get(excel_key)
    flats = []

    if excel_board:
        board = normalize_chessboard(excel_board)
        for room in board.get("rooms", []):
            room_key = str(
                room.get("room_key") or room.get("room_number") or ""
            ).strip()
            if not room_key:
                continue
            flats.append({
                "num": room_key,
                "status": normalize_work_status(
                    system_data.get(room_key, {}).get("status")
                ),
            })
        return flats

    object_settings = objects.get(address, {})
    for number in range(1, int(object_settings.get("flats", 0)) + 1):
        flats.append({
            "num": number,
            "status": normalize_work_status(
                system_data.get(str(number), {}).get("status")
            ),
        })

    return flats


def build_object_summary(data, contractor, address, system):
    """Build a live summary without persisting a second statistics copy."""
    flats = get_selection_flats(data, contractor, address, system)
    counts = {
        color: sum(1 for flat in flats if flat["status"] == color)
        for color in ("green", "yellow", "red", "white")
    }
    total = len(flats)
    progress = round(counts["green"] / total * 100) if total else 0
    spill = data.get("spill_data", {}).get(
        f"{address}_{system}",
        {"top": 0, "bottom": 0, "single": 0},
    )

    result = {
        "contractor": contractor,
        "object": address,
        "system": system,
        "total": total,
        "done": counts["green"],
        "check": counts["yellow"],
        "denied": counts["red"],
        "notdone": counts["white"],
        "progress": progress,
    }
    if system == "ТС":
        result["upper_rozliv"] = int(spill.get("top", 0) or 0)
        result["lower_rozliv"] = int(spill.get("bottom", 0) or 0)
    else:
        result["rozliv"] = int(spill.get("single", 0) or 0)
    return result


flat_data = load_data()
sync_imported_objects(flat_data)

@app.route("/")
def home():
    global flat_data

    if "user" not in session:
        return redirect("/login")

    # Gunicorn workers do not share memory. Always use the latest disk state.
    flat_data = load_data()
    refresh_contractors()
    sync_imported_objects(flat_data)

    current_contractor = request.args.get(
        "contractor",
        ""
    )

    if "user" in session:
        username = session["user"]

        if users[username]["role"] == "contractor":
            current_contractor = users[username]["name"]

    current_object = request.args.get(
        "object",
        ""
    )

    current_system = request.args.get(
        "system",
        ""
    )
    common_works = []


    available_objects = contractors.get(
        current_contractor,
        []
    )

    show_object = (
        current_contractor != ""
        and current_object != ""
        and current_system != ""
    )
    current_excel_chessboard = None

    if show_object:
        excel_key = f"{current_contractor}|{current_object}|{current_system}"

        current_excel_chessboard = flat_data.get(
            "excel_chessboards",
             {}
        ).get(excel_key)

    flats = []
    progress = 0

    green_count = 0
    yellow_count = 0
    red_count = 0
    white_count = 0

    if show_object:

        object_data = flat_data.get(
            current_object,
            {}
        )

        system_data = object_data.get(
            current_system,
            {}
        )

        if current_excel_chessboard:
            current_excel_chessboard = prepare_chessboard_for_display(
                current_excel_chessboard,
                system_data,
                current_contractor,
                current_object,
                current_system,
            )

        flats = get_selection_flats(
            flat_data,
            current_contractor,
            current_object,
            current_system,
        )

        done_count = sum(
            1 for flat in flats
            if flat["status"] == "green"
        )

        green_count = sum(
            1 for flat in flats
           if flat["status"] == "green"
        )

        yellow_count = sum(
            1 for flat in flats
            if flat["status"] == "yellow"
        )

        red_count = sum(
            1 for flat in flats
            if flat["status"] == "red"
        )

        white_count = sum(
            1 for flat in flats
            if flat["status"] == "white"
        )
        if flats:
            progress = round(
                done_count / len(flats) * 100
            )

    # Розливы
    spill_data = flat_data.get("spill_data", {})

    object_key = f"{current_object}_{current_system}"

    current_spill = spill_data.get(
        object_key,
        {
            "top": 0,
            "bottom": 0,
            "single": 0
        }
    )

    return render_template(
        "index.html",
        flats=flats,
        progress=progress,
        spill=current_spill,
        green_count=green_count,
        yellow_count=yellow_count,
        red_count=red_count,
        white_count=white_count,
        common_works=common_works,
        objects=objects,
        contractors=contractors,
        current_contractor=current_contractor,
        available_objects=available_objects,
        current_object=current_object,
        current_system=current_system,
        current_excel_chessboard=current_excel_chessboard,
        floors=objects.get(
            current_object,
            {}
        ).get(
            "floors",
            0
        ),
        is_admin=current_user_is_admin(),
        is_super_admin=current_user_is_super_admin(),
        show_object=show_object
    )


@app.route("/api/object-summary/options")
def object_summary_options():
    global flat_data

    if "user" not in session:
        abort(401)
    if not current_user_can("view_object_summary"):
        abort(403)

    flat_data = load_data()
    refresh_contractors()
    sync_imported_objects(flat_data)

    contractor = request.args.get("contractor", "").strip()
    address = request.args.get("object", "").strip()
    contractor_objects = []
    systems = []

    if contractor:
        if contractor not in contractors:
            return jsonify({"error": "Подрядчик не найден."}), 404
        contractor_objects = list(contractors[contractor])

    if address:
        if not contractor or address not in contractor_objects:
            return jsonify({"error": "Объект не принадлежит подрядчику."}), 403
        systems = list(objects.get(address, {}).get("systems", []))

    return jsonify({
        "contractors": list(contractors),
        "objects": contractor_objects,
        "systems": systems,
    })


@app.route("/api/object-summary")
def object_summary():
    global flat_data

    if "user" not in session:
        abort(401)
    if not current_user_can("view_object_summary"):
        abort(403)

    contractor = request.args.get("contractor", "").strip()
    address = request.args.get("object", "").strip()
    system = request.args.get("system", "").strip()

    flat_data = load_data()
    refresh_contractors()
    sync_imported_objects(flat_data)

    if contractor not in contractors:
        return jsonify({"error": "Подрядчик не найден."}), 404
    if address not in contractors[contractor]:
        return jsonify({"error": "Объект не принадлежит подрядчику."}), 403
    if system not in objects.get(address, {}).get("systems", []):
        return jsonify({"error": "Инженерная система не найдена."}), 404

    return jsonify(build_object_summary(
        flat_data,
        contractor,
        address,
        system,
    ))


@app.route(
    "/flat/<int:flat_num>",
    methods=["GET", "POST"]
)
def flat(flat_num):
    return room_card_response(str(flat_num), "apartment")


@app.route(
    "/room/<path:room_key>",
    methods=["GET", "POST"]
)
def imported_room(room_key):
    return room_card_response(str(room_key), "non_residential")


def room_card_response(room_key, room_kind):
    """Use MAYAK's existing room card for imported non-residential slots."""

    global flat_data
    flat_data = load_data()

    room_key = str(room_key).strip()
    if not room_key:
        abort(404)

    if room_kind == "apartment":
        room_label = f"Квартира №{room_key}"
        room_short_label = f"Квартира {room_key}"
    else:
        room_label = f"Нежилое помещение {room_key.split('::', 1)[0]}"
        room_short_label = f"Помещение {room_key.split('::', 1)[0]}"

    current_contractor = request.args.get(
        "contractor",
        ""
    )

    current_object = request.args.get(
        "object",
        ""
    )

    current_system = request.args.get(
        "system",
        ""
    )

    if request.method == "POST":

        current_contractor = request.form.get(
            "contractor",
            ""
        )

        current_object = request.form.get(
            "object",
            ""
        )

        current_system = request.form.get(
            "system",
            ""
        )

        photos = request.files.getlist("photos")

        photo_names = []

        for photo in photos:

            if not photo or not photo.filename:
                continue

            original_name = secure_filename(photo.filename)

            if "." in original_name:
                ext = original_name.rsplit(".", 1)[1].lower()
            else:
                ext = "jpg"

            filename = f"{uuid.uuid4()}.{ext}"

            photo.save(
                os.path.join(
                    UPLOAD_FOLDER,
                    filename
                )
            )

            photo_names.append(filename)

        status = request.form.get("status")

        if status == "green" and not current_user_can("complete_work"):
            abort(403)

        can_edit_date = current_user_can("edit_completion_date")
        def mutate_flat(data):
            system_data = data.setdefault(current_object, {}).setdefault(current_system, {})
            old = system_data.get(room_key, {})
            all_photos = list(old.get("photos", [])) + photo_names
            if status == "yellow" and not all_photos:
                raise ValueError("photo_required")
            new_date = request.form.get("date")
            if not can_edit_date and old.get("date"):
                new_date = old["date"]
            system_data[room_key] = {
                "status": status,
                "contractor": current_contractor,
                "date": new_date,
                "photos": all_photos,
            }
        try:
            update_data(mutate_flat)
        except Exception as error:
            for filename in photo_names:
                path = os.path.join(UPLOAD_FOLDER, filename)
                if os.path.exists(path):
                    os.remove(path)
            if isinstance(error, ValueError) and str(error) == "photo_required":
                return render_template(
                    "flat.html", flat_num=room_key,
                    room_label=room_label,
                    room_short_label=room_short_label,
                    flat_info={"status": "white", "contractor": current_contractor, "date": "", "photos": []},
                    current_contractor=current_contractor, current_object=current_object,
                    current_system=current_system,
                    error="Для отправки квартиры на проверку необходимо загрузить хотя бы одну фотографию."
                )
            raise

        return redirect(
            f"/?contractor={current_contractor}"
            f"&object={current_object}"
            f"&system={current_system}"
        )

    flat_info = (
        flat_data
        .get(current_object, {})
        .get(current_system, {})
        .get(
            room_key,
            {
                "status": "white",
                "contractor": "",
                "date": "",
                "photos": []
            }
        )
    )

    return render_template(
        "flat.html",
        flat_num=room_key,
        room_label=room_label,
        room_short_label=room_short_label,
        flat_info=flat_info,
        current_contractor=current_contractor,
        current_object=current_object,
        current_system=current_system
    )

@app.route(
    "/common/<work_name>",
    methods=["GET", "POST"]
)
def common_work(work_name):
    global flat_data
    flat_data = load_data()

    current_contractor = request.args.get(
        "contractor",
        ""
    )

    current_object = request.args.get(
        "object",
        ""
    )

    current_system = request.args.get(
        "system",
        ""
    )

    if request.method == "POST":

        current_contractor = request.form.get(
            "contractor",
            ""
        )

        current_object = request.form.get(
            "object",
            ""
        )

        current_system = request.form.get(
            "system",
            ""
        )

        photos = request.files.getlist("photos")

        photo_names = []

        for photo in photos:

            if photo and photo.filename:

                original_name = secure_filename(photo.filename)
                if "." in original_name:
                    ext = original_name.rsplit(".", 1)[1].lower()
                else:
                    ext = "jpg"
                filename = f"{uuid.uuid4()}.{ext}"

                photo.save(
                    os.path.join(
                        UPLOAD_FOLDER,
                        filename
                    )
                )

                photo_names.append(filename)

        can_edit_date = current_user_can("edit_completion_date")
        def mutate_common(data):
            common = data.setdefault(current_object, {}).setdefault(current_system, {}).setdefault("_common", {})
            old = common.get(work_name, {"status": "white", "date": "", "photos": []})
            new_date = request.form.get("date")
            if not can_edit_date and old.get("date"):
                new_date = old["date"]
            common[work_name] = {
                "status": request.form.get("status"), "date": new_date,
                "photos": list(old.get("photos", [])) + photo_names,
            }
        try:
            update_data(mutate_common)
        except Exception:
            for filename in photo_names:
                path = os.path.join(UPLOAD_FOLDER, filename)
                if os.path.exists(path):
                    os.remove(path)
            raise

    work_info = (
        flat_data
        .get(current_object, {})
        .get(current_system, {})
        .get("_common", {})
        .get(
            work_name,
            {
                "status": "white",
                "date": "",
                "photos": []
            }
        )
    )

    return render_template(
        "common.html",
        work_name=work_name,
        work_info=work_info,
        current_contractor=current_contractor,
        current_object=current_object,
        current_system=current_system
    )

@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        username = request.form.get("username")
        password = request.form.get("password")

        if username in users:

            if password_is_valid(users[username], password):

                session["user"] = username
                return redirect("/")

        return "Неверный логин или пароль"

    return render_template("login.html")

@app.route("/save_spill", methods=["POST"])
def save_spill():
    global flat_data

    current_object = request.args.get("object", "")
    current_system = request.args.get("system", "")

    object_key = f"{current_object}_{current_system}"

    def mutate_spill(data):
        spill_data = data.setdefault("spill_data", {})
        if current_system == "ТС":
            spill_data[object_key] = {
                "top": int(request.form.get("top", 0)),
                "bottom": int(request.form.get("bottom", 0)),
            }
        else:
            spill_data[object_key] = {"single": int(request.form.get("single", 0))}
    update_data(mutate_spill)

    return redirect(
        f"/?contractor={request.args.get('contractor','')}"
        f"&object={current_object}"
        f"&system={current_system}"
    )

@app.route("/download_pdf")
def download_pdf():
    global flat_data

    if not current_user_can("export_pdf"):
        abort(403)

    flat_data = load_data()

    current_contractor = request.args.get("contractor", "")
    current_object = request.args.get("object", "")
    current_system = request.args.get("system", "")

    pdf_font_name = register_pdf_font()
    pdf_buffer = BytesIO()

    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=landscape(A4),
        rightMargin=1 * cm,
        leftMargin=1 * cm,
        topMargin=1 * cm,
        bottomMargin=1 * cm
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "title",
        parent=styles["Heading1"],
        fontName=pdf_font_name,
        fontSize=22,
        alignment=TA_CENTER,
        textColor=HexColor("#0d47a1"),
        spaceAfter=16
    )

    normal = ParagraphStyle(
        "normal",
        parent=styles["Normal"],
        fontName=pdf_font_name,
        fontSize=10,
        leading=14
    )

    story = []

    story.append(Paragraph("ОТЧЕТ ПО ОБЪЕКТУ", title_style))

    object_data = flat_data.get(current_object, {})
    system_data = object_data.get(current_system, {})

    flats = []

    for num in range(1, objects[current_object]["flats"] + 1):

        status = "white"

        if str(num) in system_data:
            status = system_data[str(num)]["status"]

        flats.append({
            "num": num,
            "status": status
        })

    green_count = sum(1 for flat in flats if flat["status"] == "green")
    yellow_count = sum(1 for flat in flats if flat["status"] == "yellow")
    red_count = sum(1 for flat in flats if flat["status"] == "red")
    white_count = sum(1 for flat in flats if flat["status"] == "white")

    if len(flats) > 0:
        progress = round(green_count / len(flats) * 100)
    else:
        progress = 0

    story.append(Paragraph(f"<b>Дата:</b> {datetime.now().strftime('%d.%m.%Y %H:%M')}", normal))
    story.append(Paragraph(f"<b>Подрядчик:</b> {current_contractor}", normal))
    story.append(Paragraph(f"<b>Объект:</b> {current_object}", normal))
    story.append(Paragraph(f"<b>Система:</b> {current_system}", normal))
    story.append(Paragraph(f"<b>Готовность объекта:</b> {progress} %", normal))

    story.append(Spacer(1, 0.3 * cm))

    story.append(Paragraph(f"<b>Общее количество помещений:</b> {len(flats)}", normal))
    story.append(Paragraph(f"<b>Выполнено:</b> {green_count}", normal))
    story.append(Paragraph(f"<b>Отказ / нет доступа:</b> {red_count}", normal))
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph("<b>Шахматка помещений:</b>", normal))
    story.append(Spacer(1, 0.3 * cm))

    floors_count = objects[current_object]["floors"]
    flats_count = objects[current_object]["flats"]
    flats_per_floor = flats_count // floors_count

    table_data = []

    for floor in range(floors_count, 0, -1):

        row = [f"{floor} этаж"]

        start = (floor - 1) * flats_per_floor
        end = floor * flats_per_floor

        for flat in flats[start:end]:
            row.append(f"Кв.{flat['num']}")

        table_data.append(row)
    page_width = landscape(A4)[0] - 2 * cm
    page_height = landscape(A4)[1] - 7 * cm

    max_cols = max(len(row) for row in table_data)
    max_rows = len(table_data)

    col_width = page_width / max_cols
    row_height = page_height / max_rows

    table = Table(
        table_data,
        colWidths=[col_width] * max_cols,
        rowHeights=[row_height] * max_rows
    )

    table_style = TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTNAME", (0, 0), (-1, -1), pdf_font_name),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#2c3e50")),
        ("TEXTCOLOR", (0, 0), (0, -1), colors.white),

    ])

    status_colors = {
        "green": colors.HexColor("#32CD32"),
        "yellow": colors.white,
        "red": colors.HexColor("#FF3B30"),
        "white": colors.white
    }

    for row_index, floor in enumerate(range(floors_count, 0, -1)):

        start = (floor - 1) * flats_per_floor
        end = floor * flats_per_floor

        for col_index, flat in enumerate(flats[start:end], start=1):

            table_style.add(
                "BACKGROUND",
                (col_index, row_index),
                (col_index, row_index),
                status_colors.get(flat["status"], colors.white)
            )

    table.setStyle(table_style)

    story.append(table)

    story.append(Spacer(1, 0.7 * cm))

    story.append(
        Paragraph(
            "<b>Розлив:</b>",
            normal
        )
    )

    spill_data = flat_data.get("spill_data", {})

    object_key = f"{current_object}_{current_system}"

    spill = spill_data.get(
        object_key,
        {
            "top": 0,
            "bottom": 0,
            "single": 0
        }
    )

    if current_system == "ТС":

        story.append(
            Paragraph(
                f"Верхний розлив: {spill.get('top', 0)} %",
                normal
            )
        )

        story.append(
            Paragraph(
                f"Нижний розлив: {spill.get('bottom', 0)} %",
                normal
            )
        )

    else:

        story.append(
            Paragraph(
                f"Розлив: {spill.get('single', 0)} %",
                normal
            )
        )

    doc.build(story)
    pdf_buffer.seek(0)

    return send_file(
        pdf_buffer,
        as_attachment=True,
        download_name="mayak_object_report.pdf",
        mimetype="application/pdf"
    )


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


@app.route("/import_excel", methods=["GET", "POST"])
def import_excel():
    if "user" not in session:
        return redirect("/login")

    if not current_user_can("import_excel"):
        abort(403)

    refresh_contractors()

    if request.method == "POST":
        contractor = request.form.get("contractor", "").strip()
        address = request.form.get("address", "").strip()
        system = request.form.get("system", "").strip()

        if contractor not in contractors or not address:
            abort(400)

        if system not in IMPORT_SYSTEMS:
            abort(400)

        file = request.files.get("excel_file")

        if not file or not file.filename:
            return render_template(
                "import_excel.html",
                contractors=contractors,
                import_systems=IMPORT_SYSTEMS,
                error="Выберите файл Excel."
            ), 400

        extension = os.path.splitext(file.filename)[1].lower()
        if extension not in IMPORT_EXTENSIONS:
            return render_template(
                "import_excel.html",
                contractors=contractors,
                import_systems=IMPORT_SYSTEMS,
                error="Поддерживаются файлы XLSX и XLSM."
            ), 400

        token = uuid.uuid4().hex
        temporary_excel = os.path.join(
            IMPORT_PREVIEW_FOLDER,
            f"{token}{extension}"
        )

        try:
            file.save(temporary_excel)
            chessboard = parse_excel_chessboard(temporary_excel)
        except Exception as error:
            return render_template(
                "import_excel.html",
                contractors=contractors,
                import_systems=IMPORT_SYSTEMS,
                error=f"Не удалось прочитать Excel: {error}"
            ), 400
        finally:
            if os.path.exists(temporary_excel):
                os.remove(temporary_excel)

        save_import_preview(token, {
            "contractor": contractor,
            "address": address,
            "system": system,
            "source_filename": secure_filename(file.filename),
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "chessboard": chessboard,
        })

        return redirect(url_for("import_excel_preview", token=token))

    return render_template(
        "import_excel.html",
        contractors=contractors,
        import_systems=IMPORT_SYSTEMS
    )


@app.route("/api/contractors", methods=["GET", "POST"])
def contractor_directory():
    if "user" not in session:
        abort(401)

    refresh_contractors()

    if request.method == "GET":
        return jsonify({"contractors": list(contractors)})

    if not current_user_can("manage_contractors"):
        abort(403)

    payload = request.get_json(silent=True) or request.form
    name = normalize_contractor_name(payload.get("name", ""))

    if not name:
        return jsonify({"error": "Введите название подрядчика."}), 400

    existing = next(
        (
            item for item in contractors
            if contractor_lookup_key(item) == contractor_lookup_key(name)
        ),
        None,
    )
    if existing is not None:
        return jsonify({
            "error": "Подрядчик уже существует",
            "existing_contractor": existing,
        }), 409

    def add_contractor(latest):
        duplicate = next(
            (item for item in latest if contractor_lookup_key(item) == contractor_lookup_key(name)),
            None,
        )
        if duplicate is not None:
            raise ValueError(duplicate)
        latest[name] = []
    try:
        update_contractors(add_contractor)
    except ValueError as error:
        return jsonify({"error": "Подрядчик уже существует", "existing_contractor": str(error)}), 409

    return jsonify({
        "contractor": name,
        "contractors": list(contractors),
    }), 201


@app.route("/import_excel/preview/<token>")
def import_excel_preview(token):
    if "user" not in session:
        return redirect("/login")
    if not current_user_can("import_excel"):
        abort(403)

    payload = load_import_preview(token)
    chessboard = normalize_chessboard(payload["chessboard"])
    return render_template(
        "import_excel_preview.html",
        token=token,
        payload=payload,
        chessboard=chessboard,
    )


@app.route("/import_excel/confirm/<token>", methods=["POST"])
def confirm_excel_import(token):
    global flat_data

    if "user" not in session:
        return redirect("/login")
    if not current_user_can("import_excel"):
        abort(403)

    refresh_contractors()
    payload = load_import_preview(token)
    contractor = payload.get("contractor", "")
    address = payload.get("address", "")
    system = payload.get("system", "")

    if contractor not in contractors or not address or system not in IMPORT_SYSTEMS:
        abort(400)

    assigned_contractors = [
        contractor_name
        for contractor_name, addresses in contractors.items()
        if address in addresses
    ]
    def mutate_import(data):
        return apply_import_to_data(
            data, contractor, address, system, payload["chessboard"],
            existing_object=objects.get(address), assigned_contractors=assigned_contractors,
        )
    result = update_data(mutate_import)
    sync_imported_objects(flat_data)
    delete_import_preview(token)

    return redirect(url_for("excel_chessboard", key=result["key"]))


@app.route("/import_excel/cancel/<token>", methods=["POST"])
def cancel_excel_import(token):
    if "user" not in session:
        return redirect("/login")
    if not current_user_can("import_excel"):
        abort(403)
    delete_import_preview(token)
    return redirect(url_for("import_excel"))


@app.route("/excel_chessboard")
def excel_chessboard():
    global flat_data

    if "user" not in session:
        return redirect("/login")

    flat_data = load_data()

    requested_key = request.args.get("key")
    key = requested_key or flat_data.get("last_excel_key")

    chessboard = None

    if key:
        chessboard = flat_data.get("excel_chessboards", {}).get(key)

    if not chessboard and not requested_key:
        chessboard = flat_data.get("excel_chessboard")

    if not chessboard:
        return "Шахматка ещё не загружена"

    contractor = ""
    address = ""
    system = ""
    if key and key.count("|") >= 2:
        contractor, address, system = key.split("|", 2)

    system_data = flat_data.get(address, {}).get(system, {})
    chessboard = prepare_chessboard_for_display(
        chessboard,
        system_data,
        contractor,
        address,
        system,
    )

    return render_template(
        "excel_chessboard.html",
        chessboard=chessboard,
        contractor=contractor,
        address=address,
        system=system,
    )

if __name__ == "__main__":
    app.run(
        host="127.0.0.1",
        port=5001,
        debug=os.environ.get("FLASK_DEBUG") == "1"
    )
