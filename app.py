from __future__ import annotations

from flask import Flask, jsonify, render_template, request, session, redirect, url_for, send_file, flash
from pathlib import Path
from io import BytesIO
from datetime import datetime
import os
import shutil
import socket
import threading
import time
import webbrowser

from dashboard.metrics import DashboardStore
from dashboard.xlsx_export import build_branch_request_xlsx
from dashboard.ppt_export import build_presentation

BASE = Path(__file__).resolve().parent
DATA_FILE = BASE / "data" / "MC_Dashboard_IMPORT.xlsx"
UPLOAD_DIR = BASE / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("SCM_SECRET_KEY", "change-this-secret-before-production")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
ADMIN_PASSWORD = os.environ.get("SCM_ADMIN_PASSWORD", "admin123")
ACTIVE_IMPORT = UPLOAD_DIR / "active_import.xlsx"
store = DashboardStore(ACTIVE_IMPORT if ACTIVE_IMPORT.exists() else DATA_FILE)


def role():
    return "admin" if session.get("is_admin") else "guest"


def admin_required(fn):
    from functools import wraps
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if role() != "admin":
            return jsonify({"error": "Admin access required."}), 403
        return fn(*args, **kwargs)
    return wrapper


@app.get("/")
def index():
    return render_template("index.html", role=role())


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        password = request.form.get("password", "")
        if password == ADMIN_PASSWORD:
            session["is_admin"] = True
            return redirect(url_for("index"))
        flash("Invalid admin password.", "error")
    return render_template("login.html")


@app.post("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))


@app.get("/api/bootstrap")
def api_bootstrap():
    return jsonify(store.bootstrap(role()))


@app.get("/api/area")
def api_area():
    return jsonify(store.area_dashboard(request.args.get("area", "Overall")))


@app.get("/api/branch")
def api_branch():
    branch = request.args.get("branch", "")
    return jsonify(store.branch_dashboard(branch))


@app.get("/api/branch-models")
def api_branch_models():
    branch = request.args.get("branch", "")
    return jsonify({"branch": branch, "models": store.branch_models(branch)})


@app.get("/api/model")
def api_model():
    branch = request.args.get("branch", "")
    model = request.args.get("model", "")
    item = store.model_lookup(branch, model)
    if not item:
        return jsonify({"error": "Model not found for selected branch."}), 404
    return jsonify(item)


@app.post("/admin/import")
@admin_required
def admin_import():
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "Please select an .xlsx import workbook."}), 400
    tmp = UPLOAD_DIR / "candidate_import.xlsx"
    f.save(tmp)
    check = DashboardStore.validate(tmp)
    if not check.ok:
        tmp.unlink(missing_ok=True)
        return jsonify({"error": check.message}), 400
    active = UPLOAD_DIR / "active_import.xlsx"
    shutil.move(tmp, active)
    store.load(active)
    return jsonify({"ok": True, "message": "Import completed. Dashboard data refreshed.", "generated_at": store.generated_at.isoformat()})


@app.get("/admin/export/pptx")
@admin_required
def admin_export_pptx():
    area = request.args.get("area", "Overall")
    branch = request.args.get("branch", "")
    data = store.bootstrap(role())
    ppt = build_presentation(data, store.area_dashboard(area), store.branch_dashboard(branch))
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    return send_file(BytesIO(ppt), mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation", as_attachment=True, download_name=f"SCM_IDP_Executive_Control_Tower_{stamp}.pptx")


@app.post("/admin/export/request")
@admin_required
def admin_export_request():
    payload = request.get_json(force=True, silent=True) or {}
    branch = str(payload.get("branch", "")).strip()
    items = payload.get("items") or []
    if not branch or not items:
        return jsonify({"error": "Branch and at least one requested item are required."}), 400
    verified = []
    area = ""
    for item in items:
        model = str(item.get("model", "")).strip()
        base = store.model_lookup(branch, model)
        if not base:
            continue
        qty = float(item.get("requested_qty", 0) or 0)
        new_doi = ((base["inventory"] + qty) / base["avg_daily_sale"]) if base["avg_daily_sale"] > 0 else "N/A"
        area = base["area"]
        verified.append({**base, "requested_qty": qty, "remarks": str(item.get("remarks", "")), "new_doi": round(new_doi, 2) if isinstance(new_doi, float) else new_doi})
    if not verified:
        return jsonify({"error": "No valid request lines."}), 400
    xlsx = build_branch_request_xlsx(branch, area, verified)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    safe_branch = "_".join(branch.split())
    return send_file(BytesIO(xlsx), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name=f"Branch_Request_{safe_branch}_{stamp}.xlsx")


@app.get("/health")
def health():
    return {"status": "ok", "role": role(), "records": len(store.raw_records)}


def _port_is_available(host: str, port: int) -> bool:
    """Return True only when Windows/Linux allows us to bind this host/port."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.bind((host, port))
        return True
    except OSError:
        return False


def _pick_port(host: str) -> int:
    """Choose a usable port, avoiding Windows reserved/excluded ports."""
    candidates = []
    env_port = os.environ.get("PORT")
    if env_port:
        try:
            candidates.append(int(env_port))
        except ValueError:
            print(f"WARNING: Ignoring invalid PORT={env_port!r}")

    # 5055 is the preferred local dashboard port. Alternatives are tried automatically.
    candidates.extend([5055, 8050, 8088, 8765, 9000, 5000])

    seen = set()
    for port in candidates:
        if port in seen or not (1 <= port <= 65535):
            continue
        seen.add(port)
        if _port_is_available(host, port):
            return port

    # Ask Windows for any free ephemeral port as a final fallback.
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind((host, 0))
        return sock.getsockname()[1]


def _open_browser(url: str) -> None:
    if os.environ.get("SCM_NO_BROWSER") == "1":
        return

    def opener():
        time.sleep(1.2)
        try:
            webbrowser.open(url)
        except Exception:
            pass

    threading.Thread(target=opener, daemon=True).start()


if __name__ == "__main__":
    # Local-only by default: safer and avoids Windows firewall/reserved-socket issues.
    host = os.environ.get("SCM_HOST", "127.0.0.1")
    port = _pick_port(host)
    url = f"http://{host}:{port}"

    print("=" * 68)
    print(" SCM Inventory & Distribution Planning Dashboard")
    print(f" Open in browser: {url}")
    print(" Admin login: /login")
    print(" Press CTRL+C to stop the dashboard.")
    print("=" * 68)

    _open_browser(url)
    app.run(
        host=host,
        port=port,
        debug=os.environ.get("FLASK_DEBUG") == "1",
        use_reloader=False,
    )
