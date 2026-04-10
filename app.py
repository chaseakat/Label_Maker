import os
import shutil
import time
import uuid
import zipfile
from pathlib import Path

from flask import Flask, abort, flash, redirect, render_template, request, send_file, session, url_for
from werkzeug.utils import secure_filename

from label_engine import LabelRunConfig, generate_manual_labels, run_label_generation


app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret")

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
RESULTS_DIR = BASE_DIR / "results"
RESULT_CACHE_DIR = RESULTS_DIR / "cache"
DEFAULT_OUTPUT_DIR = BASE_DIR / "generated_labels"
PERSISTENT_LOGO_DIR = BASE_DIR / "persistent_logo"

for directory in [UPLOAD_DIR, RESULTS_DIR, RESULT_CACHE_DIR, DEFAULT_OUTPUT_DIR, PERSISTENT_LOGO_DIR]:
    directory.mkdir(parents=True, exist_ok=True)

ALLOWED_EXTENSIONS = {"pdf", "png", "jpg", "jpeg", "bmp", "tif", "tiff", "webp"}
ALLOWED_IMAGE_EXTENSIONS = {"png", "jpg", "jpeg", "bmp", "webp"}


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def get_persisted_logo_paths() -> tuple[Path | None, Path | None]:
    image_path = None
    zip_path = None

    for path in PERSISTENT_LOGO_DIR.glob("current_logo.*"):
        if path.suffix.lower().lstrip(".") in ALLOWED_IMAGE_EXTENSIONS:
            image_path = path
            break

    zip_candidate = PERSISTENT_LOGO_DIR / "current_logo.zip"
    if zip_candidate.exists():
        zip_path = zip_candidate

    return image_path, zip_path


def clear_persisted_logo() -> None:
    for path in PERSISTENT_LOGO_DIR.glob("current_logo.*"):
        try:
            path.unlink()
        except FileNotFoundError:
            pass


def save_logo_inputs_to_persistent(logo_image, logo_zip):
    persisted_logo_image, persisted_logo_zip = get_persisted_logo_paths()
    logo_image_path = None
    logo_zip_path = None

    if logo_image and logo_image.filename:
        image_name = secure_filename(logo_image.filename)
        image_ext = Path(image_name).suffix.lower().lstrip(".")
        if image_ext not in ALLOWED_IMAGE_EXTENSIONS:
            raise ValueError("Invalid logo image type.")

        clear_persisted_logo()
        persisted_image = PERSISTENT_LOGO_DIR / f"current_logo.{image_ext}"
        logo_image.save(persisted_image)
        logo_image_path = persisted_image
    elif logo_zip and logo_zip.filename:
        zip_name = secure_filename(logo_zip.filename)
        if Path(zip_name).suffix.lower() != ".zip":
            raise ValueError("Invalid logo zip type.")

        clear_persisted_logo()
        persisted_zip = PERSISTENT_LOGO_DIR / "current_logo.zip"
        logo_zip.save(persisted_zip)
        logo_zip_path = persisted_zip
    else:
        logo_image_path = persisted_logo_image
        logo_zip_path = persisted_logo_zip

    return logo_image_path, logo_zip_path


def build_download_response(output_files):
    files = [Path(p) for p in output_files]
    if len(files) == 1:
        return send_file(files[0], as_attachment=True)

    zip_name = f"labels_{uuid.uuid4().hex}.zip"
    zip_path = RESULTS_DIR / zip_name
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in files:
            zf.write(file_path, arcname=file_path.name)
    return send_file(zip_path, as_attachment=True)


def cleanup_cached_results(max_age_seconds: int = 3600) -> None:
    now = time.time()
    for run_dir in RESULT_CACHE_DIR.iterdir():
        if not run_dir.is_dir():
            continue
        age = now - run_dir.stat().st_mtime
        if age > max_age_seconds:
            shutil.rmtree(run_dir, ignore_errors=True)


def cache_output_files(output_files):
    cleanup_cached_results()
    result_id = uuid.uuid4().hex
    target_dir = RESULT_CACHE_DIR / result_id
    target_dir.mkdir(parents=True, exist_ok=True)

    cached_names = []
    for output_file in output_files:
        src = Path(output_file)
        dst = target_dir / src.name
        shutil.copy2(src, dst)
        cached_names.append(src.name)

    session["latest_result_id"] = result_id
    session["latest_result_files"] = cached_names
    session["latest_result_created"] = int(time.time())
    return result_id, cached_names


def get_latest_cached_result():
    result_id = session.get("latest_result_id")
    files = session.get("latest_result_files") or []
    if not result_id or not files:
        return None, []

    run_dir = RESULT_CACHE_DIR / result_id
    if not run_dir.exists():
        return None, []

    existing = [name for name in files if (run_dir / name).exists()]
    if not existing:
        return None, []
    return result_id, existing


@app.route("/", methods=["GET"])
def index():
    cleanup_cached_results()
    persisted_logo_image, persisted_logo_zip = get_persisted_logo_paths()
    persisted_logo_label = "None saved"
    if persisted_logo_image:
        persisted_logo_label = f"Saved image: {persisted_logo_image.name}"
    elif persisted_logo_zip:
        persisted_logo_label = f"Saved zip: {persisted_logo_zip.name}"

    latest_result_id, latest_files = get_latest_cached_result()
    return render_template(
        "index.html",
        default_output_dir=str(DEFAULT_OUTPUT_DIR),
        persisted_logo_label=persisted_logo_label,
        latest_result_id=latest_result_id,
        latest_files=latest_files,
    )


@app.route("/run", methods=["POST"])
def run_generation():
    workorder = request.files.get("workorder_file")
    logo_image = request.files.get("logo_image")
    logo_zip = request.files.get("logo_zip")

    output_dir = (request.form.get("output_dir") or "").strip() or str(DEFAULT_OUTPUT_DIR)
    job_number_override = (request.form.get("job_number_override") or "").strip()
    job_name_override = (request.form.get("job_name_override") or "").strip()
    output_path = Path(output_dir).expanduser().resolve()
    output_path.mkdir(parents=True, exist_ok=True)

    if not workorder or not workorder.filename:
        flash("Please upload a work order PDF or image.", "error")
        return redirect(url_for("index"))

    if not allowed_file(workorder.filename):
        flash("Invalid work order file type.", "error")
        return redirect(url_for("index"))

    persisted_logo_image, persisted_logo_zip = get_persisted_logo_paths()
    has_new_logo = (logo_image and logo_image.filename) or (logo_zip and logo_zip.filename)

    if not has_new_logo and not persisted_logo_image and not persisted_logo_zip:
        flash("Please upload a logo image or a logo zip.", "error")
        return redirect(url_for("index"))

    run_id = uuid.uuid4().hex
    run_upload_dir = UPLOAD_DIR / run_id
    run_upload_dir.mkdir(parents=True, exist_ok=True)

    workorder_name = secure_filename(workorder.filename)
    workorder_path = run_upload_dir / workorder_name
    workorder.save(workorder_path)

    logo_image_path = None
    logo_zip_path = None

    try:
        logo_image_path, logo_zip_path = save_logo_inputs_to_persistent(logo_image, logo_zip)
    except ValueError as exc:
        flash(str(exc), "error")
        return redirect(url_for("index"))

    config = LabelRunConfig(
        output_dir=str(output_path),
        logo_image_path=str(logo_image_path) if logo_image_path else None,
        logo_zip_path=str(logo_zip_path) if logo_zip_path else None,
    )

    try:
        result = run_label_generation(
            str(workorder_path),
            config,
            job_number_override=job_number_override or None,
            job_name_override=job_name_override or None,
        )
    except Exception as exc:
        flash(f"Run failed: {exc}", "error")
        return redirect(url_for("index"))
    finally:
        shutil.rmtree(run_upload_dir, ignore_errors=True)

    output_files = [Path(p) for p in result["output_files"]]
    if not output_files:
        flash(
            "Run completed but no labels were generated. "
            f"Detected items={result.get('items_found', 0)}, "
            f"parts={result.get('parsed_parts_total', 0)}, "
            f"text_chars={result.get('text_chars', 0)}, "
            f"job={result.get('job_number', '').strip() or 'UNKNOWN'}.",
            "error",
        )
        return redirect(url_for("index"))

    _, cached_names = cache_output_files(output_files)
    flash(f"Labels generated and cached ({len(cached_names)} file(s)). Use Print or View Latest Labels.", "success")
    return redirect(url_for("index"))


@app.route("/manual", methods=["POST"])
def manual_generation():
    logo_image = request.files.get("manual_logo_image")
    logo_zip = request.files.get("manual_logo_zip")

    output_dir = (request.form.get("manual_output_dir") or "").strip() or str(DEFAULT_OUTPUT_DIR)
    output_path = Path(output_dir).expanduser().resolve()
    output_path.mkdir(parents=True, exist_ok=True)

    job_number = (request.form.get("manual_job_number") or "").strip()
    job_name = (request.form.get("manual_job_name") or "").strip()
    item_label = (request.form.get("manual_item_label") or "").strip()
    description = (request.form.get("manual_description") or "").strip()
    qty_raw = (request.form.get("manual_quantity") or "1").strip()
    label_size = (request.form.get("manual_label_size") or "auto").strip().lower()

    if not item_label:
        flash("Manual mode: Item label is required.", "error")
        return redirect(url_for("index"))

    try:
        quantity = int(qty_raw)
    except ValueError:
        flash("Manual mode: Quantity must be a number.", "error")
        return redirect(url_for("index"))

    if quantity < 1 or quantity > 500:
        flash("Manual mode: Quantity must be between 1 and 500.", "error")
        return redirect(url_for("index"))

    try:
        logo_image_path, logo_zip_path = save_logo_inputs_to_persistent(logo_image, logo_zip)
    except ValueError as exc:
        flash(str(exc), "error")
        return redirect(url_for("index"))

    if not logo_image_path and not logo_zip_path:
        flash("Manual mode: Upload a logo once or choose one now.", "error")
        return redirect(url_for("index"))

    config = LabelRunConfig(
        output_dir=str(output_path),
        logo_image_path=str(logo_image_path) if logo_image_path else None,
        logo_zip_path=str(logo_zip_path) if logo_zip_path else None,
    )

    try:
        output_files = generate_manual_labels(
            config=config,
            job_number=job_number or " ",
            job_name=job_name or "UNKNOWN JOB",
            item_label=item_label,
            description=description,
            quantity=quantity,
            label_size=label_size,
        )
    except Exception as exc:
        flash(f"Manual mode failed: {exc}", "error")
        return redirect(url_for("index"))

    _, cached_names = cache_output_files(output_files)
    flash(f"Manual labels generated and cached ({len(cached_names)} file(s)). Use Print or View Latest Labels.", "success")
    return redirect(url_for("index"))


@app.route("/labels/latest/<path:filename>", methods=["GET"])
def latest_label_file(filename):
    result_id, files = get_latest_cached_result()
    if not result_id:
        flash("No cached labels found. Generate labels before viewing.", "error")
        return redirect(url_for("index"))

    if filename not in files:
        abort(404)

    safe_name = Path(filename).name
    run_dir = RESULT_CACHE_DIR / result_id
    file_path = run_dir / safe_name
    if not file_path.exists():
        abort(404)
    return send_file(file_path, as_attachment=False)


@app.route("/labels/latest", methods=["GET"])
def view_latest_labels():
    result_id, files = get_latest_cached_result()
    if not result_id:
        flash("No cached labels found. Generate labels before viewing.", "error")
        return redirect(url_for("index"))
    return render_template("latest_labels.html", files=files)


@app.route("/labels/print", methods=["GET"])
def print_latest_labels():
    result_id, files = get_latest_cached_result()
    if not result_id:
        flash("No labels generated yet. Generate labels before printing.", "error")
        return redirect(url_for("index"))

    # For one file, open directly so browser PDF viewer print is one tap.
    if len(files) == 1:
        return redirect(url_for("latest_label_file", filename=files[0]))
    return redirect(url_for("view_latest_labels"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
