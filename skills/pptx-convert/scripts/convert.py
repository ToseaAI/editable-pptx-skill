#!/usr/bin/env python3
"""
Reslide conversion CLI — turn slide images or image-based PDFs into a fully
editable PowerPoint (.pptx) via the Reslide API by Tosea.

Zero third-party dependencies: standard library only (urllib, json, time,
mimetypes, argparse, os, sys, uuid).

Environment
-----------
  TOSEA_PPTX_API_KEY   (required)  Bearer API key. Create one at
                                   https://reslide.tosea.ai/keys
  TOSEA_PPTX_API_BASE  (optional)  API base URL.
                                   Default: https://reslideapi.tosea.ai

Usage
-----
  # Local PDF  -> editable .pptx (uploaded, then converted)
  python convert.py deck.pdf

  # Local slide images -> one editable .pptx (uploaded in order)
  python convert.py slide1.png slide2.png slide3.png

  # Remote PDF by URL
  python convert.py --pdf-url https://example.com/deck.pdf

  # Remote images by URL (one per slide, in order)
  python convert.py --image-urls https://ex.com/1.png https://ex.com/2.png

  # Options
  --out result.pptx     Output path (default: <input-name>.pptx / output.pptx)
  --title "Q3 deck"     Optional deck title stored on the task

Billing
-------
  1 page = 1 credit. Credits are reserved when the task is created and only
  settle when it succeeds. A failed or canceled task has zero net billable
  credits; its reservation is reconciled against the original credit pools
  subject to those pools' lifecycle rules.

API contract (envelope: {code, message, data}, code 0 == success)
-----------------------------------------------------------------
  POST /v1/uploads          multipart form field "file"    -> { upload_id, ... }
  POST /v1/tasks            { operation, input }            -> { task_id, ... }
  GET  /v1/tasks/{task_id}                                  -> { status, result, error, ... }

  operation "pdf_to_pptx"    input: { upload_id | pdf_url,  title? }
  operation "images_to_pptx" input: { upload_ids | image_urls, title? }

Exit codes: 0 on success, 1 on any error (message printed to stderr).
"""

import argparse
import hashlib
import json
import mimetypes
import os
import sys
import tempfile
import time
import uuid
from urllib import error, request

API_BASE = os.environ.get("TOSEA_PPTX_API_BASE", "https://reslideapi.tosea.ai").rstrip("/")
API_KEY = os.environ.get("TOSEA_PPTX_API_KEY")

POLL_INTERVAL = 3.0  # seconds between status polls
POLL_TIMEOUT = 15 * 60  # give up after 15 minutes
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp"}


# --------------------------------------------------------------------------- #
# Errors + small helpers
# --------------------------------------------------------------------------- #
class ApiError(Exception):
    """An error carrying the HTTP/envelope status and an optional Retry-After."""

    def __init__(self, status, message, retry_after=None):
        super().__init__(message)
        self.status = status
        self.message = message
        self.retry_after = retry_after


def log(msg):
    """Progress goes to stderr so stdout stays reserved for the result path."""
    print(msg, file=sys.stderr, flush=True)


def die(msg, code=1):
    log(f"error: {msg}")
    sys.exit(code)


def _auth_headers():
    return {"Authorization": f"Bearer {API_KEY}"}


def _parse_envelope(raw_bytes, status):
    """Return `data` for a code==0 envelope, else raise ApiError."""
    try:
        payload = json.loads(raw_bytes.decode("utf-8"))
    except Exception:
        raise ApiError(status, f"Non-JSON response from server (HTTP {status})") from None
    code = payload.get("code")
    if code == 0:
        return payload.get("data")
    # Envelope reported a failure — surface its code + message.
    raise ApiError(code or status, payload.get("message") or "Request failed")


def _send(req):
    """Perform an urllib request and unwrap the {code,message,data} envelope."""
    try:
        with request.urlopen(req) as resp:
            return _parse_envelope(resp.read(), getattr(resp, "status", 200))
    except error.HTTPError as exc:
        retry_after = exc.headers.get("Retry-After")
        ra = int(retry_after) if retry_after and str(retry_after).isdigit() else None
        # A well-formed error body is still an envelope; parse it for the message.
        try:
            _parse_envelope(exc.read(), exc.code)
        except ApiError as api_err:
            api_err.retry_after = ra
            raise
        raise ApiError(exc.code, f"HTTP {exc.code}", ra) from exc
    except error.URLError as exc:
        raise ApiError(0, f"Network error: {exc.reason}") from exc


def api_json(method, path, body=None, extra_headers=None):
    """JSON request against the API base, returns the envelope `data`."""
    headers = _auth_headers()
    if extra_headers:
        headers.update(extra_headers)
    data = None
    if body is not None:
        data = json.dumps(body).encode("utf-8")
        headers["Content-Type"] = "application/json"
    req = request.Request(f"{API_BASE}{path}", data=data, method=method, headers=headers)
    return _send(req)


# --------------------------------------------------------------------------- #
# API operations
# --------------------------------------------------------------------------- #
def upload_file(path):
    """POST a local file to /v1/uploads as multipart/form-data; return upload_id."""
    filename = os.path.basename(path)
    ctype = mimetypes.guess_type(path)[0] or "application/octet-stream"
    with open(path, "rb") as fh:
        content = fh.read()

    # Hand-build a multipart body with a unique boundary (no third-party deps).
    boundary = "----ToseaBoundary" + uuid.uuid4().hex
    preamble = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
        f"Content-Type: {ctype}\r\n\r\n"
    ).encode()
    epilogue = f"\r\n--{boundary}--\r\n".encode()
    body = preamble + content + epilogue

    headers = _auth_headers()
    headers["Content-Type"] = f"multipart/form-data; boundary={boundary}"
    req = request.Request(f"{API_BASE}/v1/uploads", data=body, method="POST", headers=headers)
    data = _send(req)
    upload_id = (data or {}).get("upload_id")
    if not upload_id:
        raise ApiError(0, f"Upload of {filename} returned no upload_id")
    return upload_id


def create_task(operation, input_obj, idempotency_key):
    """POST /v1/tasks; return the created task dict (with task_id, status)."""
    data = api_json(
        "POST",
        "/v1/tasks",
        {"operation": operation, "input": input_obj},
        {"Idempotency-Key": idempotency_key},
    )
    if not data or not data.get("task_id"):
        raise ApiError(0, "Task creation returned no task_id")
    return data


def poll_task(task_id):
    """Poll GET /v1/tasks/{id} every POLL_INTERVAL until terminal; honor 429."""
    deadline = time.time() + POLL_TIMEOUT
    last_status = None
    while True:
        try:
            task = api_json("GET", f"/v1/tasks/{task_id}")
        except ApiError as exc:
            if exc.status == 429:
                wait = exc.retry_after or POLL_INTERVAL
                log(f"rate limited; waiting {wait}s before retrying…")
                time.sleep(wait)
                continue
            raise

        status = task.get("status")
        if status != last_status:
            pct = task.get("progress")
            suffix = f" ({pct}%)" if isinstance(pct, (int, float)) else ""
            log(f"status: {status}{suffix}")
            last_status = status

        if status == "succeeded":
            return task
        if status in ("failed", "canceled"):
            raise ApiError(0, task.get("error") or f"Task {status}")
        if time.time() > deadline:
            raise ApiError(0, "Timed out waiting for the conversion to finish")
        time.sleep(POLL_INTERVAL)


def download(url, out_path, overwrite=False):
    """Atomically download a signed result URL without clobbering by default."""
    target = os.path.abspath(out_path)
    parent = os.path.dirname(target) or "."
    if os.path.exists(target) and not overwrite:
        raise ApiError(0, f"output already exists: {out_path} (use --force to replace it)")
    if not os.path.isdir(parent):
        raise ApiError(0, f"output directory does not exist: {parent}")

    temp_path = None
    req = request.Request(url)  # signed URL — no auth header required
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb",
            prefix=f".{os.path.basename(target)}.",
            suffix=".part",
            dir=parent,
            delete=False,
        ) as fh:
            temp_path = fh.name
            with request.urlopen(req) as resp:
                while chunk := resp.read(1024 * 1024):
                    fh.write(chunk)
            fh.flush()
            os.fsync(fh.fileno())

        if overwrite:
            os.replace(temp_path, target)
        else:
            # Hard-link publication is atomic and fails if another process
            # created the target after the preflight check.
            os.link(temp_path, target)
            os.unlink(temp_path)
        temp_path = None
    except FileExistsError as exc:
        raise ApiError(
            0, f"output already exists: {out_path} (use --force to replace it)"
        ) from exc
    except error.HTTPError as exc:
        raise ApiError(exc.code, f"Failed to download result (HTTP {exc.code})") from exc
    except error.URLError as exc:
        raise ApiError(0, f"Failed to download result: {exc.reason}") from exc
    except OSError as exc:
        raise ApiError(0, f"Failed to save result: {exc}") from exc
    finally:
        if temp_path:
            try:
                os.unlink(temp_path)
            except FileNotFoundError:
                pass


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def _is_image(path):
    return os.path.splitext(path)[1].lower() in IMAGE_EXTS


def _is_pdf(path):
    return os.path.splitext(path)[1].lower() == ".pdf"


def idempotency_key(args, operation):
    """Return an explicit key or a stable fingerprint of the logical input."""
    if args.idempotency_key is not None:
        key = args.idempotency_key.strip()
        if not key or "\r" in key or "\n" in key:
            die("--idempotency-key must be a non-empty single-line value")
        return key

    digest = hashlib.sha256()
    digest.update(b"reslide-cli-v1\0")
    digest.update(operation.encode())
    digest.update(b"\0")
    digest.update((args.title or "").encode())
    if args.pdf_url:
        digest.update(b"\0pdf-url\0" + args.pdf_url.encode())
    elif args.image_urls:
        for url in args.image_urls:
            digest.update(b"\0image-url\0" + url.encode())
    else:
        for path in args.inputs:
            digest.update(b"\0file\0")
            with open(path, "rb") as fh:
                while chunk := fh.read(1024 * 1024):
                    digest.update(chunk)
    return "reslide-cli-v1-" + digest.hexdigest()


def resolve_job(args):
    """Decide (operation, input_obj, default_out) from the parsed arguments."""
    title = args.title

    # 1. Remote PDF URL
    if args.pdf_url:
        inp = {"pdf_url": args.pdf_url}
        if title:
            inp["title"] = title
        return "pdf_to_pptx", inp, "output.pptx"

    # 2. Remote image URLs
    if args.image_urls:
        inp = {"image_urls": list(args.image_urls)}
        if title:
            inp["title"] = title
        return "images_to_pptx", inp, "output.pptx"

    # 3. Local files (positional)
    if args.inputs:
        for p in args.inputs:
            if not os.path.isfile(p):
                die(f"file not found: {p}")

        pdfs = [p for p in args.inputs if _is_pdf(p)]
        imgs = [p for p in args.inputs if _is_image(p)]

        # 3a. A single local PDF
        if len(args.inputs) == 1 and pdfs:
            path = args.inputs[0]
            log(f"uploading {os.path.basename(path)} …")
            upload_id = upload_file(path)
            inp = {"upload_id": upload_id}
            if title:
                inp["title"] = title
            default_out = os.path.splitext(os.path.basename(path))[0] + ".pptx"
            return "pdf_to_pptx", inp, default_out

        # 3b. One or more local images
        if imgs and not pdfs and len(imgs) == len(args.inputs):
            upload_ids = []
            for path in args.inputs:  # preserve slide order
                log(f"uploading {os.path.basename(path)} …")
                upload_ids.append(upload_file(path))
            inp = {"upload_ids": upload_ids}
            if title:
                inp["title"] = title
            default_out = os.path.splitext(os.path.basename(args.inputs[0]))[0] + ".pptx"
            return "images_to_pptx", inp, default_out

        die(
            "unsupported input mix. Pass a single .pdf, or one-or-more image "
            "files (png/jpg/…). Do not mix PDFs and images in one call."
        )

    die(
        "nothing to convert. Provide a local .pdf, image files, --pdf-url, "
        "or --image-urls. Run with -h for usage."
    )


def main():
    parser = argparse.ArgumentParser(
        prog="convert.py",
        description="Convert images or image-based PDFs into editable PPTX via Reslide.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "inputs",
        nargs="*",
        help="local PDF file, or one or more image files (in slide order)",
    )
    parser.add_argument("--pdf-url", help="convert a remote PDF by URL")
    parser.add_argument(
        "--image-urls",
        nargs="+",
        help="convert remote images by URL (one per slide, in order)",
    )
    parser.add_argument("--out", help="output .pptx path (default: derived from input)")
    parser.add_argument("--title", help="optional deck title stored on the task")
    parser.add_argument(
        "--idempotency-key",
        help="stable caller job ID (default: derived from input content/URLs and title)",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="atomically replace an existing output file",
    )
    args = parser.parse_args()

    if not API_KEY:
        die(
            "TOSEA_PPTX_API_KEY is not set. Create a key at "
            "https://reslide.tosea.ai/keys and export it, e.g.\n"
            "  export TOSEA_PPTX_API_KEY=tpx_live_…"
        )

    try:
        operation, input_obj, default_out = resolve_job(args)
        out_path = args.out or default_out
        if os.path.exists(out_path) and not args.force:
            die(f"output already exists: {out_path} (use --force to replace it)")
        idem = idempotency_key(args, operation)

        log(f"creating {operation} task on {API_BASE} …")
        task = create_task(operation, input_obj, idem)
        task_id = task["task_id"]
        log(f"task created: {task_id}")

        task = poll_task(task_id)
        result = task.get("result") or {}
        pptx_url = result.get("pptx_url")
        if not pptx_url:
            die("task succeeded but no result.pptx_url was returned")

        pages = task.get("page_count")
        if pages:
            log(f"converted {pages} page(s)")
        log(f"downloading result → {out_path} …")
        download(pptx_url, out_path, overwrite=args.force)
        log("done.")
        # The final saved path is the only thing on stdout, for easy capture.
        print(out_path)

    except ApiError as exc:
        if exc.status == 401:
            die(
                "401 Unauthorized — your API key is missing or invalid. "
                "Check TOSEA_PPTX_API_KEY (get one at https://reslide.tosea.ai/keys)."
            )
        if exc.status == 402:
            die(
                "402 Payment Required — not enough credits for this conversion. "
                "Top up at https://reslide.tosea.ai/billing (1 page = 1 credit)."
            )
        if exc.status == 429:
            die("429 Too Many Requests — rate limited. Wait and retry shortly.")
        die(f"conversion failed: {exc.message}")
    except KeyboardInterrupt:
        die("interrupted", code=130)


if __name__ == "__main__":
    main()
