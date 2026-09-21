from __future__ import annotations

"""Durable Vercel Blob state for the SCM dashboard.

v2.46.6 deliberately does not depend on a particular Vercel Python SDK release.
It speaks to the documented Blob HTTP API directly so both authentication models
work reliably:

* current Vercel OIDC: ``x-vercel-oidc-token`` + ``BLOB_STORE_ID``;
* legacy/static Blob auth: ``BLOB_READ_WRITE_TOKEN``.

Vercel supplies the OIDC token per function request, not at Python module import
time.  A ContextVar therefore holds the token for the active Flask request.  This
also avoids leaking one request's token into another request when a warm function
handles concurrent traffic.
"""

from contextvars import ContextVar
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
import json
import os
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid


@dataclass(frozen=True)
class CloudStatus:
    runtime: str
    enabled: bool
    provider: str
    detail: str
    auth_mode: str = "none"


class VercelBlobState:
    API_VERSION = "12"

    def __init__(self) -> None:
        self.is_vercel = bool(os.environ.get("VERCEL") or os.environ.get("VERCEL_ENV"))
        self.read_write_token = str(os.environ.get("BLOB_READ_WRITE_TOKEN") or "").strip()
        self.env_oidc_token = str(os.environ.get("VERCEL_OIDC_TOKEN") or "").strip()
        self.store_id = self._normalize_store_id(str(os.environ.get("BLOB_STORE_ID") or "").strip())
        self.prefix = str(os.environ.get("SCM_BLOB_PREFIX") or "scm-idp-dashboard").strip().strip("/")
        self.allow_ephemeral = str(os.environ.get("SCM_ALLOW_EPHEMERAL_VERCEL") or "0").strip() == "1"
        self.api_url = str(os.environ.get("VERCEL_BLOB_API_URL") or "https://vercel.com/api/blob").rstrip("/")
        self._request_oidc: ContextVar[str] = ContextVar("scm_vercel_oidc_token", default="")
        self.last_error = ""

    @staticmethod
    def _normalize_store_id(value: str) -> str:
        value = (value or "").strip()
        return value[6:] if value.startswith("store_") else value

    @staticmethod
    def _store_id_from_rw_token(token: str) -> str:
        # Vercel read-write tokens are shaped like
        # vercel_blob_rw_<store-id>_<secret>. Match the official SDK parser.
        parts = (token or "").split("_")
        if len(parts) >= 5 and parts[0:3] == ["vercel", "blob", "rw"]:
            return parts[3]
        return ""

    def bind_request_oidc(self, token: str | None) -> None:
        """Bind the OIDC token for the current request context.

        Vercel puts this token in ``x-vercel-oidc-token`` at function runtime.
        ContextVar makes this safe for concurrent warm-function requests.
        """
        self._request_oidc.set(str(token or "").strip())

    @property
    def request_oidc_present(self) -> bool:
        return bool(self._request_oidc.get().strip())

    def _auth(self) -> tuple[str, str, str]:
        request_oidc = self._request_oidc.get().strip()
        if request_oidc and self.store_id:
            return request_oidc, self.store_id, "oidc-request"

        # Static read-write credentials remain supported for older Blob project
        # connections and also allow module-level cold-start hydration.
        if self.read_write_token:
            store_id = self.store_id or self._store_id_from_rw_token(self.read_write_token)
            if store_id:
                return self.read_write_token, self._normalize_store_id(store_id), "read-write-token"

        # Useful for local development after ``vercel env pull``. In production
        # the runtime OIDC token normally arrives on the request header instead.
        if self.env_oidc_token and self.store_id:
            return self.env_oidc_token, self.store_id, "oidc-env"

        return "", "", "none"

    @property
    def auth_mode(self) -> str:
        return self._auth()[2]

    @property
    def enabled(self) -> bool:
        token, store_id, _mode = self._auth()
        return bool(token and store_id)

    @property
    def durable_required(self) -> bool:
        return self.is_vercel and not self.allow_ephemeral

    def status(self) -> CloudStatus:
        if not self.is_vercel:
            return CloudStatus("local", False, "local-filesystem", "Local writable state directory", "local")
        if self.enabled:
            mode = self.auth_mode
            detail = "Private Blob persistence active"
            if mode.startswith("oidc"):
                detail += " (Vercel OIDC)"
            else:
                detail += " (read-write token)"
            return CloudStatus("vercel", True, "vercel-blob-private", detail, mode)
        if self.allow_ephemeral:
            return CloudStatus("vercel", False, "ephemeral-/tmp", "Ephemeral Vercel mode explicitly enabled", "ephemeral")
        if self.store_id and not (self.request_oidc_present or self.env_oidc_token or self.read_write_token):
            return CloudStatus(
                "vercel",
                False,
                "vercel-blob-awaiting-oidc",
                "Blob store is connected, but this request did not include a Vercel OIDC token.",
                "none",
            )
        if (self.request_oidc_present or self.env_oidc_token) and not self.store_id and not self.read_write_token:
            return CloudStatus(
                "vercel",
                False,
                "vercel-blob-missing-store-id",
                "Vercel OIDC is available but BLOB_STORE_ID is missing. Reconnect the Blob store to this project and redeploy.",
                "oidc-missing-store-id",
            )
        return CloudStatus(
            "vercel",
            False,
            "not-configured",
            "Connect a Private Vercel Blob store to this project. OIDC and legacy read-write tokens are both supported.",
            "none",
        )

    def _key(self, suffix: str) -> str:
        suffix = suffix.strip().lstrip("/")
        return f"{self.prefix}/{suffix}" if self.prefix else suffix

    @staticmethod
    def _error_message(exc: urllib.error.HTTPError) -> str:
        try:
            raw = exc.read().decode("utf-8", errors="replace")
        except Exception:
            raw = ""
        message = raw.strip()
        if raw:
            try:
                data = json.loads(raw)
                if isinstance(data, dict):
                    err = data.get("error")
                    if isinstance(err, dict):
                        message = str(err.get("message") or err.get("code") or message)
                    elif err:
                        message = str(err)
            except Exception:
                pass
        if len(message) > 500:
            message = message[:500] + "…"
        return f"HTTP {exc.code}: {message or exc.reason}"

    def _api_headers(self, *, content_type: str | None = None) -> dict[str, str]:
        token, store_id, _mode = self._auth()
        if not token or not store_id:
            raise RuntimeError(self.status().detail)
        request_id = f"{store_id}:{int(time.time() * 1000)}:{uuid.uuid4().hex[:12]}"
        headers = {
            "Authorization": f"Bearer {token}",
            "x-vercel-blob-store-id": store_id,
            "x-api-version": self.API_VERSION,
            "x-api-blob-request-id": request_id,
            "x-api-blob-request-attempt": "0",
        }
        if content_type:
            headers["Content-Type"] = content_type
        return headers

    def _open_with_retry(self, req: urllib.request.Request, *, timeout: int = 45) -> bytes:
        last_exc: Exception | None = None
        for attempt in range(3):
            try:
                with urllib.request.urlopen(req, timeout=timeout) as response:  # nosec - Vercel-owned endpoint/blob host
                    return response.read()
            except urllib.error.HTTPError as exc:
                last_exc = exc
                # Do not retry authentication/configuration/client errors. Retry
                # rate limits and transient Blob/server errors briefly.
                if exc.code not in {408, 425, 429, 500, 502, 503, 504} or attempt == 2:
                    message = self._error_message(exc)
                    self.last_error = message
                    raise RuntimeError(f"Vercel Blob request failed ({message})") from exc
                retry_after = exc.headers.get("retry-after") if exc.headers else None
                try:
                    delay = min(3.0, max(0.25, float(retry_after))) if retry_after else 0.4 * (2**attempt)
                except (TypeError, ValueError):
                    delay = 0.4 * (2**attempt)
                time.sleep(delay)
            except (urllib.error.URLError, TimeoutError, OSError) as exc:
                last_exc = exc
                if attempt == 2:
                    self.last_error = str(exc)
                    raise RuntimeError(f"Vercel Blob network request failed: {exc}") from exc
                time.sleep(0.4 * (2**attempt))
        raise RuntimeError(f"Vercel Blob request failed: {last_exc}")

    def _put_bytes(self, key: str, payload: bytes, *, content_type: str | None = None) -> None:
        if not self.enabled:
            if self.durable_required:
                raise RuntimeError(self.status().detail)
            return
        params = urllib.parse.urlencode({"pathname": self._key(key)})
        headers = self._api_headers(content_type="application/octet-stream")
        # Match current @vercel/blob server PUT semantics.
        headers["x-vercel-blob-access"] = "private"
        headers["x-allow-overwrite"] = "1"
        if content_type:
            headers["x-content-type"] = content_type
        req = urllib.request.Request(
            f"{self.api_url}/?{params}",
            data=bytes(payload),
            headers=headers,
            method="PUT",
        )
        self._open_with_retry(req)

    def put_file(self, key: str, source: Path, *, content_type: str | None = None) -> None:
        self._put_bytes(key, source.read_bytes(), content_type=content_type)

    def put_json(self, key: str, payload: Any) -> None:
        body = json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True).encode("utf-8")
        self._put_bytes(key, body, content_type="application/json")

    def get_bytes(self, key: str) -> bytes | None:
        if not self.enabled:
            return None
        token, store_id, _mode = self._auth()
        pathname = urllib.parse.quote(self._key(key), safe="/-._~")
        url = f"https://{store_id}.private.blob.vercel-storage.com/{pathname}?cache=0"
        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"}, method="GET")
        try:
            return self._open_with_retry(req)
        except RuntimeError as exc:
            cause = exc.__cause__
            if isinstance(cause, urllib.error.HTTPError) and cause.code == 404:
                return None
            if "HTTP 404" in str(exc):
                return None
            raise

    def get_json(self, key: str) -> dict[str, Any] | None:
        raw = self.get_bytes(key)
        if raw is None:
            return None
        try:
            value = json.loads(raw.decode("utf-8"))
            return value if isinstance(value, dict) else None
        except Exception as exc:
            raise RuntimeError(f"Invalid cloud state JSON at {self._key(key)}: {exc}") from exc

    def delete(self, key: str) -> None:
        if not self.enabled:
            if self.durable_required:
                raise RuntimeError(self.status().detail)
            return
        payload = json.dumps({"urls": [self._key(key)]}).encode("utf-8")
        headers = self._api_headers(content_type="application/json")
        req = urllib.request.Request(
            f"{self.api_url}/delete",
            data=payload,
            headers=headers,
            method="POST",
        )
        try:
            self._open_with_retry(req)
        except RuntimeError as exc:
            if "HTTP 404" not in str(exc):
                raise

    def probe(self) -> dict[str, Any]:
        """Verify this request can write, consistently read and delete a private blob."""
        if not self.enabled:
            raise RuntimeError(self.status().detail)
        key = f"diagnostics/probe-{uuid.uuid4().hex}.json"
        payload = {"ok": True, "ts": datetime.now(timezone.utc).isoformat(), "auth": self.auth_mode}
        try:
            self.put_json(key, payload)
            read_back = self.get_json(key)
            if not read_back or read_back.get("ok") is not True:
                raise RuntimeError("Vercel Blob probe could not read back the object it just wrote.")
            return {"ok": True, "auth_mode": self.auth_mode, "store_id": self._auth()[1]}
        finally:
            try:
                self.delete(key)
            except Exception:
                # A successful put/read proves persistence. Probe cleanup is not
                # allowed to make a valid import fail.
                pass

    def download_to(self, key: str, destination: Path) -> bool:
        raw = self.get_bytes(key)
        if raw is None:
            return False
        destination.parent.mkdir(parents=True, exist_ok=True)
        fd, temp_name = tempfile.mkstemp(prefix=f".{destination.name}.", dir=destination.parent)
        os.close(fd)
        temp = Path(temp_name)
        try:
            temp.write_bytes(raw)
            os.replace(temp, destination)
        finally:
            temp.unlink(missing_ok=True)
        return True

    # ---------- Core workbook + Aging database ----------
    @property
    def core_manifest_key(self) -> str:
        return "core/current.json"

    def hydrate_core(self, state_root: Path) -> dict[str, Any] | None:
        """Restore the latest durable core revision into the disposable runtime."""
        if not self.enabled:
            return None
        manifest = self.get_json(self.core_manifest_key)
        if not manifest:
            return None

        active = state_root / "active_import.xlsx"
        aging = state_root / "aging" / "aging.db"
        marker = state_root / ".scm_no_data"

        if bool(manifest.get("empty")):
            active.unlink(missing_ok=True)
            aging.unlink(missing_ok=True)
            Path(str(aging) + "-wal").unlink(missing_ok=True)
            Path(str(aging) + "-shm").unlink(missing_ok=True)
            marker.parent.mkdir(parents=True, exist_ok=True)
            marker.write_text(str(manifest.get("published_at") or "cloud-empty"), encoding="utf-8")
            return manifest

        workbook_key = str(manifest.get("workbook") or "").strip()
        aging_key = str(manifest.get("aging_db") or "").strip()
        if not workbook_key or not aging_key:
            raise RuntimeError("Cloud core manifest is incomplete.")

        if not self.download_to(workbook_key, active):
            raise RuntimeError("Cloud core workbook referenced by the manifest was not found.")

        # The Aging module can initialize a seed DB before an OIDC token is
        # available on a Vercel cold start. Remove its WAL sidecars before
        # replacing the DB so stale pages can never be replayed onto cloud data.
        aging.parent.mkdir(parents=True, exist_ok=True)
        Path(str(aging) + "-wal").unlink(missing_ok=True)
        Path(str(aging) + "-shm").unlink(missing_ok=True)
        if not self.download_to(aging_key, aging):
            raise RuntimeError("Cloud Aging database referenced by the manifest was not found.")
        Path(str(aging) + "-wal").unlink(missing_ok=True)
        Path(str(aging) + "-shm").unlink(missing_ok=True)
        marker.unlink(missing_ok=True)
        return manifest

    def publish_core(self, workbook: Path, aging_db: Path, revision: str) -> dict[str, Any]:
        """Publish both core artifacts, then atomically move the manifest pointer."""
        if not self.enabled:
            if self.durable_required:
                raise RuntimeError(self.status().detail)
            return {"revision": revision, "empty": False, "storage": "ephemeral"}

        safe_revision = "".join(ch if ch.isalnum() or ch in "-_" else "-" for ch in revision)[:96]
        workbook_key = f"core/revisions/{safe_revision}/active_import.xlsx"
        aging_key = f"core/revisions/{safe_revision}/aging.db"
        self.put_file(
            workbook_key,
            workbook,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.put_file(aging_key, aging_db, content_type="application/x-sqlite3")
        manifest = {
            "schema": 2,
            "revision": safe_revision,
            "empty": False,
            "workbook": workbook_key,
            "aging_db": aging_key,
            "published_at": datetime.now(timezone.utc).isoformat(),
        }
        # The pointer is written last. Any earlier failure leaves the previous
        # live revision untouched.
        self.put_json(self.core_manifest_key, manifest)
        return manifest

    def publish_empty_core(self, revision: str) -> dict[str, Any]:
        manifest = {
            "schema": 2,
            "revision": revision,
            "empty": True,
            "published_at": datetime.now(timezone.utc).isoformat(),
        }
        if self.enabled:
            self.put_json(self.core_manifest_key, manifest)
        elif self.durable_required:
            raise RuntimeError(self.status().detail)
        return manifest

    # ---------- Small mutable JSON state ----------
    def hydrate_named_files(self, state_root: Path, relative_paths: Iterable[str]) -> None:
        if not self.enabled:
            return
        for rel in relative_paths:
            rel = rel.replace("\\", "/").lstrip("/")
            self.download_to(f"state/{rel}", state_root / rel)

    def persist_named_file(self, state_root: Path, path: Path, *, payload: bytes | None = None) -> None:
        try:
            rel = path.relative_to(state_root).as_posix()
        except ValueError as exc:
            raise ValueError(f"State file is outside SCM_DATA_DIR: {path}") from exc
        body = path.read_bytes() if payload is None else payload
        content_type = "application/json" if path.suffix.lower() == ".json" else None
        self._put_bytes(f"state/{rel}", body, content_type=content_type)

    def delete_named_file(self, state_root: Path, path: Path) -> None:
        try:
            rel = path.relative_to(state_root).as_posix()
        except ValueError as exc:
            raise ValueError(f"State file is outside SCM_DATA_DIR: {path}") from exc
        self.delete(f"state/{rel}")
