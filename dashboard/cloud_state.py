from __future__ import annotations

"""Durable state adapter for serverless deployments.

The SCM dashboard was originally designed around a writable local ``uploads``
directory.  Vercel Functions are intentionally stateless, so mutable state must
be kept in an external backing service.  This module uses a *private* Vercel
Blob store as the durable source of truth while keeping a disposable working
copy under ``/tmp`` for openpyxl/sqlite processing.

The core workbook + Aging database are published as an immutable revision and
then activated by replacing a tiny manifest.  That ordering prevents a partial
upload from becoming the current dashboard state.
"""

from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
import json
import os
import tempfile
import urllib.request


@dataclass(frozen=True)
class CloudStatus:
    runtime: str
    enabled: bool
    provider: str
    detail: str


class VercelBlobState:
    def __init__(self) -> None:
        self.is_vercel = bool(os.environ.get("VERCEL") or os.environ.get("VERCEL_ENV"))
        self.token = str(os.environ.get("BLOB_READ_WRITE_TOKEN") or "").strip()
        self.prefix = str(os.environ.get("SCM_BLOB_PREFIX") or "scm-idp-dashboard").strip().strip("/")
        self.allow_ephemeral = str(os.environ.get("SCM_ALLOW_EPHEMERAL_VERCEL") or "0").strip() == "1"
        self._client: Any | None = None
        self._sdk_error = ""

        if self.token:
            try:
                from vercel.blob import BlobClient  # type: ignore

                self._client = BlobClient(token=self.token)
            except TypeError:
                # SDK versions that read the token from the environment do not
                # accept it in the constructor.
                try:
                    from vercel.blob import BlobClient  # type: ignore

                    self._client = BlobClient()
                except Exception as exc:  # pragma: no cover - runtime dependency
                    self._sdk_error = str(exc)
            except Exception as exc:  # pragma: no cover - runtime dependency
                self._sdk_error = str(exc)

    @property
    def enabled(self) -> bool:
        return bool(self.token and self._client is not None)

    @property
    def durable_required(self) -> bool:
        return self.is_vercel and not self.allow_ephemeral

    def status(self) -> CloudStatus:
        if not self.is_vercel:
            return CloudStatus("local", False, "local-filesystem", "Local writable state directory")
        if self.enabled:
            return CloudStatus("vercel", True, "vercel-blob-private", "Private Blob persistence active")
        if self.token and self._sdk_error:
            return CloudStatus("vercel", False, "vercel-blob", f"Blob SDK unavailable: {self._sdk_error}")
        if self.allow_ephemeral:
            return CloudStatus("vercel", False, "ephemeral-/tmp", "Ephemeral Vercel mode explicitly enabled")
        return CloudStatus("vercel", False, "not-configured", "Connect a Private Vercel Blob store to this project")

    def _key(self, suffix: str) -> str:
        suffix = suffix.strip().lstrip("/")
        return f"{self.prefix}/{suffix}" if self.prefix else suffix

    def _put_bytes(self, key: str, payload: bytes, *, content_type: str | None = None) -> None:
        if not self.enabled:
            if self.durable_required:
                raise RuntimeError("Private Vercel Blob storage is not configured for this deployment.")
            return
        kwargs: dict[str, Any] = {
            "access": "private",
            "overwrite": True,
        }
        if content_type:
            kwargs["content_type"] = content_type
        # token is supplied explicitly as well as through the environment for
        # compatibility across Python SDK releases.
        kwargs["token"] = self.token
        self._client.put(self._key(key), payload, **kwargs)

    def put_file(self, key: str, source: Path, *, content_type: str | None = None) -> None:
        self._put_bytes(key, source.read_bytes(), content_type=content_type)

    def put_json(self, key: str, payload: Any) -> None:
        body = json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True).encode("utf-8")
        self._put_bytes(key, body, content_type="application/json")

    def delete(self, key: str) -> None:
        if not self.enabled:
            if self.durable_required:
                raise RuntimeError("Private Vercel Blob storage is not configured for this deployment.")
            return
        pathname = self._key(key)
        # The current Python SDK exposes delete on BlobClient.  Some earlier
        # versions exposed it as ``del_``; support both without weakening the
        # production dependency contract.
        deleter = getattr(self._client, "delete", None) or getattr(self._client, "del_", None)
        if deleter is None:
            try:
                from vercel.blob import delete as module_delete  # type: ignore

                module_delete(pathname, token=self.token)
                return
            except Exception as exc:  # pragma: no cover - runtime dependency
                raise RuntimeError(f"Vercel Blob delete is unavailable: {exc}") from exc
        try:
            deleter(pathname, token=self.token)
        except TypeError:
            try:
                deleter(pathname)
            except Exception as exc:
                if "not found" not in str(exc).lower() and "404" not in str(exc):
                    raise
        except Exception as exc:
            if "not found" not in str(exc).lower() and "404" not in str(exc):
                raise

    def _stream_to_bytes(self, result: Any) -> bytes:
        if result is None:
            raise FileNotFoundError
        status = int(getattr(result, "status_code", 200) or 200)
        if status == 404:
            raise FileNotFoundError
        if status != 200:
            raise RuntimeError(f"Blob read returned HTTP {status}.")
        stream = getattr(result, "stream", None)
        if stream is None:
            return b""
        if isinstance(stream, (bytes, bytearray, memoryview)):
            return bytes(stream)
        if hasattr(stream, "read"):
            return stream.read()
        if hasattr(stream, "__aiter__"):
            # Some SDK revisions expose an async stream even when metadata was
            # fetched from the synchronous client. Fetching the private blob URL
            # directly with the same bearer token keeps cold-start hydration sync.
            blob = getattr(result, "blob", None)
            url = str(getattr(blob, "url", "") or "")
            if url:
                req = urllib.request.Request(url, headers={"Authorization": f"Bearer {self.token}"})
                with urllib.request.urlopen(req, timeout=45) as response:  # nosec - trusted Vercel Blob URL
                    return response.read()
        try:
            return b"".join(bytes(chunk) for chunk in stream)
        except TypeError as exc:
            raise RuntimeError("Unsupported Vercel Blob response stream.") from exc

    def get_bytes(self, key: str) -> bytes | None:
        if not self.enabled:
            return None
        try:
            # Newer Blob SDKs support cache bypass for consistent reads after an
            # overwrite. Fall back cleanly for older Python SDK releases.
            result = self._client.get(self._key(key), access="private", token=self.token, use_cache=False)
        except TypeError:
            try:
                result = self._client.get(self._key(key), access="private", token=self.token)
            except TypeError:
                result = self._client.get(self._key(key), access="private")
        except Exception as exc:
            # Avoid importing SDK-specific exception classes during local tests.
            if "not found" in str(exc).lower() or "404" in str(exc):
                return None
            raise
        if result is None:
            return None
        try:
            return self._stream_to_bytes(result)
        except FileNotFoundError:
            return None

    def get_json(self, key: str) -> dict[str, Any] | None:
        raw = self.get_bytes(key)
        if raw is None:
            return None
        try:
            value = json.loads(raw.decode("utf-8"))
            return value if isinstance(value, dict) else None
        except Exception as exc:
            raise RuntimeError(f"Invalid cloud state JSON at {self._key(key)}: {exc}") from exc

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
            marker.parent.mkdir(parents=True, exist_ok=True)
            marker.write_text(str(manifest.get("published_at") or "cloud-empty"), encoding="utf-8")
            return manifest

        workbook_key = str(manifest.get("workbook") or "").strip()
        aging_key = str(manifest.get("aging_db") or "").strip()
        if not workbook_key or not aging_key:
            raise RuntimeError("Cloud core manifest is incomplete.")

        if not self.download_to(workbook_key, active):
            raise RuntimeError("Cloud core workbook referenced by the manifest was not found.")
        if not self.download_to(aging_key, aging):
            raise RuntimeError("Cloud Aging database referenced by the manifest was not found.")
        marker.unlink(missing_ok=True)
        return manifest

    def publish_core(self, workbook: Path, aging_db: Path, revision: str) -> dict[str, Any]:
        """Publish both core artifacts, then atomically move the manifest pointer."""
        if not self.enabled:
            if self.durable_required:
                raise RuntimeError("Private Vercel Blob storage is not configured for this deployment.")
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
            "schema": 1,
            "revision": safe_revision,
            "empty": False,
            "workbook": workbook_key,
            "aging_db": aging_key,
            "published_at": datetime.now(timezone.utc).isoformat(),
        }
        # This tiny pointer is replaced last.  Any failure before this point leaves
        # the previous manifest (and therefore the previous live revision) intact.
        self.put_json(self.core_manifest_key, manifest)
        return manifest

    def publish_empty_core(self, revision: str) -> dict[str, Any]:
        manifest = {
            "schema": 1,
            "revision": revision,
            "empty": True,
            "published_at": datetime.now(timezone.utc).isoformat(),
        }
        if self.enabled:
            self.put_json(self.core_manifest_key, manifest)
        elif self.durable_required:
            raise RuntimeError("Private Vercel Blob storage is not configured for this deployment.")
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
