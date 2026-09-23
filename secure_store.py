from __future__ import annotations

import ctypes
import os
from ctypes import wintypes
from pathlib import Path


APP_DIR = Path(os.getenv("LOCALAPPDATA", Path.home())) / "STC" / "HotelPromoConverter"
AZURE_MAPS_KEY_FILE = APP_DIR / "azure_maps_subscription_key.bin"
ENTROPY = b"STC HotelPromoConverter Azure Maps key v1"


class SecureStoreError(RuntimeError):
    """Raised when Windows cannot protect or unprotect a saved secret."""


class DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_char)),
    ]


crypt32 = ctypes.windll.crypt32
kernel32 = ctypes.windll.kernel32


def _blob_from_bytes(data: bytes) -> DATA_BLOB:
    buffer = ctypes.create_string_buffer(data)
    blob = DATA_BLOB(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_char)))
    blob._buffer = buffer
    return blob


def _bytes_from_blob(blob: DATA_BLOB) -> bytes:
    try:
        return ctypes.string_at(blob.pbData, blob.cbData)
    finally:
        kernel32.LocalFree(blob.pbData)


def _protect(secret: str) -> bytes:
    data_blob = _blob_from_bytes(secret.encode("utf-8"))
    entropy_blob = _blob_from_bytes(ENTROPY)
    output_blob = DATA_BLOB()

    success = crypt32.CryptProtectData(
        ctypes.byref(data_blob),
        "Azure Maps subscription key",
        ctypes.byref(entropy_blob),
        None,
        None,
        0,
        ctypes.byref(output_blob),
    )

    if not success:
        raise SecureStoreError("Could not encrypt Azure Maps key with Windows DPAPI.")

    return _bytes_from_blob(output_blob)


def _unprotect(encrypted: bytes) -> str:
    data_blob = _blob_from_bytes(encrypted)
    entropy_blob = _blob_from_bytes(ENTROPY)
    output_blob = DATA_BLOB()

    success = crypt32.CryptUnprotectData(
        ctypes.byref(data_blob),
        None,
        ctypes.byref(entropy_blob),
        None,
        None,
        0,
        ctypes.byref(output_blob),
    )

    if not success:
        raise SecureStoreError("Could not decrypt saved Azure Maps key with Windows DPAPI.")

    return _bytes_from_blob(output_blob).decode("utf-8")


def save_azure_maps_key(key: str) -> None:
    cleaned_key = key.strip()

    if not cleaned_key:
        delete_azure_maps_key()
        return

    APP_DIR.mkdir(parents=True, exist_ok=True)
    AZURE_MAPS_KEY_FILE.write_bytes(_protect(cleaned_key))


def load_azure_maps_key() -> str:
    if not AZURE_MAPS_KEY_FILE.exists():
        return ""

    return _unprotect(AZURE_MAPS_KEY_FILE.read_bytes())


def delete_azure_maps_key() -> None:
    AZURE_MAPS_KEY_FILE.unlink(missing_ok=True)


def has_saved_azure_maps_key() -> bool:
    return AZURE_MAPS_KEY_FILE.exists()
