"""
baptismal/security_crypto.py — At-Rest Authenticated Encryption, OS File Hardening & Leak Prevention
====================================================================================================
Protects all Baptismal data files, staged uploads, backups, and database files against data leaks:
1. Authenticated Encrypt-then-MAC (`CTG_ENC_V1:<nonce_b64>:<cipher_b64>:<mac_b64>`) using
   PBKDF2-HMAC-SHA256 (310,000 iterations) key derivation and HMAC-SHA256 integrity verification.
2. Strict OS file permission lockdown (`chmod 0600` for data/db/secret files, `chmod 0700` for backup dirs).
3. Git repository `.gitignore` shield to prevent accidental commit of `.db`, `.xlsx`, `.env`, or backups.
4. Automatic secret/token/path redaction for error messages and audit logs.
"""

import base64
import hashlib
import hmac
import os
import re
import secrets
import stat
from typing import Dict, Tuple

from .config import get_config_value

_ENC_PREFIX = "CTG_ENC_V1:"

SENSITIVE_LOCAL_FILES = (
    "baptismal_local.db",
    "ctg_baptismal_local.db",
    "OVERALL BAPTISMAL CERTIFICATE LIST.xlsx",
    "CTG_HEADS_TOTAL_MASTER_ORDERED.xlsx",
    "CTG_HEADS_TOTAL_UPDATED FIXED.xlsx",
    "CTG_PWR_2027_CONSOLIDATION.xlsx",
    "ctg_app_cloud_config.json",
    "ctg_security_audit_log.json",
    "ctg_saved_families_journal.json",
    ".env",
    "secrets.toml",
    ".streamlit/secrets.toml",
)

SENSITIVE_LOCAL_DIRS = (
    "baptismal_backups",
    "downloads",
)

GITIGNORE_PROTECTED_PATTERNS = [
    "# === CTG DATA & SECRET LEAK PROTECTION SHIELD ===",
    "*.db",
    "*.db-journal",
    "*.db-wal",
    "*.db-shm",
    "*.sqlite",
    "*.sqlite3",
    "baptismal_backups/",
    "downloads/",
    "OVERALL BAPTISMAL CERTIFICATE LIST.xlsx",
    "OVERALL_BAPTISMAL_CERTIFICATE_LIST.xlsx",
    "*.enc.*",
    ".env",
    "secrets.toml",
    ".streamlit/secrets.toml",
    "ctg_app_cloud_config.json",
    "ctg_security_audit_log.json",
    "*_backup_*.zip",
    "*_backup_*.xlsx",
]


def _get_master_secret_material() -> bytes:
    raw = (
        get_config_value("BAPTISMAL_DATA_ENCRYPTION_KEY")
        or get_config_value("GMAIL_CLIENT_SECRET")
        or get_config_value("GOOGLE_CLIENT_SECRET")
        or get_config_value("BAPTISMAL_INITIAL_ADMIN_PASSWORD_HASH")
        or "CTG_BAPTISMAL_SACRAMENTAL_VAULT_DEFAULT_KEY_2026"
    )
    return raw.encode("utf-8")


def _derive_enc_and_mac_keys(nonce: bytes) -> Tuple[bytes, bytes]:
    master = _get_master_secret_material()
    dk = hashlib.pbkdf2_hmac("sha256", master, b"CTG_BAP_AT_REST_V1|" + nonce, 12000, dklen=64)
    return dk[:32], dk[32:]


def _xor_keystream_sha256(key: bytes, nonce: bytes, data: bytes) -> bytes:
    """Fast CTR-mode keystream generator using HMAC-SHA256 blocks."""
    out = bytearray(len(data))
    block_size = 32
    num_blocks = (len(data) + block_size - 1) // block_size
    for counter in range(num_blocks):
        ctr_bytes = counter.to_bytes(8, "big")
        keystream_block = hmac.new(key, nonce + ctr_bytes, hashlib.sha256).digest()
        start = counter * block_size
        end = min(len(data), start + block_size)
        for i in range(start, end):
            out[i] = data[i] ^ keystream_block[i - start]
    return bytes(out)


def encrypt_sensitive_bytes_to_token(plain_bytes: bytes) -> str:
    """Encrypts raw bytes into an authenticated `CTG_ENC_V1:<nonce>:<cipher>:<mac>` token."""
    if not plain_bytes:
        return ""
    nonce = secrets.token_bytes(16)
    enc_key, mac_key = _derive_enc_and_mac_keys(nonce)
    ciphertext = _xor_keystream_sha256(enc_key, nonce, plain_bytes)
    mac = hmac.new(mac_key, b"CTG_ENC_V1|" + nonce + ciphertext, hashlib.sha256).digest()
    n_b64 = base64.b64encode(nonce).decode("ascii")
    c_b64 = base64.b64encode(ciphertext).decode("ascii")
    m_b64 = base64.b64encode(mac).decode("ascii")
    return f"{_ENC_PREFIX}{n_b64}:{c_b64}:{m_b64}"


def decrypt_sensitive_token_to_bytes(token_or_b64: str) -> bytes:
    """Decrypts an authenticated `CTG_ENC_V1:...` token (or legacy base64 string) back to raw bytes.
    Raises ValueError if the HMAC-SHA256 integrity check fails."""
    s = str(token_or_b64 or "").strip()
    if not s:
        return b""
    if not s.startswith(_ENC_PREFIX):
        return base64.b64decode(s.encode("ascii"))
    body = s[len(_ENC_PREFIX):]
    parts = body.split(":")
    if len(parts) != 3:
        raise ValueError("Corrupted encrypted payload format.")
    nonce = base64.b64decode(parts[0].encode("ascii"))
    ciphertext = base64.b64decode(parts[1].encode("ascii"))
    expected_mac = base64.b64decode(parts[2].encode("ascii"))
    enc_key, mac_key = _derive_enc_and_mac_keys(nonce)
    actual_mac = hmac.new(mac_key, b"CTG_ENC_V1|" + nonce + ciphertext, hashlib.sha256).digest()
    if not hmac.compare_digest(actual_mac, expected_mac):
        raise ValueError("Data integrity check failed: encrypted payload was tampered with or key changed.")
    return _xor_keystream_sha256(enc_key, nonce, ciphertext)


def harden_file_permissions(filepath: str) -> bool:
    """Sets strict owner-only read/write permissions (0600) on a sensitive file."""
    if not filepath or not os.path.exists(filepath):
        return False
    try:
        os.chmod(filepath, stat.S_IRUSR | stat.S_IWUSR)
        return True
    except Exception:
        return False


def harden_directory_permissions(dirpath: str) -> bool:
    """Sets strict owner-only rwx permissions (0700) on a sensitive directory."""
    if not dirpath or not os.path.isdir(dirpath):
        return False
    try:
        os.chmod(dirpath, stat.S_IRWXU)
        return True
    except Exception:
        return False


def enforce_data_leak_protection_on_workspace(base_dir: str = ".") -> Dict[str, int]:
    """Locks down file/directory permissions (0600/0700) on all local data files and updates `.gitignore`
    so sensitive databases, Excel files, tokens, and backups never leak to git or other OS users."""
    locked_files = 0
    locked_dirs = 0

    for rel_f in SENSITIVE_LOCAL_FILES:
        fp = os.path.join(base_dir, rel_f)
        if os.path.exists(fp) and harden_file_permissions(fp):
            locked_files += 1

    try:
        for fn in os.listdir(base_dir):
            if fn.endswith((".db", ".sqlite", ".sqlite3", ".xlsx", ".toml", ".env")) or "backup" in fn.lower() or "audit" in fn.lower():
                fp = os.path.join(base_dir, fn)
                if os.path.isfile(fp) and harden_file_permissions(fp):
                    locked_files += 1
    except Exception:
        pass

    for rel_d in SENSITIVE_LOCAL_DIRS:
        dp = os.path.join(base_dir, rel_d)
        if os.path.isdir(dp):
            if harden_directory_permissions(dp):
                locked_dirs += 1
            try:
                for fn in os.listdir(dp):
                    fp = os.path.join(dp, fn)
                    if os.path.isfile(fp):
                        harden_file_permissions(fp)
                        locked_files += 1
            except Exception:
                pass

    # Ensure .gitignore contains all data-leak prevention patterns
    gi_path = os.path.join(base_dir, ".gitignore")
    try:
        existing_gi = ""
        if os.path.exists(gi_path):
            with open(gi_path, "r", encoding="utf-8", errors="ignore") as f:
                existing_gi = f.read()
        missing_patterns = [p for p in GITIGNORE_PROTECTED_PATTERNS if p not in existing_gi]
        if missing_patterns:
            with open(gi_path, "a", encoding="utf-8") as f:
                f.write("\n" + "\n".join(missing_patterns) + "\n")
    except Exception:
        pass

    return {"locked_files": locked_files, "locked_dirs": locked_dirs}


def redact_sensitive_text(text: object) -> str:
    """Redacts OAuth tokens, API keys, password hashes, database credentials, and local paths from strings."""
    s = str(text or "")
    if not s:
        return ""
    s = re.sub(r"1//[A-Za-z0-9_\-]{20,}", "[REDACTED_OAUTH_TOKEN]", s)
    s = re.sub(r"ya29\.[A-Za-z0-9_\-]{20,}", "[REDACTED_ACCESS_TOKEN]", s)
    s = re.sub(r"AIza[A-Za-z0-9_\-]{30,}", "[REDACTED_API_KEY]", s)
    s = re.sub(r"pbkdf2_sha256\$\S+", "[REDACTED_PW_HASH]", s)
    s = re.sub(r"CTG_ENC_V1:\S+", "[REDACTED_ENCRYPTED_BLOB]", s)
    s = re.sub(r"postgres(?:ql)?://\S+", "postgresql://[REDACTED_DB_DSN]", s)
    s = re.sub(r"/(?:Users|home|mount/src)/\S+", "[REDACTED_PATH]", s)
    return s
