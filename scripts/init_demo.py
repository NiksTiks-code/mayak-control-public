#!/usr/bin/env python3
"""Initialize a local MAYAK demo without shipping reusable credentials."""

import fcntl
import getpass
import json
import os
from pathlib import Path
import re
import uuid

from werkzeug.security import generate_password_hash


ROOT = Path(__file__).resolve().parents[1]
DEMO = ROOT / "demo"


def validate_password(value):
    categories = sum((
        bool(re.search(r"[a-z]", value)),
        bool(re.search(r"[A-Z]", value)),
        bool(re.search(r"\d", value)),
        bool(re.search(r"[^A-Za-z0-9]", value)),
    ))
    if len(value) < 12 or categories < 3:
        raise ValueError("Password must contain at least 12 characters and 3 character categories")


def prompt_password(username):
    first = getpass.getpass(f"New password for {username}: ")
    second = getpass.getpass("Confirm password: ")
    if first != second:
        raise ValueError(f"Password confirmation failed for {username}")
    validate_password(first)
    return first


def atomic_write_json(path, payload, mode):
    lock_path = path.with_name(f"{path.name}.lock")
    with lock_path.open("a+") as lock:
        fcntl.flock(lock.fileno(), fcntl.LOCK_EX)
        temporary = path.with_name(f"{path.name}.{os.getpid()}.{uuid.uuid4().hex}.tmp")
        try:
            with temporary.open("w", encoding="utf-8") as handle:
                json.dump(payload, handle, ensure_ascii=False, indent=2)
                handle.write("\n")
                handle.flush()
                os.fsync(handle.fileno())
            os.chmod(temporary, mode)
            os.replace(temporary, path)
        finally:
            if temporary.exists():
                temporary.unlink()
            fcntl.flock(lock.fileno(), fcntl.LOCK_UN)


def load_example(name):
    return json.loads((DEMO / name).read_text(encoding="utf-8"))


def main():
    users_path = ROOT / "users.json"
    if users_path.exists():
        raise SystemExit("users.json already exists; refusing to overwrite it")

    admin_password = prompt_password("demo_admin")
    contractor_password = prompt_password("demo_contractor")
    users = {
        "demo_admin": {
            "password_hash": generate_password_hash(admin_password),
            "role": "super_admin",
            "name": "Portfolio Administrator",
        },
        "demo_contractor": {
            "password_hash": generate_password_hash(contractor_password),
            "role": "contractor",
            "name": "Demo Contractor",
        },
    }
    admin_password = contractor_password = None

    atomic_write_json(ROOT / "data.json", load_example("data.example.json"), 0o600)
    atomic_write_json(
        ROOT / "contractors.json",
        load_example("contractors.example.json"),
        0o600,
    )
    atomic_write_json(users_path, users, 0o600)
    print("Demo initialized successfully")
    print("Accounts created: demo_admin, demo_contractor")
    print("Passwords were not stored or displayed")


if __name__ == "__main__":
    main()
