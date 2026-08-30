"""
shared_rules.py

Sumber aturan kategori/kata kunci/alias yang dipakai BERSAMA oleh reconbot
dan bank-statement-bot (bot ini), supaya merchant/kategori/alias yang
"diajarkan" ke satu bot otomatis dikenali bot yang lain juga -- tidak perlu
duplikasi aturan terpisah di masing-masing repo.

Urutan sumber (yang pertama tersedia dipakai): Postgres (env var
DATABASE_URL, tabel shared_rules -- Postgres yang SAMA dengan stoabot di
Railway, bukan instance baru) -> file JSON lokal (shared_rules.json) ->
{} (pemanggil fallback ke default hardcode masing-masing).

Skema tabel Postgres (lihat migrate.sql di repo reconbot -- tabelnya
dibuat lewat migrasi reconbot, bot ini cuma BACA, tidak perlu migrasi
sendiri):
    CREATE TABLE shared_rules (
        key TEXT PRIMARY KEY,
        value JSONB NOT NULL,
        updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    );
Satu baris per top-level key (category_override_rules, employee_aliases,
dst) -- value-nya persis struktur yang sama dengan shared_rules.json.

Didesain defensif: kalau psycopg2 tidak terpasang, DATABASE_URL tidak
diset, atau koneksi gagal/timeout -- semua fallback otomatis ke JSON lalu
ke {} tanpa pernah crash. Bot ini HANYA membaca (tidak pernah menulis) ke
tabel shared_rules -- reconbot yang jadi sumber kebenaran untuk push data
baru ke situ.
"""

import json
import os

_JSON_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "shared_rules.json")
_cache = None


def _load_from_postgres():
    dsn = os.environ.get("DATABASE_URL")
    if not dsn:
        return None
    try:
        import psycopg2
    except ImportError:
        return None
    try:
        conn = psycopg2.connect(dsn, connect_timeout=5)
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT key, value FROM shared_rules")
                rows = cur.fetchall()
        finally:
            conn.close()
    except Exception:
        return None
    return {key: value for key, value in rows}


def _load_from_json(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


def load(path=None):
    global _cache
    if path is None and _cache is not None:
        return _cache

    data = _load_from_postgres()
    if data is None:
        data = _load_from_json(path or _JSON_PATH)
    if data is None:
        data = {}

    if path is None:
        _cache = data
    return data


def get(key, default):
    return load().get(key, default)
