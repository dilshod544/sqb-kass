"""
Конфигурация проекта Cashier Intelligence.
"""

from pathlib import Path

# ── Пути ────────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parents[2]   # корень проекта
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
