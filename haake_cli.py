#!/usr/bin/env python3
"""Запуск импорта HAAKE RheoWin .rwd в Excel."""

from haake_rheo.wizard import run


if __name__ == "__main__":
    raise SystemExit(run("import"))
