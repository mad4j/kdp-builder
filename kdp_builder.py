#!/usr/bin/env python3
"""Compatibility wrapper for the `kdpbuilder` package.

This repository previously shipped a standalone implementation in this file.
The source of truth is now the `kdpbuilder` package (CLI + library).

Use one of:
- python -m kdpbuilder ...
- python kdp_builder.py ... (legacy wrapper)
"""

from __future__ import annotations

from kdpbuilder.cli import main as _main


if __name__ == "__main__":
    _main()
