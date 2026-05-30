#!/usr/bin/env python3
"""List the 3 most recent NotebookLM notebooks for yoav8288@gmail.com.

RUN THIS ON YOUR LOCAL MACHINE — NOT in a remote container.

Prerequisites:
  - notebooklm-py installed:  uv pip install -e .
  - Authenticated profile "yoav8288" — set up via list-recent-yoav.sh
    or:  notebooklm -p yoav8288 login

The "most recent" ordering uses the order returned by the LIST_NOTEBOOKS RPC,
which is newest-first. NotebookLM exposes `created_at` but no `modified_at`,
so "recent" means most recently created.
"""

from __future__ import annotations

import asyncio
import os
import sys

os.environ.setdefault("NOTEBOOKLM_PROFILE", "yoav8288")

try:
    from notebooklm import NotebookLMClient
except ImportError:
    sys.stderr.write(
        "notebooklm package not installed. From this directory:\n"
        "  uv sync --frozen --extra browser --extra dev\n"
        "  source .venv/bin/activate\n"
    )
    sys.exit(1)


async def main() -> int:
    try:
        async with await NotebookLMClient.from_storage() as client:
            notebooks = await client.notebooks.list()
    except Exception as exc:
        sys.stderr.write(f"Auth or transport error: {exc}\n")
        sys.stderr.write("Refresh auth with:  notebooklm -p yoav8288 login\n")
        return 1

    if not notebooks:
        print("No notebooks found.")
        return 0

    print(f"{'Created':<12}  {'Title':<50}  {'Sources':<8}  ID")
    print("-" * 100)
    for nb in notebooks[:3]:
        created = nb.created_at.strftime("%Y-%m-%d") if nb.created_at else "—"
        title = (nb.title[:47] + "...") if len(nb.title) > 50 else nb.title
        print(f"{created:<12}  {title:<50}  {nb.sources_count:<8}  {nb.id}")
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
