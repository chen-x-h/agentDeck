"""Entrypoint: uv run python -m mcp_server [--mode local|remote]"""
import sys
from mcp_server.server import main

mode = "local"
if "--mode" in sys.argv:
    idx = sys.argv.index("--mode")
    if idx + 1 < len(sys.argv):
        mode = sys.argv[idx + 1]
elif "--local" in sys.argv:
    mode = "local"

main(mode=mode)
