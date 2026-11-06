"""outlook-app-mcp: an MCP server for the Outlook desktop app (macOS and Windows)."""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("outlook-app-mcp")
except PackageNotFoundError:  # running from a source tree that was never installed
    __version__ = "0+unknown"


def main() -> None:
    from .server import main as run

    run()
