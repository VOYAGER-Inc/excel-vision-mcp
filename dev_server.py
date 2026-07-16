"""
Wrapper for running with `mcp dev` or `mcp inspector`.
Resolves relative import issues by using absolute imports.
"""
import sys
from pathlib import Path

# Add src to path for absolute imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from excel_mcp.server import mcp  # noqa: E402

if __name__ == "__main__":
    mcp.run()
