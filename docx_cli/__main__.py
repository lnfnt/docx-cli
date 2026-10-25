"""Allow `python -m docx_cli`."""
import sys

from docx_cli.cli import main

if __name__ == "__main__":
    sys.exit(main())
