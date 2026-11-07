# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""Allow `python -m docx_cli`."""
import sys

from docx_cli.cli import main

if __name__ == "__main__":
    sys.exit(main())
