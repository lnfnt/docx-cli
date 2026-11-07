# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""docx-cli — deterministic .docx generation for Russian legal documents."""
from docx_cli.builder import build_document
from docx_cli.spec import DocSpec
from docx_cli.validate import validate_document

__version__ = "0.1.4"

__all__ = ["build_document", "DocSpec", "validate_document"]
