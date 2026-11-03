# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""S2-M1: split_run не молча портит run'ы с табуляциями/разрывами."""
from __future__ import annotations

import pytest


def test_split_run_with_tab_fails_closed():
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    from docx_cli.comments import CommentError, _split_run
    r = parse_xml(
        f'<w:r {nsdecls("w")}><w:t>абв</w:t><w:tab/></w:r>')
    with pytest.raises(CommentError):
        _split_run(r, 2)


def test_split_run_plain_text_still_works():
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls, qn
    from docx_cli.comments import _split_run
    p = parse_xml(
        f'<w:p {nsdecls("w")}><w:r><w:t>абвгд</w:t></w:r></w:p>')
    r = p.find(qn("w:r"))
    clone = _split_run(r, 2)
    left = "".join(t.text or "" for t in r.findall(qn("w:t")))
    right = "".join(t.text or "" for t in clone.findall(qn("w:t")))
    assert left == "аб" and right == "вгд"
