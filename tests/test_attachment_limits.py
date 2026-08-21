from __future__ import annotations

import io
from types import SimpleNamespace

import pytest
from fastmcp.exceptions import ToolError
from pypdf import PdfWriter

from mcp_microsoft.tools import attachments


def _page(text: str = "text", stream: bytes = b""):
    contents = SimpleNamespace(get_data=lambda: stream) if stream else None
    return SimpleNamespace(
        get_contents=lambda: contents,
        extract_text=lambda: text,
    )


@pytest.mark.asyncio
async def test_read_attachment_rejects_declared_oversize_before_content_fetch(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls: list[dict | None] = []

    class DummyGraph:
        async def get(self, _path: str, params: dict | None = None):
            calls.append(params)
            return {
                "name": "large.pdf",
                "contentType": "application/pdf",
                "size": attachments._MAX_READABLE_ATTACHMENT_BYTES + 1,
            }

    monkeypatch.setattr(attachments, "get_graph", lambda _profile: DummyGraph())
    with pytest.raises(ToolError, match="10 MiB"):
        await attachments.read_attachment(
            attachments.ReadAttachmentInput(message_id="m", attachment_id="a")
        )

    assert calls == [{"$select": "id,name,size,contentType,isInline"}]


@pytest.mark.asyncio
async def test_read_attachment_rejects_oversize_encoded_payload(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(attachments, "_MAX_READABLE_ATTACHMENT_BYTES", 4)

    class DummyGraph:
        async def get(self, _path: str, params: dict | None = None):
            if params is not None:
                return {"name": "large.txt", "contentType": "text/plain", "size": 0}
            return {
                "name": "large.txt",
                "contentType": "text/plain",
                "contentBytes": "A" * 9,
            }

    monkeypatch.setattr(attachments, "get_graph", lambda _profile: DummyGraph())
    with pytest.raises(ToolError, match="10 MiB"):
        await attachments.read_attachment(
            attachments.ReadAttachmentInput(message_id="m", attachment_id="a")
        )


@pytest.mark.asyncio
async def test_read_attachment_rejects_oversize_decoded_payload(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(attachments, "_MAX_READABLE_ATTACHMENT_BYTES", 4)
    monkeypatch.setattr(attachments, "base64", SimpleNamespace(
        b64decode=lambda *_args, **_kwargs: b"12345"
    ))

    class DummyGraph:
        async def get(self, _path: str, params: dict | None = None):
            if params is not None:
                return {"name": "large.txt", "contentType": "text/plain", "size": 0}
            return {
                "name": "large.txt",
                "contentType": "text/plain",
                "contentBytes": "AAAA",
            }

    monkeypatch.setattr(attachments, "get_graph", lambda _profile: DummyGraph())
    with pytest.raises(ToolError, match="10 MiB"):
        await attachments.read_attachment(
            attachments.ReadAttachmentInput(message_id="m", attachment_id="a")
        )


def test_pdf_extraction_rejects_page_limit(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        attachments,
        "PdfReader",
        lambda _stream: SimpleNamespace(pages=[_page()] * 101),
    )
    with pytest.raises(ValueError, match="100-page"):
        attachments._extract_pdf_text_bounded(b"pdf", 100)


def test_pdf_extraction_rejects_page_stream_limit(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(attachments, "_MAX_PDF_PAGE_STREAM_BYTES", 3)
    monkeypatch.setattr(
        attachments,
        "PdfReader",
        lambda _stream: SimpleNamespace(pages=[_page(stream=b"1234")]),
    )
    with pytest.raises(ValueError, match="page exceeds"):
        attachments._extract_pdf_text_bounded(b"pdf", 100)


def test_pdf_extraction_rejects_total_stream_limit(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(attachments, "_MAX_PDF_PAGE_STREAM_BYTES", 10)
    monkeypatch.setattr(attachments, "_MAX_PDF_TOTAL_STREAM_BYTES", 5)
    monkeypatch.setattr(
        attachments,
        "PdfReader",
        lambda _stream: SimpleNamespace(
            pages=[_page(stream=b"123"), _page(stream=b"456")]
        ),
    )
    with pytest.raises(ValueError, match="total decoded"):
        attachments._extract_pdf_text_bounded(b"pdf", 100)


def test_pdf_extraction_stops_at_character_limit(monkeypatch: pytest.MonkeyPatch) -> None:
    pages = [_page("abcdef"), _page("must not be visited")]
    monkeypatch.setattr(
        attachments,
        "PdfReader",
        lambda _stream: SimpleNamespace(pages=pages),
    )
    text, page_count, truncated = attachments._extract_pdf_text_bounded(b"pdf", 4)

    assert text == "abcd"
    assert page_count == 2
    assert truncated is True


def test_pdf_worker_is_terminated_after_wall_timeout(monkeypatch: pytest.MonkeyPatch) -> None:
    process = SimpleNamespace(
        alive=False,
        start=lambda: setattr(process, "alive", True),
        terminate=lambda: setattr(process, "alive", False),
        kill=lambda: setattr(process, "alive", False),
        join=lambda timeout=None: None,
        is_alive=lambda: process.alive,
    )
    receive = SimpleNamespace(poll=lambda _timeout: False, close=lambda: None)
    send = SimpleNamespace(close=lambda: None)
    context = SimpleNamespace(
        Pipe=lambda duplex=False: (receive, send),
        Process=lambda **_kwargs: process,
    )
    monkeypatch.setattr(attachments.multiprocessing, "get_context", lambda _mode: context)

    with pytest.raises(TimeoutError, match="15-second"):
        attachments._extract_pdf_text_isolated(b"pdf", 100)
    assert process.alive is False


def test_pdf_worker_extracts_in_spawned_process() -> None:
    stream = io.BytesIO()
    writer = PdfWriter()
    writer.add_blank_page(width=72, height=72)
    writer.write(stream)

    text, page_count, truncated = attachments._extract_pdf_text_isolated(
        stream.getvalue(), 100
    )

    assert text == ""
    assert page_count == 1
    assert truncated is False
