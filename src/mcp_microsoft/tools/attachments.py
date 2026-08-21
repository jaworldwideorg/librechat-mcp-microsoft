"""
Attachment tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.

Implemented:
  - list_attachments
  - download_attachment
"""

from __future__ import annotations

import asyncio
import base64
import io
import multiprocessing
import sys
from pathlib import Path
from typing import Optional

from fastmcp.utilities.types import File
from fastmcp.exceptions import ToolError
from pypdf import PdfReader

from mcp_microsoft.common.formatting import format_size_display
from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.tooling import READ_ONLY_TOOL, WRITE_TOOL, register_tool
from mcp_microsoft.config import get_app_config
from mcp_microsoft.graph_types import GraphAttachment, parse_graph_collection
from mcp_microsoft.models import (
    AttachmentInfo,
    DownloadAttachmentResponse,
    ListAttachmentsResponse,
    ReadAttachmentResponse,
)
from mcp_microsoft.graph import get_graph

_MAX_READABLE_ATTACHMENT_BYTES = 10 * 1024 * 1024
_MAX_PDF_PAGES = 100
_MAX_PDF_PAGE_STREAM_BYTES = 8 * 1024 * 1024
_MAX_PDF_TOTAL_STREAM_BYTES = 32 * 1024 * 1024
_PDF_WORKER_MEMORY_BYTES = 512 * 1024 * 1024
_PDF_WORKER_CPU_SECONDS = 10
_PDF_WORKER_WALL_SECONDS = 15

# ---------------------------------------------------------------------------
# list_attachments
# ---------------------------------------------------------------------------

class ListAttachmentsInput(ToolRequestModel):
    message_id: str
    profile: str | None = None


class DownloadAttachmentInput(ToolRequestModel):
    message_id: str
    attachment_id: str
    save_path: Path | None = None
    profile: str | None = None


class DownloadAttachmentHttpInput(ToolRequestModel):
    """HTTP input excludes server-local paths that remote callers cannot use."""

    message_id: str
    attachment_id: str
    profile: str | None = None


class ReadAttachmentInput(ToolRequestModel):
    message_id: str
    attachment_id: str
    max_characters: int = 48_000
    profile: str | None = None


def _apply_pdf_worker_limits() -> None:
    """Apply process resource limits where the operating system supports them."""
    try:
        import resource
    except ImportError:  # pragma: no cover - Windows has no resource module
        return

    resource.setrlimit(
        resource.RLIMIT_CPU,
        (_PDF_WORKER_CPU_SECONDS, _PDF_WORKER_CPU_SECONDS + 1),
    )
    if sys.platform.startswith("linux"):
        resource.setrlimit(
            resource.RLIMIT_AS,
            (_PDF_WORKER_MEMORY_BYTES, _PDF_WORKER_MEMORY_BYTES),
        )


def _extract_pdf_text_bounded(
    raw_bytes: bytes, max_characters: int
) -> tuple[str, int, bool]:
    """Extract PDF text while enforcing page and decoded stream budgets."""
    reader = PdfReader(io.BytesIO(raw_bytes))
    page_count = len(reader.pages)
    if page_count > _MAX_PDF_PAGES:
        raise ValueError(f"PDF exceeds the {_MAX_PDF_PAGES}-page extraction limit.")

    parts: list[str] = []
    character_count = 0
    total_stream_bytes = 0
    truncated = False
    for page in reader.pages:
        contents = page.get_contents()
        stream_size = len(contents.get_data()) if contents is not None else 0
        if stream_size > _MAX_PDF_PAGE_STREAM_BYTES:
            raise ValueError("A PDF page exceeds the decoded content-stream limit.")
        total_stream_bytes += stream_size
        if total_stream_bytes > _MAX_PDF_TOTAL_STREAM_BYTES:
            raise ValueError("PDF exceeds the total decoded content-stream limit.")

        page_text = page.extract_text() or ""
        separator = "\n\n" if parts else ""
        remaining = max_characters + 1 - character_count
        addition = (separator + page_text)[:remaining]
        parts.append(addition)
        character_count += len(addition)
        if character_count > max_characters:
            truncated = True
            break

    text = "".join(parts)
    return text[:max_characters], page_count, truncated


def _pdf_worker(send_conn, raw_bytes: bytes, max_characters: int) -> None:
    """Run bounded PDF extraction in an isolated child process."""
    try:
        _apply_pdf_worker_limits()
        send_conn.send(("ok", _extract_pdf_text_bounded(raw_bytes, max_characters)))
    except BaseException as exc:
        send_conn.send(("error", str(exc) or type(exc).__name__))
    finally:
        send_conn.close()


def _extract_pdf_text_isolated(
    raw_bytes: bytes, max_characters: int
) -> tuple[str, int, bool]:
    """Extract PDF text in a disposable spawned process with a wall-time limit."""
    context = multiprocessing.get_context("spawn")
    receive_conn, send_conn = context.Pipe(duplex=False)
    process = context.Process(
        target=_pdf_worker,
        args=(send_conn, raw_bytes, max_characters),
        daemon=True,
    )
    process.start()
    send_conn.close()
    try:
        if not receive_conn.poll(_PDF_WORKER_WALL_SECONDS):
            process.terminate()
            process.join(timeout=1)
            if process.is_alive():
                process.kill()
                process.join(timeout=1)
            raise TimeoutError(
                f"PDF extraction exceeded the {_PDF_WORKER_WALL_SECONDS}-second limit."
            )
        try:
            status, payload = receive_conn.recv()
        except EOFError as exc:
            raise RuntimeError("PDF extraction worker exceeded a resource limit.") from exc
        process.join(timeout=1)
        if process.is_alive():
            process.terminate()
            process.join(timeout=1)
        if status == "error":
            raise ValueError(payload)
        return payload
    finally:
        receive_conn.close()
        if process.is_alive():
            process.terminate()
            process.join(timeout=1)
        close_process = getattr(process, "close", None)
        if close_process is not None:
            close_process()


async def list_attachments(params: ListAttachmentsInput) -> ListAttachmentsResponse:
    """
    List all attachments on an email message.

    Args:
        message_id: The Graph message ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured attachment metadata.
    """
    g = get_graph(params.profile)
    query = {
        "$select": "id,name,size,contentType,isInline",
    }

    result = await g.get(f"/me/messages/{params.message_id}/attachments", params=query)
    attachments = parse_graph_collection(result, GraphAttachment)

    items: list[AttachmentInfo] = []
    for att in attachments:
        size_bytes = att.size
        items.append(
            AttachmentInfo(
                id=att.id,
                name=att.name,
                size_bytes=size_bytes,
                size_display=format_size_display(size_bytes),
                content_type=att.content_type or "unknown",
                is_inline=att.is_inline,
            )
        )

    return ListAttachmentsResponse(
        message_id=params.message_id,
        count=len(items),
        attachments=items,
    )


# ---------------------------------------------------------------------------
# download_attachment
# ---------------------------------------------------------------------------


async def download_attachment(
    params: DownloadAttachmentInput,
) -> DownloadAttachmentResponse | File:
    """
    Download an email attachment, saving it to disk or returning a file payload.

    When running in Claude Desktop, always provide save_path (e.g. the user's
    Downloads folder) so the file is written to disk. Omitting save_path returns
    a FastMCP file object, which is only useful for programmatic embedding.

    Args:
        message_id: The Graph message ID.
        attachment_id: The Graph attachment ID.
        save_path: Optional filesystem path to save the file. Omit to return a FastMCP file payload.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        A FastMCP file when no save path is given, or structured file-save metadata.
    """
    if params.save_path is not None and get_app_config().transport == "http":
        raise ValueError(
            "save_path is not available in multi-user http mode (the server's "
            "disk is not the caller's disk); omit it to receive the file inline."
        )

    g = get_graph(params.profile)
    result = GraphAttachment.model_validate(await g.get(
        f"/me/messages/{params.message_id}/attachments/{params.attachment_id}"
    ) or {})

    att_name = result.name or "attachment"
    content_type = result.content_type or "application/octet-stream"
    content_bytes_b64: Optional[str] = result.content_bytes

    if content_bytes_b64 is None:
        return DownloadAttachmentResponse(
            success=False,
            action="download_attachment",
            message_id=params.message_id,
            attachment_id=params.attachment_id,
            filename=att_name,
            error="Attachment has no downloadable content.",
        )

    raw_bytes = base64.b64decode(content_bytes_b64)

    if params.save_path is None:
        file_format = Path(att_name).suffix.lstrip(".")
        if not file_format and "/" in content_type:
            file_format = content_type.split("/", 1)[1]
        return File(data=raw_bytes, format=file_format or None, name=att_name)

    # Resolve final output path — sanitize remote filename to prevent traversal
    resolved_path = params.save_path
    if params.save_path.is_dir():
        safe_name = Path(att_name).name  # strip directory components
        if not safe_name or safe_name.startswith("."):
            safe_name = "attachment"
        resolved_path = params.save_path / safe_name

    resolved_path.parent.mkdir(parents=True, exist_ok=True)
    with resolved_path.open("wb") as fh:
        fh.write(raw_bytes)

    size_str = format_size_display(len(raw_bytes))
    return DownloadAttachmentResponse(
        success=True,
        action="download_attachment",
        message_id=params.message_id,
        attachment_id=params.attachment_id,
        path=str(resolved_path),
        filename=att_name,
        size_bytes=len(raw_bytes),
        size_display=size_str,
        content_type=content_type,
    )


async def download_attachment_http(
    params: DownloadAttachmentHttpInput,
) -> DownloadAttachmentResponse | File:
    """Download an email attachment and return it inline to the remote client.

    The remote HTTP server cannot write to a caller's filesystem. The attachment
    is therefore always returned as an inline FastMCP file payload.

    Args:
        message_id: The Graph message ID.
        attachment_id: The Graph attachment ID.
        profile: Ignored in HTTP mode; identity comes from the caller's token.

    Returns:
        The attachment as an inline FastMCP file payload.
    """
    return await download_attachment(
        DownloadAttachmentInput(
            message_id=params.message_id,
            attachment_id=params.attachment_id,
            profile=params.profile,
        )
    )


async def read_attachment(params: ReadAttachmentInput) -> ReadAttachmentResponse:
    """Extract readable text from an email attachment.

    PDF and plain-text attachments are read directly from Microsoft Graph, so
    remote clients do not need to download and re-upload a file before an AI can
    inspect its contents. Scanned image-only PDFs require OCR and are reported
    clearly when they contain no extractable text.

    Args:
        message_id: The Graph message ID.
        attachment_id: The Graph attachment ID.
        max_characters: Maximum extracted characters returned (1–100,000).
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Extracted attachment text and PDF page count when applicable.
    """
    g = get_graph(params.profile)
    attachment_path = (
        f"/me/messages/{params.message_id}/attachments/{params.attachment_id}"
    )
    metadata = GraphAttachment.model_validate(
        await g.get(
            attachment_path,
            params={"$select": "id,name,size,contentType,isInline"},
        )
        or {}
    )
    if metadata.size > _MAX_READABLE_ATTACHMENT_BYTES:
        raise ToolError("Attachment exceeds the 10 MiB text-extraction limit.")

    attachment = GraphAttachment.model_validate(
        await g.get(attachment_path) or {}
    )
    if attachment.content_bytes is None:
        raise ToolError("Attachment has no readable file content.")

    max_encoded_size = ((_MAX_READABLE_ATTACHMENT_BYTES + 2) // 3) * 4
    if len(attachment.content_bytes) > max_encoded_size:
        raise ToolError("Attachment exceeds the 10 MiB text-extraction limit.")

    try:
        raw_bytes = base64.b64decode(attachment.content_bytes, validate=True)
    except (ValueError, TypeError) as exc:
        raise ToolError("Attachment content returned by Microsoft Graph is invalid.") from exc

    if len(raw_bytes) > _MAX_READABLE_ATTACHMENT_BYTES:
        raise ToolError("Attachment exceeds the 10 MiB text-extraction limit.")

    content_type = (attachment.content_type or metadata.content_type or "").casefold()
    suffix = Path(attachment.name).suffix.casefold()
    page_count = 0
    limit = min(max(1, params.max_characters), 100_000)
    if content_type == "application/pdf" or suffix == ".pdf":
        try:
            text, page_count, truncated = await asyncio.to_thread(
                _extract_pdf_text_isolated, raw_bytes, limit
            )
        except TimeoutError as exc:
            raise ToolError(str(exc)) from exc
        except ValueError as exc:
            raise ToolError(str(exc)) from exc
        except Exception as exc:
            raise ToolError(
                "Unable to extract text from this PDF attachment within safe resource limits."
            ) from exc
        if not text.strip():
            raise ToolError(
                "This PDF contains no extractable text; it may be scanned or image-only."
            )
    elif content_type.startswith("text/") or suffix in {
        ".csv",
        ".json",
        ".md",
        ".rtf",
        ".txt",
        ".xml",
    }:
        try:
            text = raw_bytes.decode("utf-8-sig")
        except UnicodeDecodeError as exc:
            raise ToolError("Unable to decode this text attachment as UTF-8.") from exc
        truncated = len(text) > limit
    else:
        raise ToolError(
            "Text extraction currently supports PDF and plain-text attachments; "
            "use download_attachment to retrieve other file types."
        )

    return ReadAttachmentResponse(
        message_id=params.message_id,
        attachment_id=params.attachment_id,
        filename=attachment.name,
        content_type=attachment.content_type,
        text=text[:limit],
        page_count=page_count,
        truncated=truncated,
    )


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------
def register(server) -> None:
    """Register all attachment tools with the given FastMCP server instance."""
    register_tool(server, list_attachments, annotations=READ_ONLY_TOOL)
    if get_app_config().transport == "http":
        register_tool(
            server,
            download_attachment_http,
            name="download_attachment",
            annotations=READ_ONLY_TOOL,
        )
    else:
        register_tool(server, download_attachment, annotations=WRITE_TOOL)
    register_tool(server, read_attachment, annotations=READ_ONLY_TOOL)
