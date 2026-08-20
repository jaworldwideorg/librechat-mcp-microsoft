"""
Attachment tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.

Implemented:
  - list_attachments
  - download_attachment
"""

from __future__ import annotations

import base64
from pathlib import Path
from typing import Optional

from fastmcp.utilities.types import File

from mcp_microsoft.common.formatting import format_size_display
from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.tooling import READ_ONLY_TOOL, WRITE_TOOL, register_tool
from mcp_microsoft.config import get_app_config
from mcp_microsoft.graph_types import GraphAttachment, parse_graph_collection
from mcp_microsoft.models import AttachmentInfo, DownloadAttachmentResponse, ListAttachmentsResponse
from mcp_microsoft.graph import get_graph

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
