from __future__ import annotations

import base64
from pathlib import Path

import pytest
from fastmcp.utilities.types import File

from mcp_microsoft.models import (
    CreateListItemResponse,
    GetListItemsResponse,
    SharePointFields,
    UpdateListItemResponse,
    UploadFileResponse,
    UploadSiteFileResponse,
)
import mcp_microsoft.server as server
from mcp_microsoft.runtime import reset_runtime_state
from mcp_microsoft.tools import attachments, mail, onedrive, sharepoint


@pytest.mark.asyncio
async def test_sharepoint_tools_are_registered_when_enabled(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("MCP_ENABLE_SHAREPOINT", "true")
    reset_runtime_state()
    tool_names = {
        tool.name
        for tool in await server.get_mcp_server(reset=True).list_tools(run_middleware=False)
    }
    assert "search_sharepoint_sites" in tool_names
    monkeypatch.delenv("MCP_ENABLE_SHAREPOINT", raising=False)
    reset_runtime_state()


@pytest.mark.asyncio
async def test_download_attachment_returns_fastmcp_file(monkeypatch: pytest.MonkeyPatch) -> None:
    class DummyGraph:
        async def get(self, _path: str):
            return {
                "name": "hello.txt",
                "contentType": "text/plain",
                "contentBytes": base64.b64encode(b"hello world").decode("ascii"),
            }

    monkeypatch.setattr(attachments, "get_graph", lambda _profile: DummyGraph())

    result = await attachments.download_attachment(
        attachments.DownloadAttachmentInput(
            message_id="message-id",
            attachment_id="attachment-id",
        )
    )

    assert isinstance(result, File)
    assert result.data == b"hello world"
    assert result._name == "hello.txt"


@pytest.mark.asyncio
async def test_read_attachment_extracts_pdf_text(monkeypatch: pytest.MonkeyPatch) -> None:
    class DummyGraph:
        async def get(self, _path: str):
            return {
                "name": "report.pdf",
                "contentType": "application/pdf",
                "contentBytes": base64.b64encode(b"pdf bytes").decode("ascii"),
            }

    class DummyPage:
        def __init__(self, text: str) -> None:
            self.text = text

        def extract_text(self) -> str:
            return self.text

    class DummyReader:
        def __init__(self, _stream) -> None:
            self.pages = [DummyPage("Page one"), DummyPage("Page two")]

    monkeypatch.setattr(attachments, "get_graph", lambda _profile: DummyGraph())
    monkeypatch.setattr(attachments, "PdfReader", DummyReader)

    result = await attachments.read_attachment(
        attachments.ReadAttachmentInput(
            message_id="message-id",
            attachment_id="attachment-id",
        )
    )

    assert result.filename == "report.pdf"
    assert result.page_count == 2
    assert result.text == "Page one\n\nPage two"
    assert result.truncated is False


@pytest.mark.asyncio
async def test_read_attachment_truncates_text(monkeypatch: pytest.MonkeyPatch) -> None:
    class DummyGraph:
        async def get(self, _path: str):
            return {
                "name": "notes.txt",
                "contentType": "text/plain",
                "contentBytes": base64.b64encode(b"abcdefgh").decode("ascii"),
            }

    monkeypatch.setattr(attachments, "get_graph", lambda _profile: DummyGraph())

    result = await attachments.read_attachment(
        attachments.ReadAttachmentInput(
            message_id="message-id",
            attachment_id="attachment-id",
            max_characters=4,
        )
    )

    assert result.text == "abcd"
    assert result.truncated is True


@pytest.mark.asyncio
async def test_onedrive_large_upload_does_not_read_entire_file(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    large_file = tmp_path / "large.bin"
    large_file.write_bytes(b"x" * (4 * 1024 * 1024 + 1))

    class DummyGraph:
        async def post(self, _path: str, json: dict | None = None):
            return {"uploadUrl": "https://example.invalid/upload"}

    captured: dict[str, object] = {}

    async def fake_large_upload(upload_url: str, file_path: Path, total_size: int, ctx=None):
        captured["upload_url"] = upload_url
        captured["file_path"] = file_path
        captured["total_size"] = total_size
        return {"id": "drive-item", "webUrl": "https://example.invalid/file"}

    monkeypatch.setattr(onedrive, "get_graph", lambda _profile: DummyGraph())
    monkeypatch.setattr(onedrive, "upload_large_file_via_session", fake_large_upload)
    monkeypatch.setattr(Path, "read_bytes", lambda self: (_ for _ in ()).throw(AssertionError("read_bytes should not be used for large uploads")))

    result = await onedrive.upload_file(
        onedrive.UploadFileInput(local_path=large_file)
    )

    assert isinstance(result, UploadFileResponse)
    assert result.success is True
    assert result.action == "upload_file"
    assert result.file_id == "drive-item"
    assert captured["upload_url"] == "https://example.invalid/upload"
    assert captured["file_path"] == large_file
    assert captured["total_size"] == large_file.stat().st_size


@pytest.mark.asyncio
async def test_sharepoint_large_upload_does_not_read_entire_file(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    large_file = tmp_path / "large.bin"
    large_file.write_bytes(b"x" * (4 * 1024 * 1024 + 1))

    class DummyGraph:
        async def post(self, _path: str, json: dict | None = None):
            return {"uploadUrl": "https://example.invalid/upload"}

    captured: dict[str, object] = {}

    async def fake_large_upload(upload_url: str, file_path: Path, total_size: int, ctx=None):
        captured["upload_url"] = upload_url
        captured["file_path"] = file_path
        captured["total_size"] = total_size
        return {"id": "drive-item", "webUrl": "https://example.invalid/file"}

    monkeypatch.setattr(sharepoint, "_get_sharepoint_graph", lambda _profile: DummyGraph())
    monkeypatch.setattr(sharepoint, "upload_large_file_via_session", fake_large_upload)
    monkeypatch.setattr(Path, "read_bytes", lambda self: (_ for _ in ()).throw(AssertionError("read_bytes should not be used for large uploads")))

    result = await sharepoint.upload_to_site(
        sharepoint.UploadToSiteInput(
            site_id="site-id",
            drive_id="drive-id",
            local_path=large_file,
        )
    )

    assert isinstance(result, UploadSiteFileResponse)
    assert result.success is True
    assert result.action == "upload_to_site"
    assert result.file_id == "drive-item"
    assert captured["upload_url"] == "https://example.invalid/upload"
    assert captured["file_path"] == large_file
    assert captured["total_size"] == large_file.stat().st_size


@pytest.mark.asyncio
async def test_sharepoint_list_item_tools_accept_structured_fields(monkeypatch: pytest.MonkeyPatch) -> None:
    calls: list[dict] = []

    class DummyGraph:
        async def post(self, _path: str, json: dict | None = None):
            calls.append({"method": "post", "json": json})
            return {"id": "item-123"}

        async def patch(self, _path: str, json: dict | None = None):
            calls.append({"method": "patch", "json": json})
            return {}

    monkeypatch.setattr(sharepoint, "_get_sharepoint_graph", lambda _profile: DummyGraph())

    create_result = await sharepoint.create_list_item(
        sharepoint.CreateListItemInput(
            site_id="site-id",
            list_id="list-id",
            fields=SharePointFields({"Title": "Hello"}),
        )
    )
    update_result = await sharepoint.update_list_item(
        sharepoint.UpdateListItemInput(
            site_id="site-id",
            list_id="list-id",
            item_id="item-123",
            fields=SharePointFields({"Status": "Done"}),
        )
    )

    assert isinstance(create_result, CreateListItemResponse)
    assert create_result.success is True
    assert isinstance(create_result.fields, SharePointFields)
    assert create_result.fields.root == {"Title": "Hello"}
    assert isinstance(update_result, UpdateListItemResponse)
    assert update_result.success is True
    assert update_result.updated_fields == ["Status"]
    assert isinstance(update_result.fields, SharePointFields)
    assert update_result.fields.root == {"Status": "Done"}
    assert calls == [
        {"method": "post", "json": {"fields": {"Title": "Hello"}}},
        {"method": "patch", "json": {"Status": "Done"}},
    ]


@pytest.mark.asyncio
async def test_get_list_items_returns_structured_sharepoint_fields(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyGraph:
        async def get(self, _path: str, params: dict | None = None):
            return {
                "value": [
                    {
                        "id": "item-1",
                        "createdDateTime": "2026-01-01T12:00:00Z",
                        "lastModifiedDateTime": "2026-01-02T12:00:00Z",
                        "fields": {
                            "Title": "Task",
                            "Status": "Done",
                            "_UIVersionString": "1.0",
                            "ContentType": "Item",
                        },
                    }
                ]
            }

    monkeypatch.setattr(sharepoint, "_get_sharepoint_graph", lambda _profile: DummyGraph())

    result = await sharepoint.get_list_items(
        sharepoint.GetListItemsInput(site_id="site-id", list_id="list-id")
    )

    assert isinstance(result, GetListItemsResponse)
    assert result.count == 1
    assert result.items[0].title == "Task"
    assert isinstance(result.items[0].fields, SharePointFields)
    assert result.items[0].fields.root == {"Status": "Done"}


@pytest.mark.asyncio
async def test_bulk_move_emails_parses_typed_batch_message_responses(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyGraph:
        async def batch(self, _requests: list[dict]):
            return [
                {
                    "id": "1",
                    "status": 201,
                    "body": {
                        "id": "moved-message-id",
                        "subject": "Moved",
                    },
                }
            ]

    monkeypatch.setattr(mail, "get_graph", lambda _profile: DummyGraph())

    result = await mail.bulk_move_emails(
        mail.BulkMoveEmailsInput(
            message_ids=["original-message-id"],
            destination_folder="archive",
        )
    )

    assert result.success is True
    assert result.succeeded == 1
    assert result.failed == 0
    assert result.moved == [
        mail.BulkMovedEmail(
            source_message_id="original-message-id",
            new_message_id="moved-message-id",
        )
    ]
