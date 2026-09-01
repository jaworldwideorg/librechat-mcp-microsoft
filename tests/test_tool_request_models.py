from __future__ import annotations

import pytest
from fastmcp import FastMCP

from mcp_microsoft.tools import contacts, drafts, mail, onedrive, profiles, services, sharepoint


def _inner_params_schema(tool) -> dict:
    schema = tool.parameters
    params_schema = schema["properties"]["params"]
    ref = params_schema["$ref"].split("/")[-1]
    return schema["$defs"][ref]


@pytest.mark.asyncio
async def test_create_contact_tool_uses_standardized_params_object() -> None:
    mcp = FastMCP("test-server")
    contacts.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}
    tool = tools["create_contact"]

    assert set(tool.parameters["properties"]) == {"params"}
    assert tool.parameters["required"] == ["params"]

    params_schema = _inner_params_schema(tool)
    assert "display_name" in params_schema["properties"]
    assert "display_name" in params_schema["required"]
    assert "display_name" not in tool.parameters["properties"]


@pytest.mark.asyncio
async def test_send_email_tool_hides_ctx_and_uses_params_object() -> None:
    mcp = FastMCP("test-server")
    mail.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}
    tool = tools["send_email"]

    assert set(tool.parameters["properties"]) == {"params"}
    assert "ctx" not in tool.parameters["properties"]

    params_schema = _inner_params_schema(tool)
    assert set(params_schema["required"]) == {"to", "subject", "body"}
    assert "confirm" in params_schema["properties"]


@pytest.mark.asyncio
async def test_reply_email_schema_documents_body_type_compatibility() -> None:
    mcp = FastMCP("test-server")
    mail.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}
    tool = tools["reply_email"]
    params_schema = _inner_params_schema(tool)

    body_type_schema = params_schema["properties"]["body_type"]
    assert body_type_schema["default"] == "text"
    assert body_type_schema["deprecated"] is True
    assert "do not expose a content-type selector" in body_type_schema["description"]

    output_body_type_schema = tool.output_schema["properties"]["body_type"]
    assert output_body_type_schema["deprecated"] is True
    assert "compatibility echo" in output_body_type_schema["description"]


@pytest.mark.asyncio
async def test_filter_emails_schema_exposes_safe_recipient_search() -> None:
    mcp = FastMCP("test-server")
    mail.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}
    params_schema = _inner_params_schema(tools["filter_emails"])

    recipient_schema = params_schema["properties"]["to_address"]
    string_schema = next(
        variant
        for variant in recipient_schema["anyOf"]
        if variant.get("type") == "string"
    )
    assert string_schema["maxLength"] == 320
    assert "documented to: search" in recipient_schema["description"]
    assert params_schema["additionalProperties"] is False


@pytest.mark.asyncio
async def test_mail_search_schemas_default_to_entire_mailbox() -> None:
    mcp = FastMCP("test-server")
    mail.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}
    for tool_name in ("search_emails", "filter_emails"):
        folder_schema = _inner_params_schema(tools[tool_name])["properties"]["folder"]
        assert folder_schema["default"] is None
        assert "entire mailbox" in folder_schema["description"]

    output_folder_schema = tools["filter_emails"].output_schema["properties"][
        "folder"
    ]
    assert {variant["type"] for variant in output_folder_schema["anyOf"]} == {
        "string",
        "null",
    }
    assert output_folder_schema["default"] is None
    assert "entire mailbox" in output_folder_schema["description"]
    assert "Mail.ReadWrite" in tools["search_emails"].description
    assert "Mail.ReadWrite" in tools["filter_emails"].description


@pytest.mark.asyncio
async def test_create_reply_draft_schema_and_annotations() -> None:
    mcp = FastMCP("test-server")
    drafts.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}
    tool = tools["create_reply_draft"]
    params_schema = _inner_params_schema(tool)

    assert params_schema["required"] == ["message_id"]
    assert set(params_schema["properties"]) == {
        "message_id",
        "body",
        "reply_all",
        "body_type",
        "profile",
    }
    assert params_schema["properties"]["body"]["default"] is None
    assert params_schema["properties"]["reply_all"]["default"] is False
    assert set(tool.output_schema["properties"]) == {
        "success",
        "action",
        "error",
        "draft_id",
        "original_message_id",
        "reply_all",
        "body_type",
    }
    assert set(tool.output_schema["required"]) == {"success", "action"}
    assert "Mail.ReadWrite" in tool.description
    assert "not sent" in tool.description
    assert tool.annotations.destructiveHint is False
    assert tool.annotations.readOnlyHint is not True

    get_draft_tool = tools["get_draft"]
    assert set(get_draft_tool.output_schema["properties"]) == {
        "success",
        "action",
        "error",
        "id",
        "subject",
        "to",
        "cc",
        "bcc",
        "last_modified_at",
        "last_modified_at_display",
        "body",
        "body_content_type",
        "is_draft",
    }
    assert set(get_draft_tool.output_schema["required"]) == {"success", "action"}


@pytest.mark.asyncio
async def test_update_draft_schema_documents_history_preservation() -> None:
    mcp = FastMCP("test-server")
    drafts.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}
    tool = tools["update_draft"]
    params_schema = _inner_params_schema(tool)

    assert params_schema["required"] == ["draft_id"]
    preserve_schema = params_schema["properties"]["preserve_history"]
    assert preserve_schema["default"] is False
    assert "quoted reply history is retained" in preserve_schema["description"]
    assert tool.annotations.destructiveHint is False
    assert tool.annotations.openWorldHint is True
    assert "does not send" in tool.description
    assert "Mail.ReadWrite" in tool.description


@pytest.mark.asyncio
async def test_move_or_copy_item_tool_preserves_copy_field_name() -> None:
    mcp = FastMCP("test-server")
    onedrive.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}
    tool = tools["move_or_copy_item"]
    params_schema = _inner_params_schema(tool)

    assert "copy" in params_schema["properties"]
    assert "copy_value" not in params_schema["properties"]


@pytest.mark.asyncio
async def test_zero_arg_tools_do_not_require_empty_params_wrapper() -> None:
    mcp = FastMCP("test-server")
    profiles.register(mcp)
    services.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}

    assert tools["list_ms_profiles"].parameters == {
        "additionalProperties": False,
        "properties": {},
        "type": "object",
    }
    assert tools["list_enabled_services"].parameters == {
        "additionalProperties": False,
        "properties": {},
        "type": "object",
    }


@pytest.mark.asyncio
async def test_registered_tools_expose_human_friendly_titles() -> None:
    mcp = FastMCP("test-server")
    mail.register(mcp)
    profiles.register(mcp)
    sharepoint.register(mcp)

    tools = {tool.name: tool for tool in await mcp.list_tools(run_middleware=False)}

    assert tools["read_email"].title == "Read Email"
    assert tools["list_ms_profiles"].title == "List MS Profiles"
    assert tools["search_sharepoint_sites"].title == "Search SharePoint Sites"
