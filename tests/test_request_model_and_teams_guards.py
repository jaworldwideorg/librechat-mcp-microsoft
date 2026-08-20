from __future__ import annotations

import pytest
from pydantic import ValidationError

from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.text import strip_html
from mcp_microsoft.tools import teams


class _ExampleInput(ToolRequestModel):
    display_name: str
    recipients: list[str]
    options: dict[str, str]


def test_tool_request_model_normalizes_camel_case_and_json_text() -> None:
    params = _ExampleInput.model_validate(
        {
            "displayName": "Ada",
            "recipients": "alice@example.com, bob@example.com",
            "options": '{"mode":"strict"}',
        }
    )

    assert params.display_name == "Ada"
    assert params.recipients == ["alice@example.com", "bob@example.com"]
    assert params.options == {"mode": "strict"}


def test_strip_html_removes_tags_and_decodes_entities() -> None:
    assert strip_html("<p>Hello &amp; <strong>world</strong></p>") == "Hello & world"


def test_teams_list_chats_input_rejects_unknown_chat_type() -> None:
    with pytest.raises(ValidationError):
        teams.TeamsListChatsInput(chat_type="admin")


@pytest.mark.asyncio
async def test_teams_list_meetings_normalizes_datetime_filters(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, _path: str, params: dict | None = None):
            captured["params"] = params or {}
            return {"value": []}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())

    result = await teams.teams_list_meetings(
        teams.TeamsListMeetingsInput(
            start_after="2026-04-10T14:00:00-03:00",
            start_before="2026-04-10T15:00:00",
        )
    )

    assert result.count == 0
    assert captured["params"] == {
        "startDateTime": "2026-04-10T17:00:00Z",
        "endDateTime": "2026-04-10T15:00:00Z",
        "$top": 50,
        "$select": "id,subject,start,end,isOnlineMeeting,onlineMeeting,onlineMeetingUrl,createdDateTime",
    }


@pytest.mark.asyncio
async def test_teams_list_meetings_rejects_invalid_datetime_filter() -> None:
    with pytest.raises(ValueError, match="valid ISO 8601"):
        await teams.teams_list_meetings(
            teams.TeamsListMeetingsInput(
                start_after="2026-04-10T14:00:00Z' or 1 eq 1 or '",
            )
        )
