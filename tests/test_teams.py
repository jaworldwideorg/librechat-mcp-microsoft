"""
tests/test_teams.py — Test coverage for the Teams module.

Uses monkeypatch to mock the Graph API client, following the
pattern established in test_calendar.py.
"""

from __future__ import annotations

import json

import pytest
from pydantic import ValidationError

import mcp_microsoft.server as server
from mcp_microsoft.models import (
    ChannelDetailResponse,
    ChannelMessageDetailResponse,
    ChatDetailResponse,
    CreateChannelResponse,
    CreateTeamsChatResponse,
    CreateTeamsMeetingResponse,
    MeetingDetailResponse,
    SendTeamsMessageResponse,
    TeamDetailResponse,
    TeamsListChannelMessagesResponse,
    TeamsListChannelsResponse,
    TeamsListChatsResponse,
    TeamsListChatMessagesResponse,
    TeamsListJoinedResponse,
    TeamsListMeetingsResponse,
    TeamsListRepliesResponse,
)
from mcp_microsoft.tools import teams


# ---------------------------------------------------------------------------
# Tool Registration
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_tools_are_registered(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify all Teams tools are registered when MCP_ENABLE_TEAMS is set."""
    monkeypatch.setenv("MCP_ENABLE_TEAMS", "true")
    tool_names = {tool.name for tool in await server.mcp.list_tools(run_middleware=False)}
    assert "teams_list_joined" in tool_names
    assert "teams_get" in tool_names
    assert "teams_list_channels" in tool_names
    assert "teams_get_channel" in tool_names
    assert "teams_create_channel" in tool_names
    assert "teams_list_channel_messages" in tool_names
    assert "teams_get_channel_message" in tool_names
    assert "teams_send_channel_message" in tool_names
    assert "teams_reply_to_channel_message" in tool_names
    assert "teams_list_message_replies" in tool_names
    assert "teams_list_chats" in tool_names
    assert "teams_get_chat" in tool_names
    assert "teams_list_chat_messages" in tool_names
    assert "teams_send_chat_message" in tool_names
    assert "teams_create_chat" in tool_names
    assert "teams_create_meeting" in tool_names
    assert "teams_get_meeting" in tool_names
    assert "teams_list_meetings" in tool_names


# ---------------------------------------------------------------------------
# teams_list_joined
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_list_joined_returns_structured_response(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_list_joined returns TeamsListJoinedResponse with team summaries."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "value": [
                    {
                        "id": "team-1",
                        "displayName": "Engineering",
                        "description": "Engineering team",
                        "visibility": "private",
                        "webUrl": "https://teams.microsoft.com/l/team/team-1",
                        "isArchived": False,
                    },
                    {
                        "id": "team-2",
                        "displayName": "Marketing",
                        "description": "",
                        "visibility": "public",
                        "webUrl": "https://teams.microsoft.com/l/team/team-2",
                        "isArchived": None,
                    },
                ]
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_list_joined(teams.TeamsListJoinedInput())

    assert isinstance(result, TeamsListJoinedResponse)
    assert result.count == 2
    eng = next(t for t in result.teams if t.id == "team-1")
    assert eng.display_name == "Engineering"
    assert eng.visibility == "private"
    assert eng.is_archived is False


@pytest.mark.asyncio
async def test_teams_list_joined_empty_result(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_list_joined handles an empty value array."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {"value": []}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_list_joined(teams.TeamsListJoinedInput())

    assert result.count == 0
    assert result.teams == []


@pytest.mark.asyncio
async def test_teams_list_joined_uses_correct_path(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_list_joined calls /me/joinedTeams."""
    captured: dict[str, str] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            captured["path"] = path
            return {"value": []}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_list_joined(teams.TeamsListJoinedInput())

    assert captured["path"] == "/me/joinedTeams"


# ---------------------------------------------------------------------------
# teams_get
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_get_returns_team_detail(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_get returns TeamDetailResponse with settings."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "id": "team-abc",
                "displayName": "Platform",
                "description": "Platform engineering",
                "visibility": "private",
                "webUrl": "https://teams.microsoft.com/l/team/team-abc",
                "isArchived": False,
                "memberSettings": {
                    "allowCreateUpdateChannels": True,
                    "allowDeleteChannels": False,
                },
                "guestSettings": {
                    "allowCreateUpdateChannels": False,
                    "allowDeleteChannels": False,
                },
                "funSettings": {
                    "allowGiphy": True,
                    "giphyContentRating": "moderate",
                    "allowStickersAndMemes": True,
                    "allowCustomMemes": False,
                },
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_get(teams.TeamsGetInput(team_id="team-abc"))

    assert isinstance(result, TeamDetailResponse)
    assert result.id == "team-abc"
    assert result.display_name == "Platform"
    assert result.member_settings.allow_create_update_channels is True
    assert result.member_settings.allow_delete_channels is False
    assert result.fun_settings.allow_giphy is True
    assert result.fun_settings.giphy_content_rating == "moderate"


@pytest.mark.asyncio
async def test_teams_get_uses_team_id_in_path(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_get constructs /teams/{team_id}."""
    captured: dict[str, str] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            captured["path"] = path
            return {"id": "team-xyz", "memberSettings": {}, "guestSettings": {}, "funSettings": {}}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_get(teams.TeamsGetInput(team_id="team-xyz"))

    assert captured["path"] == "/teams/team-xyz"


# ---------------------------------------------------------------------------
# teams_list_channels
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_list_channels_returns_structured_response(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_list_channels returns TeamsListChannelsResponse."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "value": [
                    {
                        "id": "chan-1",
                        "displayName": "General",
                        "description": "General channel",
                        "channelType": "standard",
                        "webUrl": "https://teams.microsoft.com/l/channel/chan-1",
                        "isFavoriteByDefault": True,
                    }
                ]
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_list_channels(teams.TeamsListChannelsInput(team_id="team-1"))

    assert isinstance(result, TeamsListChannelsResponse)
    assert result.team_id == "team-1"
    assert result.count == 1
    chan = result.channels[0]
    assert chan.display_name == "General"
    assert chan.channel_type == "standard"
    assert chan.is_favorite_by_default is True


@pytest.mark.asyncio
async def test_teams_list_channels_uses_team_id_in_path(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_list_channels builds the correct API path."""
    captured: dict[str, str] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            captured["path"] = path
            return {"value": []}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_list_channels(teams.TeamsListChannelsInput(team_id="team-99"))

    assert captured["path"] == "/teams/team-99/channels"


# ---------------------------------------------------------------------------
# teams_get_channel
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_get_channel_returns_channel_detail(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_get_channel returns ChannelDetailResponse."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "id": "chan-42",
                "displayName": "Announcements",
                "description": "Company announcements",
                "channelType": "standard",
                "webUrl": "https://teams.microsoft.com/l/channel/chan-42",
                "isFavoriteByDefault": False,
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_get_channel(
        teams.TeamsGetChannelInput(team_id="team-1", channel_id="chan-42")
    )

    assert isinstance(result, ChannelDetailResponse)
    assert result.id == "chan-42"
    assert result.display_name == "Announcements"


@pytest.mark.asyncio
async def test_teams_get_channel_builds_correct_path(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_get_channel uses /teams/{team_id}/channels/{channel_id}."""
    captured: dict[str, str] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            captured["path"] = path
            return {"id": "chan-x", "channelType": "standard"}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_get_channel(
        teams.TeamsGetChannelInput(team_id="t1", channel_id="c1")
    )

    assert captured["path"] == "/teams/t1/channels/c1"


# ---------------------------------------------------------------------------
# teams_create_channel
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_create_channel_dry_run_no_api_call(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_create_channel returns dry-run preview when confirm=False."""
    api_called = {"called": False}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            api_called["called"] = True
            return {}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_create_channel(
        teams.TeamsCreateChannelInput(
            team_id="team-1",
            display_name="New Channel",
            description="A new channel",
            confirm=False,
        )
    )

    assert isinstance(result, CreateChannelResponse)
    assert result.success is True
    assert result.dry_run is True
    assert result.requires_confirmation is True
    assert result.display_name == "New Channel"
    assert api_called["called"] is False


@pytest.mark.asyncio
async def test_teams_create_channel_confirmed_posts_to_api(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_create_channel creates the channel when confirm=True."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["path"] = path
            captured["json"] = json
            return {
                "id": "new-chan-id",
                "displayName": "New Channel",
                "description": "A new channel",
                "webUrl": "https://teams.microsoft.com/l/channel/new-chan-id",
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_create_channel(
        teams.TeamsCreateChannelInput(
            team_id="team-1",
            display_name="New Channel",
            description="A new channel",
            confirm=True,
        )
    )

    assert isinstance(result, CreateChannelResponse)
    assert result.success is True
    assert result.dry_run is False
    assert result.channel_id == "new-chan-id"
    assert captured["path"] == "/teams/team-1/channels"
    assert captured["json"]["displayName"] == "New Channel"
    assert captured["json"]["channelType"] == "standard"
    assert captured["json"]["description"] == "A new channel"


@pytest.mark.asyncio
async def test_teams_create_channel_no_description_in_payload(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_create_channel omits description when not provided."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["json"] = json
            return {"id": "chan-id"}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_create_channel(
        teams.TeamsCreateChannelInput(
            team_id="team-1",
            display_name="Minimal",
            confirm=True,
        )
    )

    assert "description" not in captured["json"]


# ---------------------------------------------------------------------------
# teams_list_channel_messages
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_list_channel_messages_returns_messages(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_list_channel_messages returns TeamsListChannelMessagesResponse."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "value": [
                    {
                        "id": "msg-1",
                        "createdDateTime": "2026-04-01T09:00:00Z",
                        "lastModifiedDateTime": "2026-04-01T09:00:00Z",
                        "from": {
                            "user": {
                                "id": "user-1",
                                "displayName": "Alice",
                            }
                        },
                        "body": {
                            "contentType": "text",
                            "content": "Hello team!",
                        },
                        "subject": "",
                        "webUrl": "https://teams.microsoft.com/l/message/msg-1",
                        "replyToId": "",
                        "importance": "normal",
                    }
                ]
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_list_channel_messages(
        teams.TeamsListChannelMessagesInput(team_id="team-1", channel_id="chan-1")
    )

    assert isinstance(result, TeamsListChannelMessagesResponse)
    assert result.team_id == "team-1"
    assert result.channel_id == "chan-1"
    assert result.count == 1
    msg = result.messages[0]
    assert msg.id == "msg-1"
    assert msg.body == "Hello team!"
    assert msg.from_display == "Alice"


@pytest.mark.asyncio
async def test_teams_list_channel_messages_truncates_long_body(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_list_channel_messages truncates body content over 500 chars."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "value": [
                    {
                        "id": "msg-long",
                        "createdDateTime": "2026-04-01T10:00:00Z",
                        "lastModifiedDateTime": "2026-04-01T10:00:00Z",
                        "from": None,
                        "body": {
                            "contentType": "text",
                            "content": "A" * 600,
                        },
                        "subject": "",
                        "webUrl": "",
                        "replyToId": "",
                        "importance": "",
                    }
                ]
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_list_channel_messages(
        teams.TeamsListChannelMessagesInput(team_id="team-1", channel_id="chan-1")
    )

    msg = result.messages[0]
    assert len(msg.body) <= 504  # 500 + "…" (3 bytes)
    assert msg.body.endswith("…")


# ---------------------------------------------------------------------------
# teams_get_channel_message
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_get_channel_message_returns_detail(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_get_channel_message returns ChannelMessageDetailResponse with extras."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "id": "msg-42",
                "createdDateTime": "2026-04-01T10:00:00Z",
                "lastModifiedDateTime": "2026-04-01T10:00:00Z",
                "from": {
                    "user": {"id": "user-2", "displayName": "Bob"}
                },
                "body": {
                    "contentType": "html",
                    "content": "<p>Detailed message</p>",
                },
                "subject": "Important Update",
                "webUrl": "https://teams.microsoft.com/l/message/msg-42",
                "replyToId": "",
                "importance": "high",
                "reactions": [
                    {
                        "reactionType": "like",
                        "createdDateTime": "2026-04-01T10:05:00Z",
                        "user": {
                            "user": {"id": "user-3", "displayName": "Carol"}
                        },
                    }
                ],
                "attachments": [],
                "mentions": [],
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_get_channel_message(
        teams.TeamsGetChannelMessageInput(
            team_id="team-1", channel_id="chan-1", message_id="msg-42"
        )
    )

    assert isinstance(result, ChannelMessageDetailResponse)
    assert result.id == "msg-42"
    assert result.subject == "Important Update"
    assert result.importance == "high"
    assert result.from_display == "Bob"
    assert len(result.reactions) == 1
    assert result.reactions[0].reaction_type == "like"
    assert result.attachments == []
    assert result.mentions == []


# ---------------------------------------------------------------------------
# teams_send_channel_message
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_send_channel_message_posts_correct_payload(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_send_channel_message POSTs the correct body and returns success."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["path"] = path
            captured["json"] = json
            return {
                "id": "sent-msg-1",
                "createdDateTime": "2026-04-01T11:00:00Z",
                "webUrl": "https://teams.microsoft.com/l/message/sent-msg-1",
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_send_channel_message(
        teams.TeamsSendChannelMessageInput(
            team_id="team-1",
            channel_id="chan-1",
            content="Hello, channel!",
            content_type="text",
        )
    )

    assert isinstance(result, SendTeamsMessageResponse)
    assert result.success is True
    assert result.id == "sent-msg-1"
    assert captured["path"] == "/teams/team-1/channels/chan-1/messages"
    assert captured["json"]["body"]["content"] == "Hello, channel!"
    assert captured["json"]["body"]["contentType"] == "text"


@pytest.mark.asyncio
async def test_teams_send_channel_message_includes_subject(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_send_channel_message includes subject when provided."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["json"] = json
            return {"id": "msg-with-subject", "createdDateTime": "2026-04-01T11:00:00Z"}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_send_channel_message(
        teams.TeamsSendChannelMessageInput(
            team_id="team-1",
            channel_id="chan-1",
            content="Announcement body",
            subject="Important Announcement",
        )
    )

    assert captured["json"]["subject"] == "Important Announcement"


@pytest.mark.asyncio
async def test_teams_send_channel_message_no_subject_when_omitted(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_send_channel_message omits subject key when not provided."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["json"] = json
            return {"id": "msg-no-subject", "createdDateTime": "2026-04-01T11:00:00Z"}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_send_channel_message(
        teams.TeamsSendChannelMessageInput(
            team_id="team-1",
            channel_id="chan-1",
            content="Plain text",
        )
    )

    assert "subject" not in captured["json"]


# ---------------------------------------------------------------------------
# teams_reply_to_channel_message
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_reply_to_channel_message_posts_to_replies_endpoint(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_reply_to_channel_message POSTs to /replies and sets parent_message_id."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["path"] = path
            captured["json"] = json
            return {
                "id": "reply-1",
                "createdDateTime": "2026-04-01T12:00:00Z",
                "webUrl": "https://teams.microsoft.com/l/message/reply-1",
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_reply_to_channel_message(
        teams.TeamsReplyToChannelMessageInput(
            team_id="team-1",
            channel_id="chan-1",
            message_id="msg-root",
            content="This is my reply",
        )
    )

    assert isinstance(result, SendTeamsMessageResponse)
    assert result.success is True
    assert result.id == "reply-1"
    assert result.parent_message_id == "msg-root"
    assert captured["path"] == "/teams/team-1/channels/chan-1/messages/msg-root/replies"
    assert captured["json"]["body"]["content"] == "This is my reply"


# ---------------------------------------------------------------------------
# teams_list_message_replies
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_list_message_replies_returns_replies(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_list_message_replies returns TeamsListRepliesResponse."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "value": [
                    {
                        "id": "reply-a",
                        "createdDateTime": "2026-04-01T12:30:00Z",
                        "lastModifiedDateTime": "2026-04-01T12:30:00Z",
                        "from": None,
                        "body": {"contentType": "text", "content": "Reply A"},
                        "subject": "",
                        "webUrl": "",
                        "replyToId": "msg-root",
                        "importance": "normal",
                    }
                ]
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_list_message_replies(
        teams.TeamsListMessageRepliesInput(
            team_id="team-1",
            channel_id="chan-1",
            message_id="msg-root",
        )
    )

    assert isinstance(result, TeamsListRepliesResponse)
    assert result.team_id == "team-1"
    assert result.channel_id == "chan-1"
    assert result.parent_message_id == "msg-root"
    assert result.count == 1
    assert result.replies[0].id == "reply-a"


@pytest.mark.asyncio
async def test_teams_list_message_replies_uses_correct_path(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_list_message_replies calls the replies sub-resource."""
    captured: dict[str, str] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            captured["path"] = path
            return {"value": []}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_list_message_replies(
        teams.TeamsListMessageRepliesInput(
            team_id="team-T",
            channel_id="chan-C",
            message_id="msg-M",
        )
    )

    assert captured["path"] == "/teams/team-T/channels/chan-C/messages/msg-M/replies"


# ---------------------------------------------------------------------------
# teams_list_chats
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_list_chats_returns_structured_response(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_list_chats returns TeamsListChatsResponse."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "value": [
                    {
                        "id": "chat-1",
                        "chatType": "oneOnOne",
                        "topic": "",
                        "createdDateTime": "2026-03-01T08:00:00Z",
                        "lastUpdatedDateTime": "2026-04-01T09:00:00Z",
                        "webUrl": "https://teams.microsoft.com/l/chat/chat-1",
                    },
                    {
                        "id": "chat-2",
                        "chatType": "group",
                        "topic": "Project Alpha",
                        "createdDateTime": "2026-03-15T10:00:00Z",
                        "lastUpdatedDateTime": "2026-04-01T10:30:00Z",
                        "webUrl": "https://teams.microsoft.com/l/chat/chat-2",
                    },
                ]
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_list_chats(teams.TeamsListChatsInput())

    assert isinstance(result, TeamsListChatsResponse)
    assert result.count == 2
    group_chat = next(c for c in result.chats if c.id == "chat-2")
    assert group_chat.topic == "Project Alpha"
    assert group_chat.chat_type == "group"


@pytest.mark.asyncio
async def test_teams_list_chats_filter_by_type(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_list_chats adds $filter param when chat_type is provided."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            captured["params"] = params or {}
            return {"value": []}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_list_chats(teams.TeamsListChatsInput(chat_type="oneOnOne"))

    assert "$filter" in captured["params"]
    assert "oneOnOne" in captured["params"]["$filter"]


@pytest.mark.asyncio
async def test_teams_list_chats_no_filter_when_type_empty(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_list_chats omits $filter when chat_type is not set."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            captured["params"] = params or {}
            return {"value": []}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_list_chats(teams.TeamsListChatsInput())

    assert "$filter" not in captured["params"]


# ---------------------------------------------------------------------------
# teams_get_chat
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_get_chat_returns_chat_with_members(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_get_chat returns ChatDetailResponse with member list."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "id": "chat-xyz",
                "chatType": "group",
                "topic": "Design Review",
                "createdDateTime": "2026-03-20T09:00:00Z",
                "lastUpdatedDateTime": "2026-04-01T14:00:00Z",
                "webUrl": "https://teams.microsoft.com/l/chat/chat-xyz",
                "members": [
                    {
                        "id": "member-1",
                        "displayName": "Alice",
                        "email": "alice@example.com",
                        "userId": "user-alice",
                        "tenantId": "tenant-1",
                        "roles": ["owner"],
                    },
                    {
                        "id": "member-2",
                        "displayName": "Bob",
                        "email": "bob@example.com",
                        "userId": "user-bob",
                        "tenantId": "tenant-1",
                        "roles": [],
                    },
                ],
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_get_chat(teams.TeamsGetChatInput(chat_id="chat-xyz"))

    assert isinstance(result, ChatDetailResponse)
    assert result.id == "chat-xyz"
    assert result.topic == "Design Review"
    assert len(result.members) == 2
    alice = next(m for m in result.members if m.display_name == "Alice")
    assert alice.email == "alice@example.com"
    assert "owner" in alice.roles


@pytest.mark.asyncio
async def test_teams_get_chat_expands_members(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_get_chat requests $expand=members."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            captured["params"] = params or {}
            return {
                "id": "chat-1",
                "chatType": "oneOnOne",
                "topic": "",
                "createdDateTime": "2026-01-01T00:00:00Z",
                "lastUpdatedDateTime": "2026-01-01T00:00:00Z",
                "webUrl": "",
                "members": [],
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_get_chat(teams.TeamsGetChatInput(chat_id="chat-1"))

    assert captured["params"].get("$expand") == "members"


# ---------------------------------------------------------------------------
# teams_list_chat_messages
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_list_chat_messages_returns_messages(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_list_chat_messages returns TeamsListChatMessagesResponse."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "value": [
                    {
                        "id": "chat-msg-1",
                        "createdDateTime": "2026-04-01T15:00:00Z",
                        "lastModifiedDateTime": "2026-04-01T15:00:00Z",
                        "from": {
                            "user": {"id": "user-1", "displayName": "Carol"}
                        },
                        "body": {"contentType": "text", "content": "Hey, how's it going?"},
                        "subject": "",
                        "webUrl": "https://teams.microsoft.com/l/message/chat-msg-1",
                        "replyToId": "",
                        "importance": "normal",
                    }
                ]
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_list_chat_messages(
        teams.TeamsListChatMessagesInput(chat_id="chat-abc")
    )

    assert isinstance(result, TeamsListChatMessagesResponse)
    assert result.chat_id == "chat-abc"
    assert result.count == 1
    assert result.messages[0].from_display == "Carol"


@pytest.mark.asyncio
async def test_teams_list_chat_messages_uses_chat_id_in_path(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_list_chat_messages calls /me/chats/{chat_id}/messages."""
    captured: dict[str, str] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            captured["path"] = path
            return {"value": []}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_list_chat_messages(
        teams.TeamsListChatMessagesInput(chat_id="chat-999")
    )

    assert captured["path"] == "/me/chats/chat-999/messages"


# ---------------------------------------------------------------------------
# teams_send_chat_message
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_send_chat_message_posts_correct_payload(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_send_chat_message POSTs to the chat and returns SendTeamsMessageResponse."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["path"] = path
            captured["json"] = json
            return {
                "id": "chat-sent-1",
                "createdDateTime": "2026-04-01T16:00:00Z",
                "webUrl": "https://teams.microsoft.com/l/message/chat-sent-1",
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_send_chat_message(
        teams.TeamsSendChatMessageInput(
            chat_id="chat-abc",
            content="Hello chat!",
            content_type="text",
        )
    )

    assert isinstance(result, SendTeamsMessageResponse)
    assert result.success is True
    assert result.id == "chat-sent-1"
    assert result.chat_id == "chat-abc"
    assert captured["path"] == "/me/chats/chat-abc/messages"
    assert captured["json"]["body"]["content"] == "Hello chat!"
    assert captured["json"]["body"]["contentType"] == "text"


@pytest.mark.asyncio
async def test_teams_send_chat_message_html_content(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_send_chat_message accepts html content_type."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["json"] = json
            return {"id": "html-msg", "createdDateTime": "2026-04-01T16:30:00Z"}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_send_chat_message(
        teams.TeamsSendChatMessageInput(
            chat_id="chat-abc",
            content="<b>Bold message</b>",
            content_type="html",
        )
    )

    assert captured["json"]["body"]["contentType"] == "html"


# ---------------------------------------------------------------------------
# teams_create_chat
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_create_chat_group_includes_topic(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_create_chat creates a group chat with topic."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {"id": "me-user-id"}

        async def post(self, path: str, json: dict = None):
            captured["path"] = path
            captured["json"] = json
            return {
                "id": "new-chat-id",
                "chatType": "group",
                "topic": "Project Beta",
                "createdDateTime": "2026-04-01T17:00:00Z",
                "lastUpdatedDateTime": "2026-04-01T17:00:00Z",
                "webUrl": "https://teams.microsoft.com/l/chat/new-chat-id",
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_create_chat(
        teams.TeamsCreateChatInput(
            members=["alice@example.com", "bob@example.com"],
            topic="Project Beta",
            chat_type="group",
        )
    )

    assert isinstance(result, CreateTeamsChatResponse)
    assert result.success is True
    assert result.id == "new-chat-id"
    assert result.topic == "Project Beta"
    assert captured["json"]["chatType"] == "group"
    assert captured["json"]["topic"] == "Project Beta"


@pytest.mark.asyncio
async def test_teams_create_chat_adds_me_as_owner(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_create_chat fetches /me and adds the caller as owner."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {"id": "my-user-id"}

        async def post(self, path: str, json: dict = None):
            captured["members"] = json.get("members", [])
            return {
                "id": "chat-new",
                "chatType": "group",
                "topic": "",
                "createdDateTime": "2026-04-01T17:00:00Z",
                "lastUpdatedDateTime": "2026-04-01T17:00:00Z",
                "webUrl": "",
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_create_chat(
        teams.TeamsCreateChatInput(
            members=["alice@example.com"],
            chat_type="oneOnOne",
        )
    )

    members = captured["members"]
    owner_members = [m for m in members if "owner" in m.get("roles", [])]
    assert len(owner_members) == 1
    assert "my-user-id" in owner_members[0]["user@odata.bind"]


# ---------------------------------------------------------------------------
# teams_create_meeting
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_create_meeting_returns_meeting_response(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_create_meeting posts correct payload and returns CreateTeamsMeetingResponse."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["path"] = path
            captured["json"] = json
            return {
                "id": "meeting-id-1",
                "subject": "Quarterly Review",
                "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/meeting-id-1",
                "startDateTime": "2026-04-10T14:00:00Z",
                "endDateTime": "2026-04-10T15:00:00Z",
                "createdDateTime": "2026-04-01T08:00:00Z",
                "joinMeetingIdSettings": {"joinMeetingId": "987654321"},
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_create_meeting(
        teams.TeamsCreateMeetingInput(
            subject="Quarterly Review",
            start_datetime="2026-04-10T14:00:00Z",
            end_datetime="2026-04-10T15:00:00Z",
        )
    )

    assert isinstance(result, CreateTeamsMeetingResponse)
    assert result.success is True
    assert result.id == "meeting-id-1"
    assert result.subject == "Quarterly Review"
    assert "teams.microsoft.com" in result.join_web_url
    assert captured["path"] == "/me/onlineMeetings"
    assert captured["json"]["subject"] == "Quarterly Review"
    assert captured["json"]["startDateTime"] == "2026-04-10T14:00:00Z"


@pytest.mark.asyncio
async def test_teams_create_meeting_with_attendees(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_create_meeting includes participants when attendees are provided."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["json"] = json
            return {
                "id": "meeting-with-attendees",
                "subject": "Team Sync",
                "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/abc",
                "startDateTime": "2026-04-15T09:00:00Z",
                "endDateTime": "2026-04-15T10:00:00Z",
                "createdDateTime": "2026-04-01T08:00:00Z",
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_create_meeting(
        teams.TeamsCreateMeetingInput(
            subject="Team Sync",
            start_datetime="2026-04-15T09:00:00Z",
            end_datetime="2026-04-15T10:00:00Z",
            attendees=["alice@example.com", "bob@example.com"],
        )
    )

    payload = captured["json"]
    assert "participants" in payload
    attendees = payload["participants"]["attendees"]
    assert len(attendees) == 2
    upns = [a["upn"] for a in attendees]
    assert "alice@example.com" in upns
    assert "bob@example.com" in upns


@pytest.mark.asyncio
async def test_teams_create_meeting_no_attendees_no_participants_key(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_create_meeting omits participants key when no attendees given."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict = None):
            captured["json"] = json
            return {
                "id": "meeting-no-att",
                "subject": "Solo",
                "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/solo",
                "startDateTime": "2026-04-20T09:00:00Z",
                "endDateTime": "2026-04-20T10:00:00Z",
                "createdDateTime": "2026-04-01T08:00:00Z",
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_create_meeting(
        teams.TeamsCreateMeetingInput(
            subject="Solo",
            start_datetime="2026-04-20T09:00:00Z",
            end_datetime="2026-04-20T10:00:00Z",
        )
    )

    assert "participants" not in captured["json"]


# ---------------------------------------------------------------------------
# teams_get_meeting
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_get_meeting_returns_meeting_detail(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_get_meeting returns MeetingDetailResponse with participants."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {
                "id": "meeting-detail-1",
                "subject": "Board Meeting",
                "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/board",
                "startDateTime": "2026-04-20T13:00:00Z",
                "endDateTime": "2026-04-20T14:00:00Z",
                "createdDateTime": "2026-04-01T08:00:00Z",
                "joinMeetingIdSettings": {"joinMeetingId": "111222333"},
                "videoTeleconferenceId": "vtc-id-1",
                "participants": {
                    "organizer": {
                        "upn": "alice@example.com",
                        "role": "presenter",
                        "identity": {
                            "user": {"id": "user-1", "displayName": "Alice"}
                        },
                    },
                    "attendees": [],
                    "producers": [],
                    "contributors": [],
                },
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_get_meeting(
        teams.TeamsGetMeetingInput(meeting_id="meeting-detail-1")
    )

    assert isinstance(result, MeetingDetailResponse)
    assert result.id == "meeting-detail-1"
    assert result.subject == "Board Meeting"
    assert result.join_meeting_id == "111222333"
    assert result.video_teleconference_id == "vtc-id-1"
    assert result.participants.organizer is not None
    assert result.participants.organizer.upn == "alice@example.com"


@pytest.mark.asyncio
async def test_teams_get_meeting_uses_correct_path(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_get_meeting calls /me/onlineMeetings/{meeting_id}."""
    captured: dict[str, str] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            captured["path"] = path
            return {
                "id": "mtg-99",
                "subject": "",
                "joinWebUrl": "",
                "startDateTime": "2026-04-01T10:00:00Z",
                "endDateTime": "2026-04-01T11:00:00Z",
                "createdDateTime": "",
                "participants": {"attendees": [], "producers": [], "contributors": []},
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_get_meeting(teams.TeamsGetMeetingInput(meeting_id="mtg-99"))

    assert captured["path"] == "/me/onlineMeetings/mtg-99"


# ---------------------------------------------------------------------------
# teams_list_meetings
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_teams_list_meetings_returns_meetings(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_list_meetings returns TeamsListMeetingsResponse with meetings."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            if path == "/me/calendarView":
                return {
                    "value": [{
                        "id": "event-1",
                        "subject": "Daily Standup",
                        "isOnlineMeeting": True,
                        "onlineMeeting": {"joinUrl": "https://teams.microsoft.com/l/meetup-join/standup"},
                    }]
                }
            return {
                "value": [
                    {
                        "id": "mtg-1",
                        "subject": "Daily Standup",
                        "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/standup",
                        "startDateTime": "2026-04-02T09:00:00Z",
                        "endDateTime": "2026-04-02T09:15:00Z",
                        "createdDateTime": "2026-04-01T08:00:00Z",
                    }
                ]
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_list_meetings(teams.TeamsListMeetingsInput())

    assert isinstance(result, TeamsListMeetingsResponse)
    assert result.count == 1
    assert result.meetings[0].subject == "Daily Standup"
    assert result.start_after != ""
    assert result.start_before != ""


@pytest.mark.asyncio
async def test_teams_list_meetings_uses_calendar_view_range(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify meeting enumeration uses supported calendarView parameters."""
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            captured["path"] = path
            captured["params"] = params or {}
            return {"value": []}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_list_meetings(
        teams.TeamsListMeetingsInput(
            start_after="2026-04-01T00:00:00Z",
            start_before="2026-04-30T00:00:00Z",
        )
    )

    assert captured["path"] == "/me/calendarView"
    assert captured["params"]["startDateTime"] == "2026-04-01T00:00:00Z"
    assert captured["params"]["endDateTime"] == "2026-04-30T00:00:00Z"
    assert "$filter" not in captured["params"]


@pytest.mark.asyncio
async def test_teams_list_meetings_empty_result(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify teams_list_meetings handles no meetings gracefully."""

    class DummyGraph:
        async def get(self, path: str, params: dict = None):
            return {"value": []}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_list_meetings(teams.TeamsListMeetingsInput())

    assert result.count == 0
    assert result.meetings == []


# ---------------------------------------------------------------------------
# ToolRequestModel input validation
# ---------------------------------------------------------------------------


def test_teams_input_camelcase_normalization() -> None:
    """Verify ToolRequestModel accepts camelCase keys and normalises to snake_case."""
    params = teams.TeamsListChannelsInput.model_validate(
        {"teamId": "team-abc", "top": 25}
    )
    assert params.team_id == "team-abc"
    assert params.top == 25


def test_teams_input_stringified_json_payload() -> None:
    """Verify ToolRequestModel parses a JSON-encoded string as the full payload."""
    raw = json.dumps({"teamId": "team-xyz", "channelId": "chan-xyz", "messageId": "msg-xyz"})
    params = teams.TeamsGetChannelMessageInput.model_validate(raw)
    assert params.team_id == "team-xyz"
    assert params.channel_id == "chan-xyz"
    assert params.message_id == "msg-xyz"


def test_teams_create_chat_members_comma_separated_string() -> None:
    """Verify TeamsCreateChatInput coerces a comma-separated members string into a list."""
    params = teams.TeamsCreateChatInput.model_validate(
        {
            "members": "alice@example.com, bob@example.com, carol@example.com",
            "chat_type": "group",
        }
    )
    assert isinstance(params.members, list)
    assert len(params.members) == 3
    assert "alice@example.com" in params.members


def test_teams_list_chats_rejects_unknown_chat_type() -> None:
    """Verify TeamsListChatsInput raises ValidationError for invalid chat_type."""
    with pytest.raises(ValidationError):
        teams.TeamsListChatsInput(chat_type="unknown_type")


def test_teams_get_input_required_field() -> None:
    """Verify TeamsGetInput raises ValidationError when team_id is missing."""
    with pytest.raises(ValidationError):
        teams.TeamsGetInput.model_validate({})


def test_teams_send_channel_message_required_fields() -> None:
    """Verify TeamsSendChannelMessageInput raises ValidationError when required fields missing."""
    with pytest.raises(ValidationError):
        teams.TeamsSendChannelMessageInput.model_validate(
            {"team_id": "team-1"}  # missing channel_id and content
        )
