from __future__ import annotations

import pytest

from mcp_microsoft.tools import calendar, mail, sharepoint, teams


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("search_input", "expected_graph_query"),
    [
        ("project update", '"project update"'),
        (
            '"Denmark" AND ("LMS" OR "deployment")',
            '"Denmark" AND ("LMS" OR "deployment")',
        ),
        ('subject:"quarterly report"', 'subject:"quarterly report"'),
    ],
)
async def test_search_emails_does_not_double_quote_kql(
    monkeypatch: pytest.MonkeyPatch,
    search_input: str,
    expected_graph_query: str,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            captured["path"] = path
            captured["params"] = params or {}
            return {"value": []}

    monkeypatch.setattr(mail, "get_graph", lambda _profile: DummyGraph())
    await mail.search_emails(mail.SearchEmailsInput(query=search_input))

    assert captured["path"] == "/me/messages"
    assert captured["params"]["$search"] == expected_graph_query


@pytest.mark.asyncio
async def test_mail_orderby_property_leads_filter(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: list[dict[str, object]] = []

    class DummyGraph:
        async def get(self, _path: str, params: dict | None = None):
            captured.append(params or {})
            return {"value": []}

    monkeypatch.setattr(mail, "get_graph", lambda _profile: DummyGraph())
    await mail.list_emails(mail.ListEmailsInput(unread_only=True))
    await mail.filter_emails(mail.FilterEmailsInput(subject_contains="status"))

    for query in captured:
        assert query["$orderby"] == "receivedDateTime desc"
        assert str(query["$filter"]).startswith("receivedDateTime ge ")


@pytest.mark.asyncio
async def test_teams_list_chats_uses_supported_orderby(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            captured["path"] = path
            captured["params"] = params or {}
            return {"value": []}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_list_chats(teams.TeamsListChatsInput())

    assert captured["path"] == "/me/chats"
    assert captured["params"]["$orderby"] == "lastMessagePreview/createdDateTime desc"
    assert "$select" not in captured["params"]


@pytest.mark.asyncio
async def test_joined_teams_omits_unsupported_query_options(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            captured["path"] = path
            captured["params"] = params
            return {"value": []}

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    await teams.teams_list_joined(teams.TeamsListJoinedInput())

    assert captured == {"path": "/me/joinedTeams", "params": None}


@pytest.mark.asyncio
async def test_list_meetings_follows_calendar_pages_and_resolves_join_url(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls: list[tuple[str, dict | None]] = []

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            calls.append((path, params))
            if path == "/me/calendarView" and params and "$skiptoken" not in params:
                return {
                    "value": [{"id": "event-1", "isOnlineMeeting": False}],
                    "@odata.nextLink": (
                        "https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=next"
                    ),
                }
            if path == "/me/calendarView":
                return {
                    "value": [
                        {
                            "id": "event-2",
                            "isOnlineMeeting": True,
                            "onlineMeeting": {
                                "joinUrl": "https://teams.example/join/1"
                            },
                        }
                    ]
                }
            return {
                "value": [
                    {
                        "id": "meeting-1",
                        "subject": "Standup",
                        "joinWebUrl": "https://teams.example/join/1",
                        "startDateTime": "2026-04-01T09:00:00Z",
                        "endDateTime": "2026-04-01T09:15:00Z",
                    }
                ]
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_list_meetings(
        teams.TeamsListMeetingsInput(
            start_after="2026-04-01T00:00:00Z",
            start_before="2026-04-02T00:00:00Z",
        )
    )

    assert result.count == 1
    assert result.meetings[0].id == "meeting-1"
    assert calls[1] == ("/me/calendarView", {"$skiptoken": "next"})
    assert calls[2][1] == {"$filter": "JoinWebUrl eq 'https://teams.example/join/1'"}


@pytest.mark.asyncio
async def test_default_calendar_view_uses_documented_path(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            captured["path"] = path
            return {"value": []}

    monkeypatch.setattr(calendar, "get_graph", lambda _profile: DummyGraph())
    await calendar.list_upcoming_events(
        calendar.ListUpcomingEventsInput(
            start_datetime="2026-04-01T00:00:00Z",
            end_datetime="2026-04-02T00:00:00Z",
        )
    )

    assert captured["path"] == "/me/calendarView"


@pytest.mark.asyncio
async def test_search_api_isolates_restricted_entity_types(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    with pytest.raises(ValueError, match="searched separately"):
        await sharepoint.search_content(
            sharepoint.SearchContentInput(
                query="status", entity_types=["driveItem", "message"]
            )
        )

    captured: dict[str, object] = {}

    class DummyGraph:
        async def post(self, path: str, json: dict | None = None):
            captured["request"] = (json or {})["requests"][0]
            return {"value": []}

    monkeypatch.setattr(
        sharepoint, "_get_sharepoint_graph", lambda _profile: DummyGraph()
    )
    await sharepoint.search_content(
        sharepoint.SearchContentInput(
            query="status", entity_types=["message"], max_results=100
        )
    )

    assert captured["request"]["size"] == 25
    assert "fields" not in captured["request"]
