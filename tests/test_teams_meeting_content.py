from __future__ import annotations

from pathlib import Path

import httpx
import pytest

import mcp_microsoft.server as server
from mcp_microsoft.config import reset_app_config
from mcp_microsoft.profiles import build_default_scopes
from mcp_microsoft.runtime import reset_runtime_state
from mcp_microsoft.tools import teams


@pytest.fixture(autouse=True)
def reset_runtime() -> None:
    reset_runtime_state()
    yield
    reset_runtime_state()


def test_build_default_scopes_only_adds_meeting_artifact_scopes_when_explicitly_enabled(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("MCP_ENABLE_TEAMS", "true")
    monkeypatch.delenv("MCP_ENABLE_TEAMS_MEETING_ARTIFACTS", raising=False)
    monkeypatch.delenv("MCP_ENABLE_TEAMS_AI_INSIGHTS", raising=False)
    reset_app_config()

    scopes = build_default_scopes()
    assert "OnlineMeetingTranscript.Read.All" not in scopes
    assert "OnlineMeetingRecording.Read.All" not in scopes
    assert "OnlineMeetingAiInsight.Read.All" not in scopes

    monkeypatch.setenv("MCP_ENABLE_TEAMS_MEETING_ARTIFACTS", "true")
    monkeypatch.setenv("MCP_ENABLE_TEAMS_AI_INSIGHTS", "true")
    reset_app_config()

    scopes = build_default_scopes()
    assert "OnlineMeetingTranscript.Read.All" in scopes
    assert "OnlineMeetingRecording.Read.All" in scopes
    assert "OnlineMeetingAiInsight.Read.All" in scopes


def test_build_default_scopes_does_not_add_meeting_content_scopes_without_teams(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("MCP_ENABLE_TEAMS", "false")
    monkeypatch.setenv("MCP_ENABLE_TEAMS_MEETING_ARTIFACTS", "true")
    monkeypatch.setenv("MCP_ENABLE_TEAMS_AI_INSIGHTS", "true")
    reset_app_config()

    scopes = build_default_scopes()
    assert "OnlineMeetings.ReadWrite" not in scopes
    assert "OnlineMeetingTranscript.Read.All" not in scopes
    assert "OnlineMeetingRecording.Read.All" not in scopes
    assert "OnlineMeetingAiInsight.Read.All" not in scopes


@pytest.mark.asyncio
async def test_meeting_content_tools_register_only_under_explicit_flags(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("MCP_ENABLE_TEAMS", "true")
    tool_names = {
        tool.name
        for tool in await server.get_mcp_server(reset=True).list_tools(run_middleware=False)
    }

    assert "teams_find_meeting_by_url" in tool_names
    assert "teams_list_meeting_transcripts" not in tool_names
    assert "teams_get_meeting_ai_insight" not in tool_names

    monkeypatch.setenv("MCP_ENABLE_TEAMS_MEETING_ARTIFACTS", "true")
    monkeypatch.setenv("MCP_ENABLE_TEAMS_AI_INSIGHTS", "true")
    reset_runtime_state()
    tool_names = {
        tool.name
        for tool in await server.get_mcp_server(reset=True).list_tools(run_middleware=False)
    }

    assert "teams_list_meeting_transcripts" in tool_names
    assert "teams_get_meeting_transcript" in tool_names
    assert "teams_list_meeting_recordings" in tool_names
    assert "teams_download_meeting_recording" in tool_names
    assert "teams_list_meeting_ai_insights" in tool_names
    assert "teams_get_meeting_ai_insight" in tool_names


@pytest.mark.asyncio
async def test_teams_find_meeting_by_url_escapes_filter_and_returns_detail(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            if path == "/me/onlineMeetings":
                captured["path"] = path
                captured["params"] = params
                return {"value": [{"id": "meeting-1"}]}
            return {
                "id": "meeting-1",
                "subject": "Planner's Sync",
                "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/abc?context=it's",
                "startDateTime": "2026-04-10T14:00:00Z",
                "endDateTime": "2026-04-10T15:00:00Z",
                "createdDateTime": "2026-04-01T10:00:00Z",
                "participants": {"attendees": [], "producers": [], "contributors": []},
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())
    result = await teams.teams_find_meeting_by_url(
        teams.TeamsFindMeetingByUrlInput(
            join_web_url="https://teams.microsoft.com/l/meetup-join/abc?context=it's"
        )
    )

    assert result.id == "meeting-1"
    assert captured["path"] == "/me/onlineMeetings"
    assert captured["params"]["$filter"] == "JoinWebUrl eq 'https://teams.microsoft.com/l/meetup-join/abc?context=it''s'"


@pytest.mark.asyncio
async def test_teams_list_and_get_meeting_transcript(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            if path.endswith("/transcripts"):
                return {
                    "value": [
                        {
                            "id": "transcript-1",
                            "meetingId": "meeting-1",
                            "callId": "call-1",
                            "createdDateTime": "2026-04-10T15:00:00Z",
                            "endDateTime": "2026-04-10T15:30:00Z",
                            "contentCorrelationId": "corr-1",
                            "transcriptContentUrl": "https://graph.microsoft.com/v1.0/.../content",
                            "meetingOrganizer": {"user": {"displayName": "Alice"}},
                        }
                    ]
                }
            return {
                "id": "transcript-1",
                "meetingId": "meeting-1",
                "callId": "call-1",
                "createdDateTime": "2026-04-10T15:00:00Z",
                "endDateTime": "2026-04-10T15:30:00Z",
                "contentCorrelationId": "corr-1",
                "transcriptContentUrl": "https://graph.microsoft.com/v1.0/.../content",
                "meetingOrganizer": {"user": {"displayName": "Alice"}},
            }

        async def get_raw(self, path: str) -> bytes:
            assert path.endswith("/content?$format=text/vtt")
            return b"WEBVTT\n\n00:00:01.000 --> 00:00:02.000\n<v Alice>Hello</v>\n"

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())

    listed = await teams.teams_list_meeting_transcripts(
        teams.TeamsListMeetingTranscriptsInput(meeting_id="meeting-1")
    )
    detail = await teams.teams_get_meeting_transcript(
        teams.TeamsGetMeetingTranscriptInput(
            meeting_id="meeting-1",
            transcript_id="transcript-1",
        )
    )

    assert listed.count == 1
    assert listed.transcripts[0].meeting_organizer_display == "Alice"
    assert detail.content_type == "text/vtt"
    assert "WEBVTT" in detail.content


@pytest.mark.asyncio
async def test_teams_list_recordings_and_download_recording(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            if path.endswith("/recordings"):
                return {
                    "value": [
                        {
                            "id": "recording-1",
                            "meetingId": "meeting-1",
                            "callId": "call-1",
                            "createdDateTime": "2026-04-10T15:00:00Z",
                            "endDateTime": "2026-04-10T16:00:00Z",
                            "contentCorrelationId": "corr-1",
                            "recordingContentUrl": "https://graph.microsoft.com/v1.0/.../content",
                            "meetingOrganizer": {"user": {"displayName": "Alice"}},
                        }
                    ]
                }
            return {
                "id": "recording-1",
                "meetingId": "meeting-1",
                "callId": "call-1",
                "createdDateTime": "2026-04-10T15:00:00Z",
                "endDateTime": "2026-04-10T16:00:00Z",
                "contentCorrelationId": "corr-1",
                "recordingContentUrl": "https://graph.microsoft.com/v1.0/.../content",
                "meetingOrganizer": {"user": {"displayName": "Alice"}},
            }

        async def get_raw(self, path: str) -> bytes:
            assert path.endswith("/recordings/recording-1/content")
            return b"recording-bytes"

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())

    listed = await teams.teams_list_meeting_recordings(
        teams.TeamsListMeetingRecordingsInput(meeting_id="meeting-1")
    )
    downloaded = await teams.teams_download_meeting_recording(
        teams.TeamsDownloadMeetingRecordingInput(
            meeting_id="meeting-1",
            recording_id="recording-1",
            destination_path=tmp_path,
        )
    )

    assert listed.count == 1
    assert listed.recordings[0].meeting_organizer_display == "Alice"
    assert Path(downloaded.path).exists()
    assert Path(downloaded.path).read_bytes() == b"recording-bytes"


@pytest.mark.asyncio
async def test_teams_download_meeting_recording_reports_organizer_only_failure(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    request = httpx.Request("GET", "https://graph.microsoft.com/v1.0/me/onlineMeetings/meeting-1/recordings/recording-1/content")
    response = httpx.Response(403, request=request)

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            return {
                "id": "recording-1",
                "meetingId": "meeting-1",
                "callId": "call-1",
            }

        async def get_raw(self, path: str) -> bytes:
            raise httpx.HTTPStatusError("Forbidden", request=request, response=response)

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())

    with pytest.raises(RuntimeError, match="meeting organizer"):
        await teams.teams_download_meeting_recording(
            teams.TeamsDownloadMeetingRecordingInput(
                meeting_id="meeting-1",
                recording_id="recording-1",
                destination_path=tmp_path / "recording.mp4",
            )
        )


@pytest.mark.asyncio
async def test_teams_list_and_get_ai_insights(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            if path == "/me":
                return {"id": "user-1"}
            if path.endswith("/aiInsights"):
                return {
                    "value": [
                        {
                            "id": "insight-1",
                            "callId": "call-1",
                            "contentCorrelationId": "corr-1",
                            "createdDateTime": "2026-04-10T15:00:00Z",
                            "endDateTime": "2026-04-10T15:30:00Z",
                        }
                    ]
                }
            return {
                "id": "insight-1",
                "callId": "call-1",
                "contentCorrelationId": "corr-1",
                "createdDateTime": "2026-04-10T15:00:00Z",
                "endDateTime": "2026-04-10T15:30:00Z",
                "meetingNotes": [
                    {
                        "title": "Agenda",
                        "text": "Discuss project status",
                        "subpoints": [{"title": "Action", "text": "Prepare follow-up"}],
                    }
                ],
                "actionItems": [
                    {
                        "title": "Send update",
                        "text": "Email the team",
                        "ownerDisplayName": "Alice",
                    }
                ],
                "viewpoint": {
                    "mentionEvents": [
                        {
                            "speaker": {"user": {"displayName": "Bob"}},
                            "eventDateTime": "2026-04-10T15:10:00Z",
                            "transcriptUtterance": "Please update Sarah.",
                        }
                    ]
                },
            }

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())

    listed = await teams.teams_list_meeting_ai_insights(
        teams.TeamsListMeetingAiInsightsInput(meeting_id="meeting-1")
    )
    detail = await teams.teams_get_meeting_ai_insight(
        teams.TeamsGetMeetingAiInsightInput(
            meeting_id="meeting-1",
            ai_insight_id="insight-1",
        )
    )

    assert listed.count == 1
    assert listed.ai_insights[0].id == "insight-1"
    assert detail.meeting_notes[0].subpoints[0].title == "Action"
    assert detail.action_items[0].owner_display_name == "Alice"
    assert detail.viewpoint.mention_events[0].speaker_display == "Bob"


@pytest.mark.asyncio
async def test_teams_ai_insights_report_copilot_license_failure(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    request = httpx.Request("GET", "https://graph.microsoft.com/v1.0/copilot/users/user-1/onlineMeetings/meeting-1/aiInsights")
    response = httpx.Response(403, request=request)

    class DummyGraph:
        async def get(self, path: str, params: dict | None = None):
            if path == "/me":
                return {"id": "user-1"}
            raise httpx.HTTPStatusError("Forbidden", request=request, response=response)

    monkeypatch.setattr(teams, "get_graph", lambda _profile: DummyGraph())

    with pytest.raises(RuntimeError, match="Copilot"):
        await teams.teams_list_meeting_ai_insights(
            teams.TeamsListMeetingAiInsightsInput(meeting_id="meeting-1")
        )
