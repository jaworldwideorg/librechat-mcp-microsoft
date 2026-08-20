"""
Calendar tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.
Endpoints live under /me/calendar and /me/calendars in Graph v1.0.

Implemented:
  - list_calendars
  - list_events
  - list_upcoming_events
  - get_event
  - create_event
  - update_event
  - delete_event
  - rsvp_event
  - get_free_busy
  - find_meeting_times
"""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any, Literal, Optional, Union
from pydantic import model_validator

from mcp_microsoft.common.formatting import format_datetime_display
from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.text import strip_html
from mcp_microsoft.common.tooling import DESTRUCTIVE_TOOL, READ_ONLY_TOOL, WRITE_TOOL, register_tool
from mcp_microsoft.feature_flags import is_deletion_disabled
from mcp_microsoft.graph_types import (
    GraphAttendee,
    GraphCalendar,
    GraphEvent,
    GraphItemBody,
    parse_graph_collection,
)
from mcp_microsoft.models import (
    AttendeeAvailabilityInfo,
    AttendeeInfo,
    CreateEventResponse,
    DeleteEventResponse,
    DisplayAddress,
    EventRecurrenceInfo,
    EventDetailResponse,
    EventSummary,
    FreeBusyResponse,
    ListCalendarsResponse,
    ListEventsResponse,
    ListUpcomingEventsResponse,
    MeetingSuggestion,
    MeetingSuggestionsResponse,
    PersonSchedule,
    RsvpEventResponse,
    ScheduleItemInfo,
    CalendarInfo,
    UpdateEventResponse,
)
from mcp_microsoft.graph import get_graph

# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

BodyType = Literal["text", "html"]
RsvpResponse = Literal["accept", "decline", "tentativelyAccept"]


def _normalize_filter_datetime(value: str) -> str:
    """Validate and normalize a caller-supplied datetime for an OData filter."""
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as exc:
        raise ValueError(
            "filter_start must be a valid ISO 8601 datetime like "
            "2026-04-01T00:00:00Z."
        ) from exc
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    else:
        parsed = parsed.astimezone(timezone.utc)
    return parsed.strftime("%Y-%m-%dT%H:%M:%SZ")


class ListCalendarsInput(ToolRequestModel):
    profile: str | None = None


class ListEventsInput(ToolRequestModel):
    max_results: int = 10
    filter_start: str | None = None
    calendar_id: str | None = None
    profile: str | None = None


class ListUpcomingEventsInput(ToolRequestModel):
    start_datetime: str
    end_datetime: str
    max_results: int = 25
    calendar_id: str | None = None
    profile: str | None = None


class GetEventInput(ToolRequestModel):
    event_id: str
    profile: str | None = None


class CreateEventInput(ToolRequestModel):
    subject: str
    start_datetime: str
    end_datetime: str
    timezone: str = "UTC"
    body: str | None = None
    body_type: BodyType = "text"
    location: str | None = None
    attendees: str | list[str] | None = None
    optional_attendees: str | list[str] | None = None
    is_all_day: bool = False
    is_online_meeting: bool = False
    calendar_id: str | None = None
    profile: str | None = None


class UpdateEventInput(ToolRequestModel):
    event_id: str
    subject: str | None = None
    start_datetime: str | None = None
    end_datetime: str | None = None
    timezone: str = "UTC"
    body: str | None = None
    body_type: BodyType = "text"
    location: str | None = None
    attendees: str | list[str] | None = None
    is_all_day: bool | None = None
    profile: str | None = None


class DeleteEventInput(ToolRequestModel):
    event_id: str
    profile: str | None = None


class RsvpEventInput(ToolRequestModel):
    event_id: str
    response: RsvpResponse
    comment: str | None = None
    send_response: bool = True
    profile: str | None = None

    @model_validator(mode="after")
    def _validate_comment_requires_send_response(self):
        if self.comment and not self.send_response:
            raise ValueError("comment requires send_response=true for Graph RSVP actions.")
        return self


class GetFreeBusyInput(ToolRequestModel):
    email_addresses: str | list[str]
    start_datetime: str
    end_datetime: str
    timezone: str = "UTC"
    profile: str | None = None


class FindMeetingTimesInput(ToolRequestModel):
    attendees: str | list[str]
    duration_minutes: int = 60
    start_datetime: str | None = None
    end_datetime: str | None = None
    timezone: str = "UTC"
    max_candidates: int = 5
    is_organizer_optional: bool = False
    profile: str | None = None


def _fmt_attendees(attendees: list[GraphAttendee]) -> str:
    """Format Graph attendee objects as a readable string."""
    parts = []
    for att in attendees or []:
        ea = att.email_address
        name = ea.name
        addr = ea.address
        status = att.status.response
        label = f"{name} <{addr}>" if name else addr
        if status and status != "none":
            label += f" [{status}]"
        parts.append(label)
    return ", ".join(parts) if parts else "(none)"


def _build_datetime(dt_str: str, tz: str) -> dict:
    """Build a Graph dateTimeTimeZone object."""
    return {"dateTime": dt_str, "timeZone": tz}


def _parse_attendees(
    emails: Optional[Union[str, list[str]]],
    attendee_type: str = "required",
) -> list[dict]:
    """Convert email addresses to Graph attendee format."""
    if not emails:
        return []
    if isinstance(emails, str):
        addresses = [addr.strip() for addr in emails.split(",") if addr.strip()]
    else:
        addresses = [addr.strip() for addr in emails if addr.strip()]
    return [
        {
            "emailAddress": {"address": addr},
            "type": attendee_type,
        }
        for addr in addresses
    ]


def _attendee_values(attendees: list[GraphAttendee]) -> list[AttendeeInfo]:
    """Normalize Graph attendee objects into simple dictionaries."""
    values: list[AttendeeInfo] = []
    for attendee in attendees or []:
        values.append(
            AttendeeInfo(
                name=attendee.email_address.name,
                address=attendee.email_address.address,
                type=attendee.type,
                response=attendee.status.response,
            )
        )
    return values


def _event_body_text(body: GraphItemBody) -> str:
    raw_body = body.content
    if body.content_type.lower() == "html" and raw_body:
        return strip_html(raw_body)
    return raw_body


def _event_summary(event: GraphEvent) -> EventSummary:
    """Normalize a Graph event into a summary payload."""
    return EventSummary(
        id=event.id,
        subject=event.subject or "(no subject)",
        start=event.start.date_time,
        start_display=format_datetime_display(event.start.date_time),
        end=event.end.date_time,
        end_display=format_datetime_display(event.end.date_time),
        timezone=event.start.time_zone,
        location=event.location.display_name,
        is_all_day=event.is_all_day,
        is_cancelled=event.is_cancelled,
        response_status=event.response_status.response,
    )


# ---------------------------------------------------------------------------
# list_calendars
# ---------------------------------------------------------------------------


async def list_calendars(params: ListCalendarsInput) -> ListCalendarsResponse:
    """
    List all calendars in the user's mailbox.

    Args:
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured calendar metadata.
    """
    g = get_graph(params.profile)
    query = {
        "$select": "id,name,color,isDefaultCalendar,canEdit",
        "$top": 50,
    }

    result = await g.get("/me/calendars", params=query)
    calendars = parse_graph_collection(result, GraphCalendar)

    return ListCalendarsResponse(
        count=len(calendars),
        calendars=[
            CalendarInfo(
                id=cal.id,
                name=cal.name or "(unnamed)",
                color=cal.color or "auto",
                is_default=cal.is_default_calendar,
                can_edit=cal.can_edit,
            )
            for cal in calendars
        ],
    )


# ---------------------------------------------------------------------------
# list_events
# ---------------------------------------------------------------------------


async def list_events(
    params: ListEventsInput,
) -> ListEventsResponse:
    """
    List events from a calendar.

    For recurring events use list_upcoming_events instead, which correctly
    expands recurring event instances via the calendarView endpoint.

    Args:
        max_results: Maximum number of events to return (1-100). Defaults to 10.
        filter_start: Optional ISO 8601 datetime to filter events starting after
                      this time. E.g. '2026-03-31T00:00:00'.
        calendar_id: Optional calendar ID. Defaults to the primary calendar.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured event summaries.
    """
    g = get_graph(params.profile)
    query: dict[str, Any] = {
        "$top": params.max_results,
        "$select": "id,subject,start,end,location,organizer,isAllDay,isCancelled,responseStatus",
        "$orderby": "start/dateTime",
    }
    if params.filter_start:
        safe_start = _normalize_filter_datetime(params.filter_start)
        query["$filter"] = f"start/dateTime ge '{safe_start}'"

    if params.calendar_id:
        path = f"/me/calendars/{params.calendar_id}/events"
    else:
        path = "/me/calendar/events"

    result = await g.get(path, params=query)
    events = parse_graph_collection(result, GraphEvent)

    return ListEventsResponse(
        calendar_id=params.calendar_id,
        filter_start=params.filter_start,
        count=len(events),
        events=[_event_summary(ev) for ev in events],
    )


# ---------------------------------------------------------------------------
# list_upcoming_events (calendarView)
# ---------------------------------------------------------------------------


async def list_upcoming_events(
    params: ListUpcomingEventsInput,
) -> ListUpcomingEventsResponse:
    """
    List upcoming events using the calendarView endpoint, which correctly
    expands recurring event instances within the specified time window.

    Args:
        start_datetime: Start of the time window (ISO 8601, e.g. '2026-03-31T00:00:00').
        end_datetime: End of the time window (ISO 8601, e.g. '2026-04-07T23:59:59').
        max_results: Maximum number of events to return (1-100). Defaults to 25.
        calendar_id: Optional calendar ID. Defaults to the primary calendar.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured event summaries within the time window.
    """
    g = get_graph(params.profile)
    query: dict[str, Any] = {
        "startDateTime": params.start_datetime,
        "endDateTime": params.end_datetime,
        "$top": params.max_results,
        "$select": "id,subject,start,end,location,organizer,isAllDay,isCancelled,responseStatus",
        "$orderby": "start/dateTime",
    }

    if params.calendar_id:
        path = f"/me/calendars/{params.calendar_id}/calendarView"
    else:
        path = "/me/calendarView"

    result = await g.get(path, params=query)
    events = parse_graph_collection(result, GraphEvent)

    return ListUpcomingEventsResponse(
        calendar_id=params.calendar_id,
        start_datetime=params.start_datetime,
        end_datetime=params.end_datetime,
        count=len(events),
        events=[_event_summary(ev) for ev in events],
    )


# ---------------------------------------------------------------------------
# get_event
# ---------------------------------------------------------------------------


async def get_event(params: GetEventInput) -> EventDetailResponse:
    """
    Fetch a calendar event by ID with full details.

    Args:
        event_id: The Graph event ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured event details.
    """
    g = get_graph(params.profile)
    query = {
        "$select": (
            "id,subject,body,start,end,location,organizer,attendees,"
            "isAllDay,isCancelled,recurrence,onlineMeeting,webLink,"
            "responseStatus,importance,sensitivity,showAs"
        ),
    }

    ev = GraphEvent.model_validate(await g.get(f"/me/events/{params.event_id}", params=query))

    subject = ev.subject or "(no subject)"
    start = format_datetime_display(ev.start.date_time)
    start_tz = ev.start.time_zone
    end = format_datetime_display(ev.end.date_time)
    location = ev.location.display_name
    organizer_ea = ev.organizer.email_address
    organizer = organizer_ea.display
    attendees_str = _fmt_attendees(ev.attendees)
    show_as = ev.show_as
    web_link = ev.web_link
    join_url = ev.online_meeting.join_url if ev.online_meeting else ""
    recurrence = EventRecurrenceInfo.model_validate(
        ev.recurrence.model_dump(by_alias=True)
    ) if ev.recurrence else None
    recurrence_str = ""
    if recurrence:
        pattern = ev.recurrence.pattern if ev.recurrence else None
        recurrence_str = f"{pattern.type} (interval: {pattern.interval})" if pattern else ""

    return EventDetailResponse(
        id=params.event_id,
        subject=subject,
        start=ev.start.date_time,
        start_display=start,
        end=ev.end.date_time,
        end_display=end,
        timezone=start_tz,
        location=location,
        organizer=DisplayAddress(
            display=organizer,
            name=organizer_ea.name,
            address=organizer_ea.address,
        ),
        attendees=_attendee_values(ev.attendees),
        attendees_display=attendees_str,
        is_all_day=ev.is_all_day,
        is_cancelled=ev.is_cancelled,
        show_as=show_as,
        web_link=web_link,
        join_url=join_url,
        body=_event_body_text(ev.body),
        body_content_type=ev.body.content_type.lower(),
        recurrence=recurrence,
        recurrence_display=recurrence_str,
        importance=ev.importance,
        sensitivity=ev.sensitivity,
    )


# ---------------------------------------------------------------------------
# create_event
# ---------------------------------------------------------------------------


async def create_event(
    params: CreateEventInput,
) -> CreateEventResponse:
    """
    Create a new calendar event.

    Args:
        subject: Event title.
        start_datetime: Start time in ISO 8601 format (e.g. '2026-04-01T10:00:00').
        end_datetime: End time in ISO 8601 format (e.g. '2026-04-01T11:00:00').
        timezone: IANA timezone string (e.g. 'America/Sao_Paulo', 'UTC'). Defaults to 'UTC'.
        body: Optional event description/body.
        body_type: 'text' or 'html'. Defaults to 'text'.
        location: Optional location display name.
        attendees: Optional required attendee email(s). Comma-separated string or list.
        optional_attendees: Optional attendee email(s). Comma-separated string or list.
        is_all_day: When True, create an all-day event. Defaults to False.
        is_online_meeting: When True, create as Teams online meeting. Defaults to False.
        calendar_id: Optional calendar ID. Defaults to the primary calendar.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured event creation confirmation.
    """
    g = get_graph(params.profile)
    event: dict = {
        "subject": params.subject,
        "start": _build_datetime(params.start_datetime, params.timezone),
        "end": _build_datetime(params.end_datetime, params.timezone),
        "isAllDay": params.is_all_day,
    }

    if params.body:
        event["body"] = {
            "contentType": "HTML" if params.body_type.lower() == "html" else "Text",
            "content": params.body,
        }

    if params.location:
        event["location"] = {"displayName": params.location}

    all_attendees = []
    all_attendees.extend(_parse_attendees(params.attendees, "required"))
    all_attendees.extend(_parse_attendees(params.optional_attendees, "optional"))
    if all_attendees:
        event["attendees"] = all_attendees

    if params.is_online_meeting:
        event["isOnlineMeeting"] = True
        event["onlineMeetingProvider"] = "teamsForBusiness"

    if params.calendar_id:
        path = f"/me/calendars/{params.calendar_id}/events"
    else:
        path = "/me/calendar/events"

    created = GraphEvent.model_validate(await g.post(path, json=event) or {})
    event_id = created.id or "unknown"
    web_link = created.web_link
    join_url = created.online_meeting.join_url if created.online_meeting else ""

    return CreateEventResponse(
        success=True,
        action="create_event",
        event_id=event_id,
        subject=params.subject,
        start_datetime=params.start_datetime,
        end_datetime=params.end_datetime,
        timezone=params.timezone,
        calendar_id=params.calendar_id,
        web_link=web_link,
        join_url=join_url,
    )


# ---------------------------------------------------------------------------
# update_event
# ---------------------------------------------------------------------------


async def update_event(
    params: UpdateEventInput,
) -> UpdateEventResponse:
    """
    Update an existing calendar event. Only provided fields are changed.

    Args:
        event_id: The Graph event ID to update.
        subject: Replace event title (omit to leave unchanged).
        start_datetime: Replace start time (ISO 8601, omit to leave unchanged).
        end_datetime: Replace end time (ISO 8601, omit to leave unchanged).
        timezone: IANA timezone for start/end times. Defaults to 'UTC'.
        body: Replace event body (omit to leave unchanged).
        body_type: 'text' or 'html'. Defaults to 'text'.
        location: Replace location (omit to leave unchanged).
        attendees: Replace attendee list (omit to leave unchanged).
        is_all_day: Set all-day flag (omit to leave unchanged).
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured event update confirmation.
    """
    g = get_graph(params.profile)
    patch: dict = {}

    if params.subject is not None:
        patch["subject"] = params.subject
    if params.start_datetime is not None:
        patch["start"] = _build_datetime(params.start_datetime, params.timezone)
    if params.end_datetime is not None:
        patch["end"] = _build_datetime(params.end_datetime, params.timezone)
    if params.body is not None:
        patch["body"] = {
            "contentType": "HTML" if params.body_type.lower() == "html" else "Text",
            "content": params.body,
        }
    if params.location is not None:
        patch["location"] = {"displayName": params.location}
    if params.attendees is not None:
        patch["attendees"] = _parse_attendees(params.attendees, "required")
    if params.is_all_day is not None:
        patch["isAllDay"] = params.is_all_day

    if not patch:
        return UpdateEventResponse(
            success=False,
            action="update_event",
            event_id=params.event_id,
            updated_fields=[],
            error="No fields to update.",
        )

    result = await g.patch(f"/me/events/{params.event_id}", json=patch)

    updated_id = (result or {}).get("id", params.event_id)
    updated_fields = ", ".join(patch.keys())
    return UpdateEventResponse(
        success=True,
        action="update_event",
        event_id=updated_id,
        updated_fields=list(patch.keys()),
        updated_fields_display=updated_fields,
    )


# ---------------------------------------------------------------------------
# delete_event
# ---------------------------------------------------------------------------


async def delete_event(params: DeleteEventInput) -> DeleteEventResponse:
    """
    Delete a calendar event.

    Args:
        event_id: The Graph event ID to delete.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured event deletion confirmation.
    """
    g = get_graph(params.profile)
    await g.delete(f"/me/events/{params.event_id}")
    return DeleteEventResponse(success=True, action="delete_event", event_id=params.event_id)


# ---------------------------------------------------------------------------
# rsvp_event
# ---------------------------------------------------------------------------


async def rsvp_event(
    params: RsvpEventInput,
) -> RsvpEventResponse:
    """
    Respond to a calendar event invitation (accept, decline, or tentatively accept).

    Args:
        event_id: The Graph event ID to respond to.
        response: One of 'accept', 'decline', or 'tentativelyAccept'.
        comment: Optional comment to include with the response. Requires send_response=True.
        send_response: When True (default), send the response to the organizer.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured RSVP confirmation.
    """
    g = get_graph(params.profile)
    payload: dict = {"sendResponse": params.send_response}
    if params.comment:
        payload["comment"] = params.comment

    await g.post(f"/me/events/{params.event_id}/{params.response}", json=payload)
    return RsvpEventResponse(
        success=True,
        action="rsvp_event",
        event_id=params.event_id,
        response=params.response,
        comment=params.comment,
        send_response=params.send_response,
    )


# ---------------------------------------------------------------------------
# get_free_busy
# ---------------------------------------------------------------------------


async def get_free_busy(
    params: GetFreeBusyInput,
) -> FreeBusyResponse:
    """
    Check free/busy availability for one or more people.

    Args:
        email_addresses: Email address(es) to check. Comma-separated string or list.
        start_datetime: Start of the time window (ISO 8601).
        end_datetime: End of the time window (ISO 8601).
        timezone: IANA timezone. Defaults to 'UTC'.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured free/busy data for each requested person.
    """
    g = get_graph(params.profile)
    if isinstance(params.email_addresses, str):
        schedules = [addr.strip() for addr in params.email_addresses.split(",") if addr.strip()]
    else:
        schedules = [addr.strip() for addr in params.email_addresses if addr.strip()]

    payload = {
        "schedules": schedules,
        "startTime": _build_datetime(params.start_datetime, params.timezone),
        "endTime": _build_datetime(params.end_datetime, params.timezone),
        "availabilityViewInterval": 30,  # 30-minute slots
    }

    result = await g.post("/me/calendar/getSchedule", json=payload)
    schedule_items = (result or {}).get("value", [])

    people: list[PersonSchedule] = []
    for sched in schedule_items:
        email = sched.get("scheduleId", "unknown")
        availability = sched.get("availabilityView", "")
        items = sched.get("scheduleItems", [])
        people.append(
            PersonSchedule(
                email=email,
                availability_view=availability,
                legend={
                    "0": "Free",
                    "1": "Tentative",
                    "2": "Busy",
                    "3": "OOF",
                    "4": "Working Elsewhere",
                },
                schedule_items=[
                    ScheduleItemInfo(
                        status=item.get("status", ""),
                        start=item.get("start", {}).get("dateTime"),
                        start_display=format_datetime_display(item.get("start", {}).get("dateTime")),
                        end=item.get("end", {}).get("dateTime"),
                        end_display=format_datetime_display(item.get("end", {}).get("dateTime")),
                        subject=item.get("subject") or "",
                    )
                    for item in items
                ],
            )
        )

    return FreeBusyResponse(
        start_datetime=params.start_datetime,
        end_datetime=params.end_datetime,
        timezone=params.timezone,
        people=people,
    )


# ---------------------------------------------------------------------------
# find_meeting_times
# ---------------------------------------------------------------------------


async def find_meeting_times(
    params: FindMeetingTimesInput,
) -> MeetingSuggestionsResponse:
    """
    Find available meeting time suggestions for a set of attendees.

    Uses the Microsoft Graph findMeetingTimes endpoint to suggest times
    when all (or most) attendees are available.

    Args:
        attendees: Attendee email address(es). Comma-separated string or list.
        duration_minutes: Meeting duration in minutes. Defaults to 60.
        start_datetime: Optional start of the search window (ISO 8601).
        end_datetime: Optional end of the search window (ISO 8601).
        timezone: IANA timezone. Defaults to 'UTC'.
        max_candidates: Maximum number of time suggestions to return. Defaults to 5.
        is_organizer_optional: When True, organizer's schedule is not considered. Defaults to False.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured meeting-time suggestions.
    """
    g = get_graph(params.profile)
    attendee_list = _parse_attendees(params.attendees, "required")

    payload: dict = {
        "attendees": attendee_list,
        "meetingDuration": f"PT{params.duration_minutes}M",
        "maxCandidates": params.max_candidates,
        "isOrganizerOptional": params.is_organizer_optional,
    }

    if params.start_datetime and params.end_datetime:
        payload["timeConstraint"] = {
            "activityDomain": "work",
            "timeSlots": [
                {
                    "start": _build_datetime(params.start_datetime, params.timezone),
                    "end": _build_datetime(params.end_datetime, params.timezone),
                }
            ],
        }

    result = await g.post("/me/findMeetingTimes", json=payload)

    suggestions = (result or {}).get("meetingTimeSuggestions", [])
    emptiness = (result or {}).get("emptySuggestionsReason", "")

    normalized: list[MeetingSuggestion] = []
    for sug in suggestions:
        confidence = sug.get("confidence", 0)
        meeting_ts = sug.get("meetingTimeSlot", {})
        normalized.append(
            MeetingSuggestion(
                confidence=confidence,
                start=meeting_ts.get("start", {}).get("dateTime"),
                start_display=format_datetime_display(meeting_ts.get("start", {}).get("dateTime")),
                end=meeting_ts.get("end", {}).get("dateTime"),
                end_display=format_datetime_display(meeting_ts.get("end", {}).get("dateTime")),
                attendee_availability=[
                    AttendeeAvailabilityInfo(
                        email=aa.get("attendee", {}).get("emailAddress", {}).get("address", ""),
                        availability=aa.get("availability", ""),
                    )
                    for aa in sug.get("attendeeAvailability", [])
                ],
            )
        )

    return MeetingSuggestionsResponse(
        count=len(normalized),
        suggestions=normalized,
        empty_suggestions_reason=emptiness or None,
        timezone=params.timezone,
    )


def register(server) -> None:
    """Register all calendar tools with the given FastMCP server instance."""
    register_tool(server, list_calendars, annotations=READ_ONLY_TOOL)
    register_tool(server, list_events, annotations=READ_ONLY_TOOL)
    register_tool(server, list_upcoming_events, annotations=READ_ONLY_TOOL)
    register_tool(server, get_event, annotations=READ_ONLY_TOOL)
    register_tool(server, create_event, annotations=WRITE_TOOL)
    register_tool(server, update_event, annotations=WRITE_TOOL)
    if not is_deletion_disabled():
        register_tool(server, delete_event, annotations=DESTRUCTIVE_TOOL)
    register_tool(server, rsvp_event, annotations=WRITE_TOOL)
    register_tool(server, get_free_busy, annotations=READ_ONLY_TOOL)
    register_tool(server, find_meeting_times, annotations=READ_ONLY_TOOL)
