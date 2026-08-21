from __future__ import annotations

from typing import Any

from pydantic import BaseModel, ConfigDict, Field, RootModel


class MCPModel(BaseModel):
    model_config = ConfigDict(populate_by_name=True)


class FlexibleMCPModel(MCPModel):
    model_config = ConfigDict(populate_by_name=True, extra="allow")


class Address(MCPModel):
    name: str = ""
    address: str = ""


class DisplayAddress(Address):
    display: str = ""


class AttachmentInfo(MCPModel):
    id: str = ""
    name: str = ""
    size_bytes: int = 0
    size_display: str = ""
    content_type: str = ""
    is_inline: bool = False
    size_kb: int = 0


class MessageSummary(MCPModel):
    id: str = ""
    subject: str = ""
    from_: DisplayAddress = Field(default_factory=DisplayAddress, alias="from")
    received_at: str | None = None
    received_at_display: str = ""
    is_read: bool = False
    has_attachments: bool = False
    importance: str = ""
    preview: str = ""


class ListEmailsResponse(MCPModel):
    folder: str = ""
    count: int = 0
    messages: list[MessageSummary] = Field(default_factory=list)
    next_page_token: str | None = None
    has_more: bool = False


class ReadEmailResponse(MCPModel):
    id: str = ""
    subject: str = ""
    from_: DisplayAddress = Field(default_factory=DisplayAddress, alias="from")
    to: list[Address] = Field(default_factory=list)
    to_display: str = ""
    cc: list[Address] = Field(default_factory=list)
    cc_display: str = ""
    received_at: str | None = None
    received_at_display: str = ""
    is_read: bool = False
    conversation_id: str = ""
    importance: str = ""
    body: str = ""
    body_content_type: str = ""
    attachments: list[AttachmentInfo] = Field(default_factory=list)


class ReadEmailSummaryResponse(MCPModel):
    id: str = ""
    subject: str = ""
    from_: DisplayAddress = Field(default_factory=DisplayAddress, alias="from")
    received_at: str | None = None
    received_at_display: str = ""
    is_read: bool = False
    preview: str = ""


class SearchEmailsResponse(MCPModel):
    query: str = ""
    folder: str | None = None
    count: int = 0
    messages: list[MessageSummary] = Field(default_factory=list)


class ActionResult(MCPModel):
    success: bool
    action: str
    error: str | None = None


class SendEmailResponse(ActionResult):
    to: list[str] = Field(default_factory=list)
    cc: list[str] = Field(default_factory=list)
    bcc: list[str] = Field(default_factory=list)
    reply_to: list[str] = Field(default_factory=list)
    subject: str = ""
    body_type: str = ""
    saved_to_sent_items: bool = False


class ReplyEmailResponse(ActionResult):
    message_id: str = ""
    body_type: str = ""


class ForwardEmailResponse(ActionResult):
    message_id: str = ""
    to: list[str] = Field(default_factory=list)
    comment: str = ""


class MarkEmailReadResponse(ActionResult):
    message_id: str = ""
    is_read: bool = False


class MoveEmailResponse(ActionResult):
    message_id: str = ""
    new_message_id: str = ""
    destination_folder: str = ""


class TrashEmailResponse(MoveEmailResponse):
    soft_delete: bool = True
    profile: str | None = None


class DeleteEmailResponse(ActionResult):
    message_id: str = ""
    irreversible: bool = True


class BulkEmailFailure(MCPModel):
    message_id: str = ""
    status: int = 0
    code: str | None = None
    error: str = ""


class BulkMovedEmail(MCPModel):
    source_message_id: str = ""
    new_message_id: str = ""


class BulkMoveEmailsResponse(ActionResult):
    destination_folder: str = ""
    total: int = 0
    succeeded: int = 0
    failed: int = 0
    moved: list[BulkMovedEmail] = Field(default_factory=list)
    failures: list[BulkEmailFailure] = Field(default_factory=list)


class BulkTrashEmailsResponse(ActionResult):
    total: int = 0
    succeeded: int = 0
    failed: int = 0
    moved: list[BulkMovedEmail] = Field(default_factory=list)
    failures: list[BulkEmailFailure] = Field(default_factory=list)


class BulkDeleteEmailsResponse(ActionResult):
    total: int = 0
    succeeded: int = 0
    failed: int = 0
    irreversible: bool = True
    failures: list[BulkEmailFailure] = Field(default_factory=list)


class DraftSummary(MCPModel):
    id: str = ""
    subject: str = ""
    to: list[Address] = Field(default_factory=list)
    last_modified_at: str | None = None
    last_modified_at_display: str = ""
    preview: str = ""


class CreateDraftResponse(ActionResult):
    draft_id: str = ""
    to: list[str] = Field(default_factory=list)
    cc: list[str] = Field(default_factory=list)
    subject: str = ""
    body_type: str = ""


class ListDraftsResponse(MCPModel):
    count: int = 0
    drafts: list[DraftSummary] = Field(default_factory=list)


class DraftDetailResponse(MCPModel):
    id: str = ""
    subject: str = ""
    to: list[Address] = Field(default_factory=list)
    cc: list[Address] = Field(default_factory=list)
    bcc: list[Address] = Field(default_factory=list)
    last_modified_at: str | None = None
    last_modified_at_display: str = ""
    body: str = ""
    body_content_type: str = ""
    is_draft: bool = True


class UpdateDraftResponse(ActionResult):
    draft_id: str = ""
    updated_fields: list[str] = Field(default_factory=list)
    updated_fields_display: str = ""


class SendDraftResponse(ActionResult):
    draft_id: str = ""


class MailFolderInfo(MCPModel):
    id: str = ""
    display_name: str = ""
    display_label: str = ""
    unread_count: int = 0
    total_count: int = 0
    child_folder_count: int = 0
    is_child: bool = False


class ListFoldersResponse(MCPModel):
    count: int = 0
    include_child_folders: bool = False
    folders: list[MailFolderInfo] = Field(default_factory=list)


class CreateFolderResponse(ActionResult):
    folder_id: str = ""
    display_name: str = ""
    parent_folder_id: str | None = None


class DeleteFolderResponse(ActionResult):
    folder_id: str = ""
    irreversible: bool = True


class ListAttachmentsResponse(MCPModel):
    message_id: str = ""
    count: int = 0
    attachments: list[AttachmentInfo] = Field(default_factory=list)


class DownloadAttachmentResponse(ActionResult):
    message_id: str = ""
    attachment_id: str = ""
    filename: str = ""
    path: str | None = None
    size_bytes: int = 0
    size_display: str = ""
    content_type: str = ""


class ReadAttachmentResponse(MCPModel):
    message_id: str = ""
    attachment_id: str = ""
    filename: str = ""
    content_type: str = ""
    text: str = ""
    page_count: int = 0
    truncated: bool = False


class CalendarInfo(MCPModel):
    id: str = ""
    name: str = ""
    color: str = ""
    is_default: bool = False
    can_edit: bool = False


class ListCalendarsResponse(MCPModel):
    count: int = 0
    calendars: list[CalendarInfo] = Field(default_factory=list)


class AttendeeInfo(MCPModel):
    name: str = ""
    address: str = ""
    type: str = ""
    response: str = ""


class EventSummary(MCPModel):
    id: str = ""
    subject: str = ""
    start: str | None = None
    start_display: str = ""
    end: str | None = None
    end_display: str = ""
    timezone: str = ""
    location: str = ""
    is_all_day: bool = False
    is_cancelled: bool = False
    response_status: str = ""


class ListEventsResponse(MCPModel):
    calendar_id: str | None = None
    filter_start: str | None = None
    count: int = 0
    events: list[EventSummary] = Field(default_factory=list)


class ListUpcomingEventsResponse(MCPModel):
    calendar_id: str | None = None
    start_datetime: str = ""
    end_datetime: str = ""
    count: int = 0
    events: list[EventSummary] = Field(default_factory=list)


class EventDetailResponse(MCPModel):
    id: str = ""
    subject: str = ""
    start: str | None = None
    start_display: str = ""
    end: str | None = None
    end_display: str = ""
    timezone: str = ""
    location: str = ""
    organizer: DisplayAddress = Field(default_factory=DisplayAddress)
    attendees: list[AttendeeInfo] = Field(default_factory=list)
    attendees_display: str = ""
    is_all_day: bool = False
    is_cancelled: bool = False
    show_as: str = ""
    web_link: str = ""
    join_url: str = ""
    body: str = ""
    body_content_type: str = ""
    recurrence: "EventRecurrenceInfo | None" = None
    recurrence_display: str = ""
    importance: str = ""
    sensitivity: str = ""


class EventRecurrencePatternInfo(MCPModel):
    type: str = ""
    interval: int = 1


class EventRecurrenceRangeInfo(FlexibleMCPModel):
    type: str = ""
    start_date: str = ""
    end_date: str = ""
    recurrence_time_zone: str = ""
    number_of_occurrences: int | None = None


class EventRecurrenceInfo(MCPModel):
    pattern: EventRecurrencePatternInfo | None = None
    range_: EventRecurrenceRangeInfo | None = Field(default=None, alias="range")


class CreateEventResponse(ActionResult):
    event_id: str = ""
    subject: str = ""
    start_datetime: str = ""
    end_datetime: str = ""
    timezone: str = ""
    calendar_id: str | None = None
    web_link: str = ""
    join_url: str = ""


class UpdateEventResponse(ActionResult):
    event_id: str = ""
    updated_fields: list[str] = Field(default_factory=list)
    updated_fields_display: str = ""


class DeleteEventResponse(ActionResult):
    event_id: str = ""


class RsvpEventResponse(ActionResult):
    event_id: str = ""
    response: str = ""
    comment: str | None = None
    send_response: bool = True


class ScheduleItemInfo(MCPModel):
    status: str = ""
    start: str | None = None
    start_display: str = ""
    end: str | None = None
    end_display: str = ""
    subject: str = ""


class PersonSchedule(MCPModel):
    email: str = ""
    availability_view: str = ""
    legend: dict[str, str] = Field(default_factory=dict)
    schedule_items: list[ScheduleItemInfo] = Field(default_factory=list)


class FreeBusyResponse(MCPModel):
    start_datetime: str = ""
    end_datetime: str = ""
    timezone: str = ""
    people: list[PersonSchedule] = Field(default_factory=list)


class AttendeeAvailabilityInfo(MCPModel):
    email: str = ""
    availability: str = ""


class MeetingSuggestion(MCPModel):
    confidence: float | int = 0
    start: str | None = None
    start_display: str = ""
    end: str | None = None
    end_display: str = ""
    attendee_availability: list[AttendeeAvailabilityInfo] = Field(default_factory=list)


class MeetingSuggestionsResponse(MCPModel):
    count: int = 0
    suggestions: list[MeetingSuggestion] = Field(default_factory=list)
    empty_suggestions_reason: str | None = None
    timezone: str = ""


class DriveItemInfo(MCPModel):
    id: str = ""
    name: str = ""
    size_bytes: int = 0
    size_display: str = ""
    last_modified_at: str | None = None
    last_modified_at_display: str = ""
    web_url: str = ""
    is_folder: bool = False
    child_count: int = 0
    mime_type: str = ""
    parent_path: str = ""


class ListDriveItemsResponse(MCPModel):
    folder_id: str | None = None
    count: int = 0
    items: list[DriveItemInfo] = Field(default_factory=list)
    has_more: bool = False


class DriveItemDetailResponse(MCPModel):
    id: str = ""
    name: str = ""
    type: str = ""
    size_bytes: int = 0
    size_display: str = ""
    created_at: str | None = None
    created_at_display: str = ""
    created_by: str = ""
    modified_at: str | None = None
    modified_at_display: str = ""
    modified_by: str = ""
    path: str = ""
    web_url: str = ""
    child_count: int = 0
    mime_type: str = ""


class SearchDriveResponse(MCPModel):
    query: str = ""
    count: int = 0
    items: list[DriveItemInfo] = Field(default_factory=list)


class CreateDriveFolderResponse(ActionResult):
    folder_id: str = ""
    name: str = ""
    parent_folder_id: str | None = None
    web_url: str = ""


class UploadFileResponse(ActionResult):
    filename: str = ""
    size_bytes: int = 0
    size_display: str = ""
    file_id: str = ""
    web_url: str = ""
    parent_folder_id: str | None = None
    path: str | None = None


class DownloadFileResponse(ActionResult):
    item_id: str = ""
    path: str = ""
    filename: str = ""
    size_bytes: int = 0
    size_display: str = ""
    expected_size_bytes: int = 0


class DeleteDriveItemResponse(ActionResult):
    item_id: str = ""
    soft_delete: bool = True


class MoveOrCopyItemResponse(ActionResult):
    item_id: str = ""
    new_item_id: str = ""
    destination_folder_id: str = ""
    name: str | None = None
    status: str | None = None


class SharePointSiteInfo(MCPModel):
    id: str = ""
    display_name: str = ""
    description: str = ""
    web_url: str = ""
    created_at: str | None = None
    created_at_display: str = ""
    last_modified_at: str | None = None
    last_modified_at_display: str = ""


class SearchSharePointSitesResponse(MCPModel):
    query: str = ""
    count: int = 0
    sites: list[SharePointSiteInfo] = Field(default_factory=list)
    has_more: bool = False


class SharePointSiteDetailResponse(SharePointSiteInfo):
    pass


class SharePointLibraryInfo(MCPModel):
    id: str = ""
    name: str = ""
    description: str = ""
    drive_type: str = ""
    web_url: str = ""


class ListSiteLibrariesResponse(MCPModel):
    site_id: str = ""
    count: int = 0
    libraries: list[SharePointLibraryInfo] = Field(default_factory=list)


class ListSiteFilesResponse(MCPModel):
    site_id: str = ""
    drive_id: str = ""
    folder_id: str | None = None
    count: int = 0
    items: list[DriveItemInfo] = Field(default_factory=list)
    has_more: bool = False


class SiteFileDetailResponse(MCPModel):
    site_id: str = ""
    drive_id: str = ""
    id: str = ""
    name: str = ""
    type: str = ""
    size_bytes: int = 0
    size_display: str = ""
    created_at: str | None = None
    created_at_display: str = ""
    created_by: str = ""
    modified_at: str | None = None
    modified_at_display: str = ""
    modified_by: str = ""
    path: str = ""
    child_count: int = 0
    mime_type: str = ""
    web_url: str = ""


class UploadSiteFileResponse(ActionResult):
    site_id: str = ""
    drive_id: str = ""
    folder_id: str | None = None
    filename: str = ""
    size_bytes: int = 0
    size_display: str = ""
    file_id: str = ""
    web_url: str = ""
    path: str | None = None


class DownloadSiteFileResponse(ActionResult):
    site_id: str = ""
    drive_id: str = ""
    item_id: str = ""
    path: str = ""
    filename: str = ""
    size_bytes: int = 0
    size_display: str = ""


class SharePointListInfo(MCPModel):
    id: str = ""
    display_name: str = ""
    description: str = ""
    web_url: str = ""
    template: str = ""


class ListSiteListsResponse(MCPModel):
    site_id: str = ""
    count: int = 0
    lists: list[SharePointListInfo] = Field(default_factory=list)


class SharePointListItemInfo(MCPModel):
    id: str = ""
    title: str = ""
    created_at: str | None = None
    created_at_display: str = ""
    modified_at: str | None = None
    modified_at_display: str = ""
    fields: "SharePointFields" = Field(default_factory=lambda: SharePointFields({}))


class GetListItemsResponse(MCPModel):
    site_id: str = ""
    list_id: str = ""
    count: int = 0
    items: list[SharePointListItemInfo] = Field(default_factory=list)
    has_more: bool = False


class SharePointFields(RootModel[dict[str, Any]]):
    root: dict[str, Any]


class CreateListItemResponse(ActionResult):
    site_id: str = ""
    list_id: str = ""
    item_id: str = ""
    fields: SharePointFields = Field(default_factory=lambda: SharePointFields({}))


class UpdateListItemResponse(ActionResult):
    site_id: str = ""
    list_id: str = ""
    item_id: str = ""
    updated_fields: list[str] = Field(default_factory=list)
    fields: SharePointFields = Field(default_factory=lambda: SharePointFields({}))


class SearchHit(MCPModel):
    """A single result from the Microsoft Search API."""

    hit_id: str = ""
    rank: int = 0
    # Human-readable excerpt with search-term highlights (HTML tags stripped).
    summary: str = ""
    # Type of the matched resource: "driveItem", "listItem", "message", "event", "site".
    resource_type: str = ""
    # Display name / filename.
    name: str = ""
    web_url: str = ""
    # ISO-8601 last modified timestamp.
    last_modified_at: str = ""
    last_modified_by: str = ""
    # File size in bytes (0 for non-file resources).
    size_bytes: int = 0
    # MIME type for driveItem hits; empty for other types.
    mime_type: str = ""
    # SharePoint site id and path; populated for driveItem and listItem.
    site_id: str = ""
    drive_id: str = ""
    parent_path: str = ""


class SearchContentResponse(MCPModel):
    """Response from the Microsoft Search API (POST /search/query)."""

    query: str = ""
    entity_types: list[str] = Field(default_factory=list)
    total: int = 0
    count: int = 0
    hits: list[SearchHit] = Field(default_factory=list)
    more_results_available: bool = False
    # Pass as `skip` on the next call to retrieve the next page.
    next_skip: int | None = None


class DeleteListItemResponse(ActionResult):
    site_id: str = ""
    list_id: str = ""
    item_id: str = ""


class ProfileInfo(MCPModel):
    name: str = ""
    client_id_masked: str = ""
    tenant_id: str = ""
    is_default: bool = False
    is_authenticated: bool | None = None
    cache_path: str = ""


class ListProfilesResponse(MCPModel):
    default_profile: str | None = None
    count: int = 0
    profiles: list[ProfileInfo] = Field(default_factory=list)


class AddedProfileInfo(MCPModel):
    name: str = ""
    client_id: str = ""
    tenant_id: str = ""
    cache_path: str = ""
    is_default: bool = False


class AddProfileResponse(ActionResult):
    profile: AddedProfileInfo | None = None


class RemoveProfileResponse(ActionResult):
    profile: str = ""


class AuthenticateProfileResponse(ActionResult):
    profile: str | None = None
    tenant_id: str = ""
    cache_path: str = ""
    # "authenticated" | "awaiting_user" | "error"
    status: str = ""
    user_code: str | None = None
    verification_uri: str | None = None
    expires_in_seconds: int | None = None
    device_code_message: str | None = None
    instructions: str | None = None


class SetDefaultProfileResponse(ActionResult):
    profile: str = ""


# ---------------------------------------------------------------------------
# Contacts models
# ---------------------------------------------------------------------------


class ContactEmailAddress(MCPModel):
    address: str = ""
    name: str = ""


class ContactInfo(MCPModel):
    id: str = ""
    display_name: str = ""
    given_name: str = ""
    surname: str = ""
    email_addresses: list[ContactEmailAddress] = Field(default_factory=list)
    mobile_phone: str = ""
    business_phones: list[str] = Field(default_factory=list)
    job_title: str = ""
    company_name: str = ""
    department: str = ""


class ListContactsResponse(MCPModel):
    count: int = 0
    folder_id: str | None = None
    contacts: list[ContactInfo] = Field(default_factory=list)
    next_page_token: str | None = None
    has_more: bool = False
    pages_scanned: int = 0


class GetContactResponse(MCPModel):
    id: str = ""
    display_name: str = ""
    given_name: str = ""
    surname: str = ""
    email_addresses: list[ContactEmailAddress] = Field(default_factory=list)
    mobile_phone: str = ""
    business_phones: list[str] = Field(default_factory=list)
    job_title: str = ""
    company_name: str = ""
    department: str = ""
    notes: str = ""


class CreateContactResponse(ActionResult):
    contact_id: str = ""
    display_name: str = ""
    folder_id: str | None = None


class UpdateContactResponse(ActionResult):
    contact_id: str = ""
    updated_fields: list[str] = Field(default_factory=list)
    updated_fields_display: str = ""


class DeleteContactResponse(ActionResult):
    contact_id: str = ""


class ContactFolderInfo(MCPModel):
    id: str = ""
    display_name: str = ""
    parent_folder_id: str = ""
    total_item_count: int = 0


class ListContactFoldersResponse(MCPModel):
    count: int = 0
    folders: list[ContactFolderInfo] = Field(default_factory=list)


class SearchContactsResponse(MCPModel):
    query: str = ""
    count: int = 0
    contacts: list[ContactInfo] = Field(default_factory=list)
    next_page_token: str | None = None
    has_more: bool = False
    pages_scanned: int = 0


class GetContactPhotoResponse(ActionResult):
    contact_id: str = ""
    photo_base64: str = ""
    saved_path: str = ""
    size_bytes: int = 0


class ServiceFlagsResponse(MCPModel):
    mail: bool = True
    calendar: bool = True
    contacts: bool = True
    onedrive: bool = True
    sharepoint: bool = False
    teams: bool = False
    teams_meeting_artifacts: bool = False
    teams_ai_insights: bool = False
    drafts: bool = True
    folders: bool = True
    attachments: bool = True


class TeamInfo(MCPModel):
    id: str = ""
    display_name: str = ""
    description: str = ""
    visibility: str = ""
    web_url: str = ""
    is_archived: bool | None = None


class TeamsListJoinedResponse(MCPModel):
    count: int = 0
    teams: list[TeamInfo] = Field(default_factory=list)
    next_link: str | None = None


class TeamMemberSettingsInfo(FlexibleMCPModel):
    allow_create_update_channels: bool | None = None
    allow_delete_channels: bool | None = None
    allow_add_remove_apps: bool | None = None
    allow_create_update_remove_tabs: bool | None = None
    allow_create_update_remove_connectors: bool | None = None


class TeamGuestSettingsInfo(FlexibleMCPModel):
    allow_create_update_channels: bool | None = None
    allow_delete_channels: bool | None = None


class TeamFunSettingsInfo(FlexibleMCPModel):
    allow_giphy: bool | None = None
    giphy_content_rating: str = ""
    allow_stickers_and_memes: bool | None = None
    allow_custom_memes: bool | None = None


class TeamDetailResponse(TeamInfo):
    member_settings: TeamMemberSettingsInfo = Field(default_factory=TeamMemberSettingsInfo)
    guest_settings: TeamGuestSettingsInfo = Field(default_factory=TeamGuestSettingsInfo)
    fun_settings: TeamFunSettingsInfo = Field(default_factory=TeamFunSettingsInfo)


class ChannelInfo(MCPModel):
    id: str = ""
    display_name: str = ""
    description: str = ""
    channel_type: str = ""
    web_url: str = ""
    is_favorite_by_default: bool = False


class TeamsListChannelsResponse(MCPModel):
    team_id: str = ""
    count: int = 0
    channels: list[ChannelInfo] = Field(default_factory=list)
    next_link: str | None = None


class ChannelDetailResponse(ChannelInfo):
    pass


class CreateChannelResponse(ActionResult):
    team_id: str = ""
    channel_id: str = ""
    display_name: str = ""
    description: str = ""
    web_url: str = ""
    dry_run: bool = False
    requires_confirmation: bool = False


class TeamsMessageInfo(MCPModel):
    id: str = ""
    created_at: str = ""
    created_at_display: str = ""
    last_modified_at: str = ""
    from_display: str = ""
    body: str = ""
    body_content_type: str = ""
    subject: str = ""
    web_url: str = ""
    reply_to_id: str = ""
    importance: str = ""


class TeamsListChannelMessagesResponse(MCPModel):
    team_id: str = ""
    channel_id: str = ""
    count: int = 0
    messages: list[TeamsMessageInfo] = Field(default_factory=list)
    next_link: str | None = None


class TeamsReactionInfo(FlexibleMCPModel):
    reaction_type: str = ""
    created_at: str = ""
    user_display: str = ""


class TeamsMessageAttachmentInfo(FlexibleMCPModel):
    id: str = ""
    name: str = ""
    content_type: str = ""
    content_url: str = ""
    thumbnail_url: str = ""


class TeamsMentionInfo(FlexibleMCPModel):
    id: int | str | None = None
    mention_text: str = ""
    mentioned_display: str = ""


class ChannelMessageDetailResponse(TeamsMessageInfo):
    reactions: list[TeamsReactionInfo] = Field(default_factory=list)
    attachments: list[TeamsMessageAttachmentInfo] = Field(default_factory=list)
    mentions: list[TeamsMentionInfo] = Field(default_factory=list)


class SendTeamsMessageResponse(ActionResult):
    id: str = ""
    parent_message_id: str = ""
    chat_id: str = ""
    web_url: str = ""
    created_at: str = ""
    created_at_display: str = ""


class TeamsListRepliesResponse(MCPModel):
    team_id: str = ""
    channel_id: str = ""
    parent_message_id: str = ""
    count: int = 0
    replies: list[TeamsMessageInfo] = Field(default_factory=list)
    next_link: str | None = None


class ChatInfo(MCPModel):
    id: str = ""
    chat_type: str = ""
    topic: str = ""
    created_at: str = ""
    created_at_display: str = ""
    last_updated_at: str = ""
    last_updated_at_display: str = ""
    web_url: str = ""


class TeamsListChatsResponse(MCPModel):
    count: int = 0
    chats: list[ChatInfo] = Field(default_factory=list)
    next_link: str | None = None


class ChatMemberInfo(MCPModel):
    id: str = ""
    display_name: str = ""
    email: str = ""
    roles: list[str] = Field(default_factory=list)


class ChatDetailResponse(ChatInfo):
    members: list[ChatMemberInfo] = Field(default_factory=list)


class TeamsListChatMessagesResponse(MCPModel):
    chat_id: str = ""
    count: int = 0
    messages: list[TeamsMessageInfo] = Field(default_factory=list)
    next_link: str | None = None


class CreateTeamsChatResponse(ActionResult):
    id: str = ""
    chat_type: str = ""
    topic: str = ""
    web_url: str = ""
    created_at: str = ""
    created_at_display: str = ""


class OnlineMeetingInfo(MCPModel):
    id: str = ""
    subject: str = ""
    join_web_url: str = ""
    join_meeting_id: str = ""
    start: str = ""
    end: str = ""
    start_display: str = ""
    end_display: str = ""
    created_at: str = ""


class CreateTeamsMeetingResponse(ActionResult, OnlineMeetingInfo):
    pass


class TeamsMeetingParticipantInfo(FlexibleMCPModel):
    upn: str = ""
    role: str = ""
    identity_display: str = ""


class TeamsMeetingParticipantsInfo(FlexibleMCPModel):
    organizer: TeamsMeetingParticipantInfo | None = None
    attendees: list[TeamsMeetingParticipantInfo] = Field(default_factory=list)
    producers: list[TeamsMeetingParticipantInfo] = Field(default_factory=list)
    contributors: list[TeamsMeetingParticipantInfo] = Field(default_factory=list)


class MeetingDetailResponse(OnlineMeetingInfo):
    participants: TeamsMeetingParticipantsInfo = Field(default_factory=TeamsMeetingParticipantsInfo)
    video_teleconference_id: str = ""


class TeamsListMeetingsResponse(MCPModel):
    start_after: str = ""
    start_before: str = ""
    count: int = 0
    meetings: list[OnlineMeetingInfo] = Field(default_factory=list)
    skipped_count: int = 0
    warnings: list[str] = Field(default_factory=list)
    next_link: str | None = None


class MeetingTranscriptInfo(MCPModel):
    id: str = ""
    meeting_id: str = ""
    call_id: str = ""
    created_at: str = ""
    end_at: str = ""
    content_correlation_id: str = ""
    transcript_content_url: str = ""
    meeting_organizer_display: str = ""


class TeamsListMeetingTranscriptsResponse(MCPModel):
    meeting_id: str = ""
    count: int = 0
    transcripts: list[MeetingTranscriptInfo] = Field(default_factory=list)
    next_link: str | None = None


class MeetingTranscriptDetailResponse(MeetingTranscriptInfo):
    content_type: str = ""
    content: str = ""


class MeetingRecordingInfo(MCPModel):
    id: str = ""
    meeting_id: str = ""
    call_id: str = ""
    created_at: str = ""
    end_at: str = ""
    content_correlation_id: str = ""
    recording_content_url: str = ""
    meeting_organizer_display: str = ""


class TeamsListMeetingRecordingsResponse(MCPModel):
    meeting_id: str = ""
    count: int = 0
    recordings: list[MeetingRecordingInfo] = Field(default_factory=list)
    next_link: str | None = None


class DownloadMeetingRecordingResponse(ActionResult):
    meeting_id: str = ""
    recording_id: str = ""
    path: str = ""
    filename: str = ""
    size_bytes: int = 0
    size_display: str = ""


class MeetingAiInsightInfo(MCPModel):
    id: str = ""
    meeting_id: str = ""
    call_id: str = ""
    created_at: str = ""
    end_at: str = ""
    content_correlation_id: str = ""


class MeetingAiInsightNoteSubpoint(MCPModel):
    title: str = ""
    text: str = ""


class MeetingAiInsightNote(MCPModel):
    title: str = ""
    text: str = ""
    subpoints: list[MeetingAiInsightNoteSubpoint] = Field(default_factory=list)


class MeetingAiInsightActionItem(MCPModel):
    title: str = ""
    text: str = ""
    owner_display_name: str = ""


class MeetingAiInsightMentionEvent(MCPModel):
    speaker_display: str = ""
    event_at: str = ""
    transcript_utterance: str = ""


class MeetingAiInsightViewpoint(MCPModel):
    mention_events: list[MeetingAiInsightMentionEvent] = Field(default_factory=list)


class TeamsListMeetingAiInsightsResponse(MCPModel):
    meeting_id: str = ""
    count: int = 0
    ai_insights: list[MeetingAiInsightInfo] = Field(default_factory=list)
    next_link: str | None = None


class MeetingAiInsightDetailResponse(MeetingAiInsightInfo):
    meeting_notes: list[MeetingAiInsightNote] = Field(default_factory=list)
    action_items: list[MeetingAiInsightActionItem] = Field(default_factory=list)
    viewpoint: MeetingAiInsightViewpoint = Field(default_factory=MeetingAiInsightViewpoint)
