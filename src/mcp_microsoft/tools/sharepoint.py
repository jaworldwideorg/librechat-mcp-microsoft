"""
SharePoint tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.
Endpoints live under /sites in Graph v1.0.

Note: SharePoint is only available with work/organizational Microsoft 365
accounts. Personal Outlook.com/Live accounts do not support SharePoint.

Implemented:
  - search_content            ← Microsoft Search API (tenant-wide full-text)
  - search_sharepoint_sites
  - get_sharepoint_site
  - list_site_libraries
  - list_site_files
  - get_site_file
  - upload_to_site
  - download_from_site
  - list_site_lists
  - get_list_items
  - create_list_item
  - update_list_item
  - delete_list_item
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any
from urllib.parse import quote

from fastmcp.exceptions import ToolError
from fastmcp.server.context import Context

from mcp_microsoft.common.formatting import drive_item_payload, format_datetime_display, format_size_display
from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.text import strip_html
from mcp_microsoft.common.transfer import upload_large_file_via_session
from mcp_microsoft.common.tooling import DESTRUCTIVE_TOOL, READ_ONLY_TOOL, WRITE_TOOL, register_tool
from mcp_microsoft.feature_flags import is_deletion_disabled
from mcp_microsoft.graph_types import (
    GraphDrive,
    GraphDriveItem,
    GraphSharePointList,
    GraphSharePointListItem,
    GraphSite,
    graph_identity_display,
    parse_graph_collection,
)
from mcp_microsoft.models import (
    CreateListItemResponse,
    DeleteListItemResponse,
    DownloadSiteFileResponse,
    GetListItemsResponse,
    ListSiteFilesResponse,
    ListSiteLibrariesResponse,
    ListSiteListsResponse,
    SearchContentResponse,
    SearchHit,
    SearchSharePointSitesResponse,
    SharePointFields,
    SharePointLibraryInfo,
    SharePointListInfo,
    SharePointListItemInfo,
    SharePointSiteDetailResponse,
    SharePointSiteInfo,
    SiteFileDetailResponse,
    UpdateListItemResponse,
    UploadSiteFileResponse,
)
from mcp_microsoft.config import get_app_config
from mcp_microsoft.graph import get_graph
from mcp_microsoft.profiles import _PERSONAL_TENANT_IDS, get_profile_manager

_log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

_4MB = 4 * 1024 * 1024  # Simple upload threshold


class UploadToSiteInput(ToolRequestModel):
    site_id: str
    drive_id: str
    local_path: Path | None = None
    folder_id: str | None = None
    filename: str | None = None
    content_base64: str | None = None
    uploaded_file: str | None = None
    profile: str | None = None


class CreateListItemInput(ToolRequestModel):
    site_id: str
    list_id: str
    fields: SharePointFields
    profile: str | None = None


class UpdateListItemInput(ToolRequestModel):
    site_id: str
    list_id: str
    item_id: str
    fields: SharePointFields
    profile: str | None = None


class SearchSharepointSitesInput(ToolRequestModel):
    query: str = ""
    max_results: int = 25
    profile: str | None = None


class GetSharepointSiteInput(ToolRequestModel):
    site_id: str
    profile: str | None = None


class ListSiteLibrariesInput(ToolRequestModel):
    site_id: str
    profile: str | None = None


class ListSiteFilesInput(ToolRequestModel):
    site_id: str
    drive_id: str
    folder_id: str | None = None
    max_results: int = 25
    profile: str | None = None


class GetSiteFileInput(ToolRequestModel):
    site_id: str
    drive_id: str
    item_id: str
    profile: str | None = None


class DownloadFromSiteInput(ToolRequestModel):
    site_id: str
    drive_id: str
    item_id: str
    destination_path: Path
    profile: str | None = None


class ListSiteListsInput(ToolRequestModel):
    site_id: str
    max_results: int = 25
    profile: str | None = None


class GetListItemsInput(ToolRequestModel):
    site_id: str
    list_id: str
    max_results: int = 25
    profile: str | None = None


class DeleteListItemInput(ToolRequestModel):
    site_id: str
    list_id: str
    item_id: str
    profile: str | None = None


class SearchContentInput(ToolRequestModel):
    query: str
    entity_types: list[str] | None = None
    site_id: str | None = None
    max_results: int = 25
    skip: int = 0
    profile: str | None = None


def _site_payload(site: GraphSite) -> SharePointSiteInfo:
    """Normalize a typed SharePoint site into a structured payload."""
    return SharePointSiteInfo(
        id=site.id,
        display_name=site.display_name or "(unnamed)",
        description=site.description or "",
        web_url=site.web_url,
        created_at=site.created_date_time,
        created_at_display=format_datetime_display(site.created_date_time),
        last_modified_at=site.last_modified_date_time,
        last_modified_at_display=format_datetime_display(site.last_modified_date_time),
    )



_CONSUMER_TENANT_ERROR = (
    "SharePoint tools require a work or school Microsoft 365 account. "
    "Use a profile configured for an organization tenant."
)


def _reject_consumer_tenant_from_token() -> None:
    """Reject consumer tenants in http mode using the caller's token claims.

    http mode has no profile to inspect, so the tenant is taken from the ``tid``
    claim embedded in the validated FastMCP access token. When the claims are
    unavailable or omit ``tid``, proceed — Graph itself returns 401/403 if the
    account is unsupported.
    """
    from fastmcp.server.dependencies import get_access_token

    access_token = get_access_token()
    if access_token is None:
        return
    tid = (access_token.claims.get("tid") or "").strip().lower()
    if tid and tid in _PERSONAL_TENANT_IDS:
        raise ValueError(_CONSUMER_TENANT_ERROR)


def _get_sharepoint_graph(profile: str | None):
    """Return a Graph client for SharePoint, rejecting consumer tenants.

    stdio mode reads the tenant from the resolved profile. http mode has no
    profile (ProfileManager may hold zero profiles), so it derives the tenant
    from the caller's bearer-token ``tid`` claim instead.
    """
    if get_app_config().transport == "http":
        _reject_consumer_tenant_from_token()
        return get_graph(profile)

    cfg = get_profile_manager().resolve_profile(profile)
    tenant_id = (cfg.tenant_id or "").strip().lower()
    if tenant_id in _PERSONAL_TENANT_IDS:
        raise ValueError(_CONSUMER_TENANT_ERROR)
    return get_graph(profile)


# ---------------------------------------------------------------------------
# search_sharepoint_sites
# ---------------------------------------------------------------------------


async def search_sharepoint_sites(
    params: SearchSharepointSitesInput,
) -> SearchSharePointSitesResponse:
    """
    Search SharePoint sites the user has access to.

    When query is empty, uses a wildcard search to discover accessible sites.
    Requires a work/organizational Microsoft 365 account.

    Args:
        query: Search query string. Leave empty to discover accessible sites via wildcard.
        max_results: Maximum number of sites to return (1-200). Defaults to 25.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured SharePoint site results.
    """
    g = _get_sharepoint_graph(params.profile)
    search_term = params.query if params.query else "*"
    query_params: dict[str, Any] = {
        "search": search_term,
        "$top": params.max_results,
        "$select": "id,displayName,description,webUrl",
    }

    result = await g.get("/sites", params=query_params)
    sites = parse_graph_collection(result, GraphSite)

    return SearchSharePointSitesResponse(
        query=params.query,
        count=len(sites),
        sites=[_site_payload(site) for site in sites],
        has_more=result.get("@odata.nextLink") is not None,
    )


# ---------------------------------------------------------------------------
# get_sharepoint_site
# ---------------------------------------------------------------------------


async def get_sharepoint_site(
    params: GetSharepointSiteInput,
) -> SharePointSiteDetailResponse:
    """
    Get details of a specific SharePoint site.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID (e.g. 'contoso.sharepoint.com,site-guid,web-guid').
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured site details.
    """
    g = _get_sharepoint_graph(params.profile)
    query = {
        "$select": "id,displayName,description,webUrl,createdDateTime,lastModifiedDateTime",
    }

    site = GraphSite.model_validate(await g.get(f"/sites/{params.site_id}", params=query))

    return SharePointSiteDetailResponse(
        id=site.id or params.site_id,
        display_name=site.display_name or "(unnamed)",
        description=site.description or "",
        created_at=site.created_date_time,
        created_at_display=format_datetime_display(site.created_date_time),
        last_modified_at=site.last_modified_date_time,
        last_modified_at_display=format_datetime_display(site.last_modified_date_time),
        web_url=site.web_url,
    )


# ---------------------------------------------------------------------------
# list_site_libraries
# ---------------------------------------------------------------------------


async def list_site_libraries(
    params: ListSiteLibrariesInput,
) -> ListSiteLibrariesResponse:
    """
    List document libraries (drives) in a SharePoint site.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured document library metadata.
    """
    g = _get_sharepoint_graph(params.profile)
    query = {
        "$select": "id,name,description,driveType,webUrl",
    }

    result = await g.get(f"/sites/{params.site_id}/drives", params=query)
    drives = parse_graph_collection(result, GraphDrive)

    return ListSiteLibrariesResponse(
        site_id=params.site_id,
        count=len(drives),
        libraries=[
            SharePointLibraryInfo(
                id=drive.id,
                name=drive.name or "(unnamed)",
                description=drive.description or "",
                drive_type=drive.drive_type,
                web_url=drive.web_url,
            )
            for drive in drives
        ],
    )


# ---------------------------------------------------------------------------
# list_site_files
# ---------------------------------------------------------------------------


async def list_site_files(
    params: ListSiteFilesInput,
) -> ListSiteFilesResponse:
    """
    List files and folders in a SharePoint document library.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        drive_id: The document library (drive) ID.
        folder_id: Optional folder ID to list contents of. When omitted,
                   lists the root of the document library.
        max_results: Maximum number of items to return (1-200). Defaults to 25.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured document-library item data.
    """
    g = _get_sharepoint_graph(params.profile)
    query: dict[str, Any] = {
        "$top": params.max_results,
        "$select": "id,name,size,file,folder,lastModifiedDateTime,webUrl",
        "$orderby": "name",
    }

    if params.folder_id:
        path = f"/drives/{params.drive_id}/items/{params.folder_id}/children"
    else:
        path = f"/drives/{params.drive_id}/root/children"

    result = await g.get(path, params=query)
    items = parse_graph_collection(result, GraphDriveItem)

    return ListSiteFilesResponse(
        site_id=params.site_id,
        drive_id=params.drive_id,
        folder_id=params.folder_id,
        count=len(items),
        items=[drive_item_payload(item) for item in items],
        has_more=result.get("@odata.nextLink") is not None,
    )


# ---------------------------------------------------------------------------
# get_site_file
# ---------------------------------------------------------------------------


async def get_site_file(
    params: GetSiteFileInput,
) -> SiteFileDetailResponse:
    """
    Get metadata for a file or folder in a SharePoint document library.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        drive_id: The document library (drive) ID.
        item_id: The DriveItem ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured file/folder details.
    """
    g = _get_sharepoint_graph(params.profile)
    query = {
        "$select": (
            "id,name,size,file,folder,lastModifiedDateTime,createdDateTime,"
            "webUrl,parentReference,createdBy,lastModifiedBy"
        ),
    }

    item = GraphDriveItem.model_validate(await g.get(
        f"/drives/{params.drive_id}/items/{params.item_id}",
        params=query,
    ))

    return SiteFileDetailResponse(
        site_id=params.site_id,
        drive_id=params.drive_id,
        id=item.id or params.item_id,
        name=item.name or "(unnamed)",
        type="Folder" if item.folder else "File",
        size_bytes=item.size or 0,
        size_display=format_size_display(item.size),
        created_at=item.created_date_time,
        created_at_display=format_datetime_display(item.created_date_time),
        created_by=graph_identity_display(item.created_by),
        modified_at=item.last_modified_date_time,
        modified_at_display=format_datetime_display(item.last_modified_date_time),
        modified_by=graph_identity_display(item.last_modified_by),
        path=item.parent_reference.path if item.parent_reference else "",
        child_count=item.folder.child_count if item.folder else 0,
        mime_type=item.file.mime_type if item.file else "",
        web_url=item.web_url,
    )


# ---------------------------------------------------------------------------
# upload_to_site
# ---------------------------------------------------------------------------


async def upload_to_site(
    params: UploadToSiteInput,
    ctx: Context | None = None,
) -> UploadSiteFileResponse:
    """
    Upload a local file to a SharePoint document library.

    Files under 4 MB use the simple PUT upload. Larger files use a
    resumable upload session automatically.

    Requires a work/organizational Microsoft 365 account.

    IMPORTANT: local_path must be a file on the machine running this MCP
    server (the user's computer), NOT a container or sandbox path. If your
    content only exists in memory, first write it to a file on the user's
    filesystem (e.g. their home directory or a temp folder), then pass
    that path here.

    Alternatively, pass file content as base64 via content_base64 when
    local filesystem access is not available (e.g. container environments).
    When using content_base64, filename is required.

    You may also pass uploaded_file with the name of a file previously uploaded
    via the file-upload UI (see list_files); its bytes never passed through the
    model's context window.

    Args:
        site_id: The SharePoint site ID.
        drive_id: The document library (drive) ID.
        local_path: Optional local file path on the MCP host machine.
        folder_id: Optional destination folder ID inside the document library.
        filename: Optional upload filename override. Required when using `content_base64`.
        content_base64: Optional base64-encoded file content when no local path is available.
        uploaded_file: Optional name of a file previously uploaded via the
            file-upload UI (see list_files). Mutually exclusive with local_path
            and content_base64.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured upload confirmation.
    """
    if params.local_path is not None and get_app_config().transport == "http":
        raise ToolError(
            "local_path is not available in multi-user http mode (the server's "
            "disk is not the caller's disk); use content_base64 instead."
        )

    import base64
    import tempfile

    # File-upload UI source: resolve the stored bytes by name. Mutually
    # exclusive with the local_path / content_base64 sources.
    uploaded_bytes: bytes | None = None
    default_upload_name: str | None = None
    if params.uploaded_file is not None:
        if params.local_path is not None or params.content_base64 is not None:
            raise ToolError(
                "uploaded_file cannot be combined with local_path or "
                "content_base64; provide exactly one file source."
            )
        from mcp_microsoft.uploads import resolve_uploaded_file

        uploaded_bytes, _content_type = resolve_uploaded_file(params.uploaded_file)
        default_upload_name = params.uploaded_file

    g = _get_sharepoint_graph(params.profile)
    temp_local_path: Path | None = None
    local_path = params.local_path

    try:
        # File-upload UI: stream the resolved bytes through a temp file so the
        # existing small-PUT / chunked-session paths are reused unchanged.
        if uploaded_bytes is not None:
            name_for_suffix = params.filename or default_upload_name or "upload"
            suffix = Path(name_for_suffix).suffix
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                # Record the path before writing so a write failure still cleans up.
                temp_local_path = Path(tmp.name)
                tmp.write(uploaded_bytes)
            local_path = temp_local_path

        # Base64 fallback: decode into a generated temp file instead of trusting the caller's path.
        elif (local_path is None or not local_path.is_file()) and params.content_base64:
            if not params.filename:
                return UploadSiteFileResponse(success=False, action="upload_to_site", path=str(local_path), error="filename is required when using content_base64.")
            try:
                raw = base64.b64decode(params.content_base64, validate=True)
            except Exception as e:
                return UploadSiteFileResponse(success=False, action="upload_to_site", path=str(local_path), error=f"Invalid base64: {e}")
            suffix = Path(params.filename).suffix
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                # Record the path before writing so a write failure still cleans up.
                temp_local_path = Path(tmp.name)
                tmp.write(raw)
            local_path = temp_local_path

        if local_path is None:
            return UploadSiteFileResponse(success=False, action="upload_to_site", error="Provide local_path or content_base64 with filename.")

        if not local_path.is_file():
            return UploadSiteFileResponse(success=False, action="upload_to_site", path=str(local_path), error="File not found.")

        upload_name = params.filename or default_upload_name or local_path.name
        encoded_name = quote(upload_name, safe="")
        file_size = local_path.stat().st_size

        base = f"/drives/{params.drive_id}"

        if file_size <= _4MB:
            file_bytes = local_path.read_bytes()
            if params.folder_id:
                path = f"{base}/items/{params.folder_id}:/{encoded_name}:/content"
            else:
                path = f"{base}/root:/{encoded_name}:/content"

            result = GraphDriveItem.model_validate(await g.put(path, content=file_bytes) or {})
        else:
            if params.folder_id:
                session_path = f"{base}/items/{params.folder_id}:/{encoded_name}:/createUploadSession"
            else:
                session_path = f"{base}/root:/{encoded_name}:/createUploadSession"

            session_payload = {
                "item": {
                    "@microsoft.graph.conflictBehavior": "rename",
                    "name": upload_name,
                }
            }
            session = await g.post(session_path, json=session_payload)
            upload_url = (session or {}).get("uploadUrl", "")

            if not upload_url:
                return UploadSiteFileResponse(success=False, action="upload_to_site", path=str(local_path), error="No upload URL returned.")

            result = GraphDriveItem.model_validate(
                await upload_large_file_via_session(upload_url, local_path, file_size, ctx) or {}
            )
        size_str = format_size_display(file_size)

        return UploadSiteFileResponse(
            success=True,
            action="upload_to_site",
            site_id=params.site_id,
            drive_id=params.drive_id,
            folder_id=params.folder_id,
            filename=upload_name,
            size_bytes=file_size,
            size_display=size_str,
            file_id=result.id or "unknown",
            web_url=result.web_url,
        )
    finally:
        if temp_local_path is not None:
            try:
                temp_local_path.unlink(missing_ok=True)
            except OSError:
                pass


# ---------------------------------------------------------------------------
# download_from_site
# ---------------------------------------------------------------------------


async def download_from_site(
    params: DownloadFromSiteInput,
) -> DownloadSiteFileResponse:
    """
    Download a file from a SharePoint document library to a local path.

    Requires a work/organizational Microsoft 365 account.

    IMPORTANT: destination_path must be on the machine running this MCP
    server (the user's computer). Use an absolute path on the user's
    filesystem (e.g. their home directory or Downloads folder).

    Args:
        site_id: The SharePoint site ID.
        drive_id: The document library (drive) ID.
        item_id: The DriveItem ID of the file to download.
        destination_path: Absolute path on the user's local machine. If a
                          directory is given, the original filename is used.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured download confirmation.
    """
    g = _get_sharepoint_graph(params.profile)
    base = f"/drives/{params.drive_id}/items/{params.item_id}"

    # Get item metadata for filename
    item = GraphDriveItem.model_validate(await g.get(base, params={"$select": "id,name,size"}) or {})
    filename = item.name or "download"

    # Resolve output path
    dest = params.destination_path
    if dest.is_dir():
        safe_name = Path(filename).name
        if not safe_name or safe_name.startswith("."):
            safe_name = "download"
        dest = dest / safe_name

    # Download content
    content = await g.get_raw(f"{base}/content")

    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(content)

    size_str = format_size_display(len(content))
    return DownloadSiteFileResponse(
        success=True,
        action="download_from_site",
        site_id=params.site_id,
        drive_id=params.drive_id,
        item_id=params.item_id,
        path=str(dest),
        filename=filename,
        size_bytes=len(content),
        size_display=size_str,
    )


# ---------------------------------------------------------------------------
# list_site_lists
# ---------------------------------------------------------------------------


async def list_site_lists(
    params: ListSiteListsInput,
) -> ListSiteListsResponse:
    """
    List all lists in a SharePoint site.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        max_results: Maximum number of lists to return. Defaults to 25.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured SharePoint list metadata.
    """
    g = _get_sharepoint_graph(params.profile)
    query: dict[str, Any] = {
        "$top": params.max_results,
        "$select": "id,displayName,description,webUrl,list",
    }

    result = await g.get(f"/sites/{params.site_id}/lists", params=query)
    lists = parse_graph_collection(result, GraphSharePointList)

    return ListSiteListsResponse(
        site_id=params.site_id,
        count=len(lists),
        lists=[
            SharePointListInfo(
                id=lst.id,
                display_name=lst.display_name or "(unnamed)",
                description=lst.description or "",
                web_url=lst.web_url,
                template=lst.list_.template if lst.list_ else "",
            )
            for lst in lists
        ],
    )


# ---------------------------------------------------------------------------
# get_list_items
# ---------------------------------------------------------------------------


async def get_list_items(
    params: GetListItemsInput,
) -> GetListItemsResponse:
    """
    Get items from a SharePoint list.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        list_id: The SharePoint list ID.
        max_results: Maximum number of items to return. Defaults to 25.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured SharePoint list-item data.
    """
    g = _get_sharepoint_graph(params.profile)
    query: dict[str, Any] = {
        "$top": params.max_results,
        "$expand": "fields",
        "$select": "id,createdDateTime,lastModifiedDateTime",
    }

    result = await g.get(
        f"/sites/{params.site_id}/lists/{params.list_id}/items",
        params=query,
    )
    items = parse_graph_collection(result, GraphSharePointListItem)

    normalized: list[SharePointListItemInfo] = []
    for item in items:
        item_id = item.id
        created = format_datetime_display(item.created_date_time)
        modified = format_datetime_display(item.last_modified_date_time)
        fields = item.fields

        # Filter out internal/system fields
        user_fields = {
            k: v for k, v in fields.items()
            if not k.startswith("@") and not k.startswith("_")
        }

        title = user_fields.pop("Title", user_fields.pop("title", f"Item {item_id}"))
        filtered_fields = {
            key: value
            for key, value in user_fields.items()
            if key not in ("id", "ContentType", "Attachments", "Edit", "LinkTitleNoMenu", "LinkTitle")
        }
        normalized.append(
            SharePointListItemInfo(
                id=item_id,
                title=title,
                created_at=item.created_date_time,
                created_at_display=created,
                modified_at=item.last_modified_date_time,
                modified_at_display=modified,
                fields=SharePointFields(filtered_fields),
            )
        )

    return GetListItemsResponse(
        site_id=params.site_id,
        list_id=params.list_id,
        count=len(normalized),
        items=normalized,
        has_more=result.get("@odata.nextLink") is not None,
    )


# ---------------------------------------------------------------------------
# create_list_item
# ---------------------------------------------------------------------------


async def create_list_item(
    params: CreateListItemInput,
) -> CreateListItemResponse:
    """
    Add an item to a SharePoint list.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        list_id: The SharePoint list ID.
        fields: Field values for the new item as a JSON object.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured list-item creation confirmation.
    """
    g = _get_sharepoint_graph(params.profile)

    payload = {"fields": params.fields.root}
    result = GraphSharePointListItem.model_validate(await g.post(
        f"/sites/{params.site_id}/lists/{params.list_id}/items",
        json=payload,
    ) or {})
    return CreateListItemResponse(
        success=True,
        action="create_list_item",
        site_id=params.site_id,
        list_id=params.list_id,
        item_id=result.id or "unknown",
        fields=SharePointFields(params.fields.root),
    )


# ---------------------------------------------------------------------------
# update_list_item
# ---------------------------------------------------------------------------


async def update_list_item(
    params: UpdateListItemInput,
) -> UpdateListItemResponse:
    """
    Update a SharePoint list item.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        list_id: The SharePoint list ID.
        item_id: The list item ID to update.
        fields: Field values to patch as a JSON object.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured list-item update confirmation.
    """
    g = _get_sharepoint_graph(params.profile)

    await g.patch(
        f"/sites/{params.site_id}/lists/{params.list_id}/items/{params.item_id}/fields",
        json=params.fields.root,
    )

    return UpdateListItemResponse(
        success=True,
        action="update_list_item",
        site_id=params.site_id,
        list_id=params.list_id,
        item_id=params.item_id,
        updated_fields=list(params.fields.root.keys()),
        fields=SharePointFields(params.fields.root),
    )


# ---------------------------------------------------------------------------
# delete_list_item
# ---------------------------------------------------------------------------


async def delete_list_item(
    params: DeleteListItemInput,
) -> DeleteListItemResponse:
    """
    Delete an item from a SharePoint list.

    Requires a work/organizational Microsoft 365 account.

    Args:
        site_id: The SharePoint site ID.
        list_id: The SharePoint list ID.
        item_id: The list item ID to delete.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured list-item deletion confirmation.
    """
    g = _get_sharepoint_graph(params.profile)
    await g.delete(f"/sites/{params.site_id}/lists/{params.list_id}/items/{params.item_id}")
    return DeleteListItemResponse(
        success=True,
        action="delete_list_item",
        site_id=params.site_id,
        list_id=params.list_id,
        item_id=params.item_id,
    )


# ---------------------------------------------------------------------------
# search_content
# ---------------------------------------------------------------------------

_SEARCH_ENTITY_TYPES = frozenset(
    {"driveItem", "listItem", "site", "message", "event", "chatMessage"}
)
_SEARCH_ISOLATED_ENTITY_TYPES = frozenset({"message", "event", "chatMessage"})

# Fields requested from Graph for driveItem hits.
_DRIVE_ITEM_FIELDS = [
    "id", "name", "webUrl", "size",
    "lastModifiedDateTime", "lastModifiedBy",
    "parentReference", "file",
]

# Fields requested for listItem hits (SharePoint list rows).
_LIST_ITEM_FIELDS = [
    "id", "webUrl", "lastModifiedDateTime", "lastModifiedBy",
    "parentReference", "fields",
]


def _parse_hit(hit: dict) -> SearchHit:
    resource = hit.get("resource") or {}
    odata_type = resource.get("@odata.type", "")
    # "#microsoft.graph.driveItem" -> "driveItem"
    resource_type = odata_type.split(".")[-1] if odata_type else ""

    last_modified_by_obj = resource.get("lastModifiedBy") or {}
    last_modified_by = (
        (last_modified_by_obj.get("user") or {}).get("displayName", "")
        or (last_modified_by_obj.get("application") or {}).get("displayName", "")
    )

    parent_ref = resource.get("parentReference") or {}
    file_obj = resource.get("file") or {}

    return SearchHit(
        hit_id=hit.get("hitId", ""),
        rank=hit.get("rank", 0),
        summary=strip_html(hit.get("summary", "")),
        resource_type=resource_type,
        name=resource.get("name") or resource.get("displayName") or resource.get("subject") or "",
        web_url=resource.get("webUrl", ""),
        last_modified_at=resource.get("lastModifiedDateTime") or resource.get("receivedDateTime") or "",
        last_modified_by=last_modified_by,
        size_bytes=resource.get("size") or 0,
        mime_type=(file_obj.get("mimeType") or ""),
        site_id=parent_ref.get("siteId", ""),
        drive_id=parent_ref.get("driveId", ""),
        parent_path=parent_ref.get("path", ""),
    )


async def search_content(
    params: SearchContentInput,
) -> SearchContentResponse:
    """
    Full-text search across Microsoft 365 content using the Microsoft Search API.

    Uses the same full-text index as SharePoint's built-in search — searching
    file content, metadata, and list fields across all sites the user has
    access to, not just navigation hierarchy.

    Supports KQL (Keyword Query Language) for advanced filtering:
        "budget 2024"                      — full-text across all content
        "author:John budget"               — by author
        "filetype:xlsx budget"             — Excel files about budget
        "contenttype:Document project"     — Word docs about a project
        "modified:2024-01-01..2024-03-31"  — modified in a date range
        "path:https://tenant.sharepoint.com/sites/Finance budget"
                                           — scoped to a specific site URL

    Args:
        query: KQL search query string.
        entity_types: Resource types to search. Any combination of:
            "driveItem"   — files in SharePoint libraries and OneDrive (default)
            "listItem"    — SharePoint list rows (non-file content)
            "site"        — SharePoint sites
            "message"     — emails (requires Mail.Read scope)
            "event"       — calendar events (requires Calendars.Read scope)
            Defaults to ["driveItem"].
        site_id: Optional SharePoint site ID to restrict the search to a
            specific site. When provided, the site's webUrl is fetched and a
            KQL ``path:"<webUrl>"`` clause is injected into the query string
            so that only items under that site's URL path are returned.
            When omitted the search is tenant-wide.
        max_results: Maximum results to return per page (1-500). Defaults to 25.
        skip: Number of results to skip for pagination. Pass the next_skip
            value from a previous response to fetch the next page.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Ranked search hits with name, URL, summary excerpt, author, and
        file metadata. Use next_skip for subsequent pages.
    """
    entity_types = params.entity_types or ["driveItem"]

    unknown = [et for et in entity_types if et not in _SEARCH_ENTITY_TYPES]
    if unknown:
        raise ValueError(
            f"Unknown entity type(s): {unknown}. "
            f"Valid types: {sorted(_SEARCH_ENTITY_TYPES)}"
        )

    isolated = _SEARCH_ISOLATED_ENTITY_TYPES.intersection(entity_types)
    if isolated and len(entity_types) != 1:
        raise ValueError(
            f"Graph requires {min(isolated)!r} to be searched separately; "
            "it cannot be combined with other entity types."
        )

    max_page_size = 25 if isolated else 500
    max_results = max(1, min(max_page_size, params.max_results))

    # Build the $search request object.
    search_request: dict = {
        "entityTypes": entity_types,
        "query": {"queryString": params.query},
        "from": params.skip,
        "size": max_results,
    }

    # Request only the fields we need to keep responses lean.
    if entity_types == ["driveItem"]:
        search_request["fields"] = _DRIVE_ITEM_FIELDS
    elif entity_types == ["listItem"]:
        search_request["fields"] = _LIST_ITEM_FIELDS

    g = _get_sharepoint_graph(params.profile)

    # Scope to a specific site when requested.
    # contentSources is for external connector items only — use a KQL
    # path: clause with the site's webUrl to restrict SharePoint search scope.
    if params.site_id:
        site_data = GraphSite.model_validate(
            await g.get(f"/sites/{params.site_id}", params={"$select": "webUrl"}) or {}
        )
        site_url = site_data.web_url
        if site_url:
            current_query = search_request["query"].get("queryString", "")
            search_request["query"]["queryString"] = f'{current_query} path:"{site_url}"'.strip()

    result = await g.post("/search/query", json={"requests": [search_request]})

    # Graph wraps responses in value[0].hitsContainers[0].
    containers = ((result.get("value") or [{}])[0]).get("hitsContainers") or []
    container = containers[0] if containers else {}

    raw_hits: list[dict] = container.get("hits") or []
    total: int = container.get("total") or 0
    more: bool = container.get("moreResultsAvailable") or False

    hits = [_parse_hit(h) for h in raw_hits]
    next_skip = params.skip + len(hits) if more else None

    return SearchContentResponse(
        query=params.query,
        entity_types=entity_types,
        total=total,
        count=len(hits),
        hits=hits,
        more_results_available=more,
        next_skip=next_skip,
    )


def register(server) -> None:
    """Register all SharePoint tools with the given FastMCP server instance."""
    register_tool(server, search_content, annotations=READ_ONLY_TOOL)
    register_tool(server, search_sharepoint_sites, annotations=READ_ONLY_TOOL)
    register_tool(server, get_sharepoint_site, annotations=READ_ONLY_TOOL)
    register_tool(server, list_site_libraries, annotations=READ_ONLY_TOOL)
    register_tool(server, list_site_files, annotations=READ_ONLY_TOOL)
    register_tool(server, get_site_file, annotations=READ_ONLY_TOOL)
    register_tool(server, upload_to_site, annotations=WRITE_TOOL)
    if get_app_config().transport == "http":
        _log.info(
            "download_from_site not registered (http transport; server disk "
            "is not the caller's disk)"
        )
    else:
        register_tool(server, download_from_site, annotations=WRITE_TOOL)
    register_tool(server, list_site_lists, annotations=READ_ONLY_TOOL)
    register_tool(server, get_list_items, annotations=READ_ONLY_TOOL)
    register_tool(server, create_list_item, annotations=WRITE_TOOL)
    register_tool(server, update_list_item, annotations=WRITE_TOOL)
    if not is_deletion_disabled():
        register_tool(server, delete_list_item, annotations=DESTRUCTIVE_TOOL)
