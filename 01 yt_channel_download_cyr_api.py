"""
YouTube Data API Research Client
Version: 1.0

Purpose
-------
This application is a custom Python client developed exclusively for an academic
research project investigating public audit and corruption exposure in Ukraine
through YouTube content.

The application retrieves only publicly available information through the
official YouTube Data API v3 and does not modify, upload, delete, or interact
with YouTube content or user accounts.

Research Workflow
-----------------
1. Read a list of YouTube channels from a text file.
2. Resolve each channel to its official Channel ID.
3. Retrieve channel metadata.
4. Retrieve the uploads playlist associated with each channel.
5. Download metadata for every public video.
6. Classify videos into regular videos, livestreams and possible Shorts.
7. Retrieve public playlists and podcast playlists.
8. Retrieve playlist membership information.
9. Export all collected data to JSON and Excel for subsequent statistical,
   computational and qualitative analysis.

YouTube Data API Methods Used
-----------------------------
• channels.list
• playlistItems.list
• playlists.list
• videos.list

Data Collected
--------------
The application collects only publicly available metadata, including:

• channel title and description
• channel statistics
• video metadata
• publication dates
• video duration
• view, like and comment counts
• playlist metadata
• thumbnail URLs
• topic categories

No private user information is requested or stored.

Output
------
The application exports structured datasets in JSON and Microsoft Excel (.xlsx)
format. These datasets are used for computational social science research,
natural language processing, statistical analysis and reproducible academic
research.

Compliance
----------
Review the YouTube Data API Terms of Service before collecting data.
This client accesses only publicly available resources through the official API,
implements request pagination and rate limiting, and is intended solely for
non-commercial academic research.

Author
------
Igor Kudinov
PhD Research Project
2026
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any
from urllib.parse import quote, unquote, urljoin, urlparse

import pandas as pd
import requests


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_API_KEY_FILE = str(BASE_DIR / "keys" / "youtube_api_key.txt")
DEFAULT_CHANNELS_FILE = str(BASE_DIR / "channels.txt")
DEFAULT_OUTPUT_DIR = str(BASE_DIR / "outputs" / "youtube")

YOUTUBE_API_BASE = "https://www.googleapis.com/youtube/v3"
YOUTUBE_BASE_URL = "https://www.youtube.com"
REQUEST_TIMEOUT = 30
MAX_PAGE_SIZE = 50
MAX_VIDEO_IDS_PER_REQUEST = 50
EXCEL_MAX_CELL_LENGTH = 32767

CHANNEL_PARTS = "snippet,contentDetails,statistics,brandingSettings,status"
PLAYLIST_PARTS = "snippet,contentDetails,status"
PLAYLIST_ITEM_PARTS = "snippet,contentDetails,status"
VIDEO_PARTS = "snippet,contentDetails,statistics,status,liveStreamingDetails,topicDetails,recordingDetails"

SHORTS_THREE_MINUTE_CUTOFF = date(2024, 10, 15)
DEFAULT_SHORTS_MAX_SECONDS_AFTER_CUTOFF = 180
DEFAULT_SHORTS_MAX_SECONDS_BEFORE_CUTOFF = 60

KNOWN_TABS = {
    "videos",
    "streams",
    "shorts",
    "featured",
    "playlists",
    "podcasts",
    "community",
    "channels",
    "about",
    "live",
    "search",
}

ILLEGAL_EXCEL_CHARS_RE = re.compile(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]")
ISO_DURATION_RE = re.compile(
    r"^P"
    r"(?:(?P<days>\d+(?:[.,]\d+)?)D)?"
    r"(?:T"
    r"(?:(?P<hours>\d+(?:[.,]\d+)?)H)?"
    r"(?:(?P<minutes>\d+(?:[.,]\d+)?)M)?"
    r"(?:(?P<seconds>\d+(?:[.,]\d+)?)S)?"
    r")?$"
)


@dataclass
class ChannelContext:
    source: str
    channel_id: str
    channel_title: str
    channel_stem: str
    channel_url: str
    uploads_playlist_id: str
    custom_url: str


class YouTubeAPIError(RuntimeError):
    """Raised when YouTube Data API returns an error response."""


class YouTubeDataAPI:
    def __init__(self, api_key: str, delay: float = 0.0) -> None:
        self.api_key = api_key
        self.delay = delay
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Accept": "application/json",
                "User-Agent": (
                    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                ),
            }
        )

    def get(self, resource: str, params: dict[str, Any]) -> dict[str, Any]:
        request_params = dict(params)
        request_params["key"] = self.api_key

        if self.delay:
            time.sleep(self.delay)

        try:
            response = self.session.get(
                f"{YOUTUBE_API_BASE}/{resource.lstrip('/')}",
                params=request_params,
                timeout=REQUEST_TIMEOUT,
            )
        except requests.RequestException as exc:
            # Requests errors can contain the URL query, including the API key.
            raise YouTubeAPIError(f"{resource}: network request failed ({type(exc).__name__})") from None

        try:
            payload = response.json()
        except ValueError:
            raise YouTubeAPIError(f"{resource}: server returned invalid JSON (HTTP {response.status_code})") from None
        if not isinstance(payload, dict):
            raise YouTubeAPIError(f"{resource}: server returned an unexpected response format")

        if response.status_code >= 400:
            error = payload.get("error", {}) if isinstance(payload, dict) else {}
            message = str(error.get("message") or response.text.strip() or response.reason).replace(self.api_key, "[redacted]")
            raise YouTubeAPIError(f"{resource}: HTTP {response.status_code}: {message}")

        return payload

    def list_all(self, resource: str, params: dict[str, Any]) -> list[dict[str, Any]]:
        items: list[dict[str, Any]] = []
        page_token: str | None = None

        while True:
            page_params = dict(params)
            page_params["maxResults"] = MAX_PAGE_SIZE
            if page_token:
                page_params["pageToken"] = page_token

            payload = self.get(resource, page_params)
            items.extend(payload.get("items", []))

            page_token = payload.get("nextPageToken")
            if not page_token:
                return items


def validate_file(path_str: str, label: str) -> Path:
    path = Path(path_str)

    if not path.exists():
        raise FileNotFoundError(f"{label} not found: {path}")

    if path.is_file() and path.stat().st_size == 0:
        raise ValueError(f"{label} is empty: {path}")

    return path


def read_api_key(path_str: str | None) -> str | None:
    if not path_str:
        return None

    path = Path(path_str)
    if not path.exists():
        return None

    value = path.read_text(encoding="utf-8").strip()
    return value or None


def load_channels(file_path: str) -> list[str]:
    with open(file_path, "r", encoding="utf-8") as file:
        return [
            line.strip()
            for line in file
            if line.strip() and not line.lstrip().startswith("#")
        ]


def encode_youtube_path(path: str) -> str:
    decoded_path = unquote(path)
    return quote(decoded_path, safe="/@:")


def normalize_channel_source(channel_value: str) -> str:
    channel_value = channel_value.strip().rstrip("/")

    if not channel_value:
        return channel_value

    if channel_value.startswith(("http://", "https://")):
        parsed = urlparse(channel_value)
        decoded_path = unquote(parsed.path)
        path_parts = [part for part in decoded_path.split("/") if part]

        if path_parts and path_parts[-1] in KNOWN_TABS:
            path_parts = path_parts[:-1]

        clean_path = "/" + "/".join(path_parts)
        encoded_path = encode_youtube_path(clean_path)
        return f"{parsed.scheme}://{parsed.netloc}{encoded_path}"

    decoded_value = unquote(channel_value)

    if decoded_value.startswith("@"):
        encoded_handle = quote(decoded_value, safe="@")
        return f"{YOUTUBE_BASE_URL}/{encoded_handle}"

    if decoded_value.startswith(("channel/", "user/", "c/")):
        encoded_path = quote(decoded_value, safe="/@:")
        return f"{YOUTUBE_BASE_URL}/{encoded_path}"

    if re.fullmatch(r"UC[\w-]{20,}", decoded_value):
        return f"{YOUTUBE_BASE_URL}/channel/{decoded_value}"

    encoded_handle = quote("@" + decoded_value, safe="@")
    return f"{YOUTUBE_BASE_URL}/{encoded_handle}"


def safe_filename(name: Any) -> str:
    name = unquote(str(name))
    bad_chars = '<>:"/\\|?*'

    for ch in bad_chars:
        name = name.replace(ch, "_")

    name = name.replace("https___", "")
    name = name.replace("http___", "")
    name = name.replace("www.youtube.com_", "")
    name = name.replace("youtube.com_", "")

    while "__" in name:
        name = name.replace("__", "_")

    return name.strip("_ ").strip() or "youtube_channel"


def get_channel_file_stem(channel_value: str, fallback_title: str = "") -> str:
    channel_value = channel_value.strip().rstrip("/")

    if channel_value.startswith(("http://", "https://")):
        parsed = urlparse(channel_value)
        decoded_path = unquote(parsed.path)
        path_parts = [part for part in decoded_path.split("/") if part]

        if path_parts and path_parts[-1] in KNOWN_TABS:
            path_parts = path_parts[:-1]

        if path_parts:
            return safe_filename("_".join(path_parts))

        return safe_filename(parsed.netloc)

    if channel_value:
        return safe_filename(channel_value.replace("@", ""))

    return safe_filename(fallback_title)


def make_channel_url(channel_id: str) -> str:
    return f"{YOUTUBE_BASE_URL}/channel/{channel_id}"


def make_video_url(video_id: str) -> str:
    return f"{YOUTUBE_BASE_URL}/watch?v={video_id}"


def make_playlist_url(playlist_id: str) -> str:
    return f"{YOUTUBE_BASE_URL}/playlist?list={playlist_id}"


def best_thumbnail(thumbnails: dict[str, Any] | None) -> str:
    if not thumbnails:
        return ""

    for key in ("maxres", "standard", "high", "medium", "default"):
        item = thumbnails.get(key)
        if isinstance(item, dict) and item.get("url"):
            return str(item["url"])

    return ""


def parse_youtube_datetime(value: str | None) -> datetime | None:
    if not value:
        return None

    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None


def parse_youtube_date(value: str | None) -> date | None:
    parsed = parse_youtube_datetime(value)
    return parsed.date() if parsed else None


def parse_iso8601_duration_seconds(value: str | None) -> int | None:
    if not value:
        return None

    match = ISO_DURATION_RE.match(value)
    if not match:
        return None

    def number(group_name: str) -> float:
        group_value = match.group(group_name)
        if not group_value:
            return 0.0
        return float(group_value.replace(",", "."))

    seconds = (
        number("days") * 86400
        + number("hours") * 3600
        + number("minutes") * 60
        + number("seconds")
    )

    return int(seconds)


def format_duration(seconds: int | None) -> str:
    if seconds is None:
        return ""

    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    secs = seconds % 60

    if hours:
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"

    return f"{minutes:02d}:{secs:02d}"


def to_int(value: Any) -> int | None:
    if value in (None, ""):
        return None

    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def is_stream_video(video_item: dict[str, Any]) -> bool:
    snippet = video_item.get("snippet", {})
    live_details = video_item.get("liveStreamingDetails", {})
    live_broadcast_content = snippet.get("liveBroadcastContent")

    return bool(live_details) or live_broadcast_content in {"live", "upcoming"}


def is_possible_short_video(
    video_item: dict[str, Any],
    seconds_before_cutoff: int,
    seconds_after_cutoff: int,
) -> bool:
    content_details = video_item.get("contentDetails", {})
    snippet = video_item.get("snippet", {})
    duration_seconds = parse_iso8601_duration_seconds(content_details.get("duration"))

    if duration_seconds is None:
        return False

    published_date = parse_youtube_date(snippet.get("publishedAt"))

    if published_date and published_date >= SHORTS_THREE_MINUTE_CUTOFF:
        return duration_seconds <= seconds_after_cutoff

    return duration_seconds <= seconds_before_cutoff


def shorts_classification_note(row: dict[str, Any]) -> str:
    if not row.get("is_possible_short"):
        return ""

    published_at = row.get("published_at")
    published_date = parse_youtube_date(published_at)

    if published_date and published_date >= SHORTS_THREE_MINUTE_CUTOFF:
        return "possible_short_duration_only_after_2024_10_15"

    return "possible_short_duration_only_before_2024_10_15"


def sanitize_excel_value(value: Any) -> Any:
    if value is None:
        return ""

    if isinstance(value, (int, float, bool)):
        return value

    if isinstance(value, (list, tuple, set)):
        text = " | ".join(str(item) for item in value)
    elif isinstance(value, dict):
        text = json.dumps(value, ensure_ascii=False)
    else:
        text = str(value)

    text = ILLEGAL_EXCEL_CHARS_RE.sub("", text)
    text = text.replace("\r", " ").replace("\n", " ").strip()
    text = re.sub(r"\s+", " ", text)

    if len(text) > EXCEL_MAX_CELL_LENGTH:
        text = text[: EXCEL_MAX_CELL_LENGTH - 3] + "..."

    return text


def rows_for_excel(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        {key: sanitize_excel_value(value) for key, value in row.items()}
        for row in rows
    ]


def save_rows(rows: list[dict[str, Any]], base_path: Path, sheet_name: str) -> Path | None:
    if not rows:
        return None

    base_path.parent.mkdir(parents=True, exist_ok=True)

    json_path = base_path.parent / (base_path.name + ".json")
    xlsx_path = base_path.parent / (base_path.name + ".xlsx")

    with open(json_path, "w", encoding="utf-8") as file:
        json.dump(rows, file, ensure_ascii=False, indent=2)

    df = pd.DataFrame(rows_for_excel(rows))

    with pd.ExcelWriter(
        xlsx_path,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False, "strings_to_formulas": False}},
    ) as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])

    return xlsx_path


def extract_channel_id_from_page(channel_url: str) -> tuple[str | None, str]:
    response = requests.get(
        channel_url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            )
        },
        timeout=REQUEST_TIMEOUT,
    )
    response.raise_for_status()
    html = response.text

    channel_id_match = (
        re.search(r'"externalId":"(UC[\w-]+)"', html)
        or re.search(r'"channelId":"(UC[\w-]+)"', html)
        or re.search(r"https://www\.youtube\.com/channel/(UC[\w-]+)", html)
    )
    canonical_match = re.search(r'"canonicalBaseUrl":"([^"]+)"', html)

    channel_id = channel_id_match.group(1) if channel_id_match else None
    canonical_url = channel_url

    if canonical_match:
        canonical_url = urljoin(
            YOUTUBE_BASE_URL,
            canonical_match.group(1).replace("\\u0026", "&").replace("\\/", "/"),
        )
    elif channel_id:
        canonical_url = make_channel_url(channel_id)

    return channel_id, canonical_url


def resolve_channel(api: YouTubeDataAPI, source: str) -> tuple[ChannelContext, dict[str, Any]]:
    normalized = normalize_channel_source(source)
    parsed = urlparse(normalized)
    parts = [part for part in unquote(parsed.path).strip("/").split("/") if part]

    channel_item: dict[str, Any] | None = None
    canonical_url = normalized

    if len(parts) >= 2 and parts[0] == "channel" and parts[1].startswith("UC"):
        channel_item = fetch_channel_by_id(api, parts[1])
        canonical_url = make_channel_url(parts[1])

    elif parts and parts[0].startswith("@"):
        channel_item = fetch_channel_by_handle(api, parts[0])
        canonical_url = normalized

    elif len(parts) >= 2 and parts[0] == "user":
        channel_item = fetch_channel_by_username(api, parts[1])

    if channel_item is None:
        channel_id, canonical_url = extract_channel_id_from_page(normalized)
        if not channel_id:
            raise RuntimeError(f"Could not resolve channel id from {source}")
        channel_item = fetch_channel_by_id(api, channel_id)

    snippet = channel_item.get("snippet", {})
    content_details = channel_item.get("contentDetails", {})
    related_playlists = content_details.get("relatedPlaylists", {})
    channel_id = channel_item.get("id", "")
    uploads_playlist_id = related_playlists.get("uploads", "")

    if not channel_id:
        raise RuntimeError(f"YouTube API returned channel without id for {source}")

    if not uploads_playlist_id:
        raise RuntimeError(f"YouTube API returned channel without uploads playlist for {source}")

    channel_title = snippet.get("title") or channel_id
    custom_url = snippet.get("customUrl") or ""
    channel_stem = get_channel_file_stem(source, fallback_title=channel_title)

    context = ChannelContext(
        source=source,
        channel_id=channel_id,
        channel_title=channel_title,
        channel_stem=channel_stem,
        channel_url=canonical_url or make_channel_url(channel_id),
        uploads_playlist_id=uploads_playlist_id,
        custom_url=custom_url,
    )

    return context, channel_item


def fetch_channel_by_id(api: YouTubeDataAPI, channel_id: str) -> dict[str, Any]:
    payload = api.get(
        "channels",
        {
            "part": CHANNEL_PARTS,
            "id": channel_id,
        },
    )
    items = payload.get("items", [])

    if not items:
        raise RuntimeError(f"Channel not found by id: {channel_id}")

    return items[0]


def fetch_channel_by_handle(api: YouTubeDataAPI, handle: str) -> dict[str, Any] | None:
    payload = api.get(
        "channels",
        {
            "part": CHANNEL_PARTS,
            "forHandle": handle,
        },
    )
    items = payload.get("items", [])
    return items[0] if items else None


def fetch_channel_by_username(api: YouTubeDataAPI, username: str) -> dict[str, Any] | None:
    payload = api.get(
        "channels",
        {
            "part": CHANNEL_PARTS,
            "forUsername": username,
        },
    )
    items = payload.get("items", [])
    return items[0] if items else None


def build_channel_metadata_row(context: ChannelContext, channel_item: dict[str, Any]) -> dict[str, Any]:
    snippet = channel_item.get("snippet", {})
    statistics = channel_item.get("statistics", {})
    branding = channel_item.get("brandingSettings", {}).get("channel", {})
    status = channel_item.get("status", {})

    return {
        "source": context.source,
        "channel_id": context.channel_id,
        "channel_title": context.channel_title,
        "channel_url": context.channel_url,
        "custom_url": context.custom_url,
        "uploads_playlist_id": context.uploads_playlist_id,
        "uploads_playlist_url": make_playlist_url(context.uploads_playlist_id),
        "published_at": snippet.get("publishedAt", ""),
        "description": snippet.get("description", ""),
        "country": branding.get("country") or snippet.get("country", ""),
        "view_count": to_int(statistics.get("viewCount")),
        "subscriber_count": to_int(statistics.get("subscriberCount")),
        "hidden_subscriber_count": statistics.get("hiddenSubscriberCount", ""),
        "video_count": to_int(statistics.get("videoCount")),
        "privacy_status": status.get("privacyStatus", ""),
        "made_for_kids": status.get("madeForKids", ""),
        "thumbnail_url": best_thumbnail(snippet.get("thumbnails")),
    }


def fetch_upload_playlist_items(api: YouTubeDataAPI, uploads_playlist_id: str) -> list[dict[str, Any]]:
    return api.list_all(
        "playlistItems",
        {
            "part": PLAYLIST_ITEM_PARTS,
            "playlistId": uploads_playlist_id,
        },
    )


def fetch_video_details(api: YouTubeDataAPI, video_ids: list[str]) -> dict[str, dict[str, Any]]:
    details: dict[str, dict[str, Any]] = {}

    for idx in range(0, len(video_ids), MAX_VIDEO_IDS_PER_REQUEST):
        chunk = video_ids[idx : idx + MAX_VIDEO_IDS_PER_REQUEST]
        payload = api.get(
            "videos",
            {
                "part": VIDEO_PARTS,
                "id": ",".join(chunk),
            },
        )

        for item in payload.get("items", []):
            video_id = item.get("id")
            if video_id:
                details[video_id] = item

    return details


def build_upload_rows(
    context: ChannelContext,
    playlist_items: list[dict[str, Any]],
    video_details: dict[str, dict[str, Any]],
    seconds_before_cutoff: int,
    seconds_after_cutoff: int,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    for fallback_position, playlist_item in enumerate(playlist_items):
        item_snippet = playlist_item.get("snippet", {})
        item_content = playlist_item.get("contentDetails", {})
        item_status = playlist_item.get("status", {})
        resource_id = item_snippet.get("resourceId", {})
        video_id = item_content.get("videoId") or resource_id.get("videoId") or ""

        video_item = video_details.get(video_id, {})
        snippet = video_item.get("snippet", {})
        content_details = video_item.get("contentDetails", {})
        statistics = video_item.get("statistics", {})
        status = video_item.get("status", {})
        live_details = video_item.get("liveStreamingDetails", {})
        recording_details = video_item.get("recordingDetails", {})
        topic_details = video_item.get("topicDetails", {})

        duration_iso = content_details.get("duration", "")
        duration_seconds = parse_iso8601_duration_seconds(duration_iso)
        is_stream = is_stream_video(video_item)
        is_possible_short = is_possible_short_video(
            video_item,
            seconds_before_cutoff=seconds_before_cutoff,
            seconds_after_cutoff=seconds_after_cutoff,
        )

        row = {
            "channel": context.source,
            "channel_stem": context.channel_stem,
            "channel_id": context.channel_id,
            "channel_title": context.channel_title,
            "channel_url": context.channel_url,
            "source_section": "uploads",
            "api_section": classify_video_section(is_stream, is_possible_short),
            "playlist_item_id": playlist_item.get("id", ""),
            "uploads_playlist_id": context.uploads_playlist_id,
            "uploads_position": item_snippet.get("position", fallback_position),
            "video_id": video_id,
            "video_title": snippet.get("title") or item_snippet.get("title", ""),
            "video_url": make_video_url(video_id) if video_id else "",
            "description": snippet.get("description", ""),
            "published_at": snippet.get("publishedAt") or item_content.get("videoPublishedAt", ""),
            "playlist_item_published_at": item_snippet.get("publishedAt", ""),
            "video_owner_channel_id": item_snippet.get("videoOwnerChannelId", ""),
            "video_owner_channel_title": item_snippet.get("videoOwnerChannelTitle", ""),
            "duration_iso": duration_iso,
            "duration_seconds": duration_seconds,
            "duration_hms": format_duration(duration_seconds),
            "dimension": content_details.get("dimension", ""),
            "definition": content_details.get("definition", ""),
            "caption": content_details.get("caption", ""),
            "licensed_content": content_details.get("licensedContent", ""),
            "projection": content_details.get("projection", ""),
            "live_broadcast_content": snippet.get("liveBroadcastContent", ""),
            "scheduled_start_time": live_details.get("scheduledStartTime", ""),
            "actual_start_time": live_details.get("actualStartTime", ""),
            "actual_end_time": live_details.get("actualEndTime", ""),
            "concurrent_viewers": to_int(live_details.get("concurrentViewers")),
            "recording_date": recording_details.get("recordingDate", ""),
            "view_count": to_int(statistics.get("viewCount")),
            "like_count": to_int(statistics.get("likeCount")),
            "comment_count": to_int(statistics.get("commentCount")),
            "favorite_count": to_int(statistics.get("favoriteCount")),
            "privacy_status": status.get("privacyStatus") or item_status.get("privacyStatus", ""),
            "license": status.get("license", ""),
            "embeddable": status.get("embeddable", ""),
            "public_stats_viewable": status.get("publicStatsViewable", ""),
            "made_for_kids": status.get("madeForKids", ""),
            "category_id": snippet.get("categoryId", ""),
            "default_language": snippet.get("defaultLanguage", ""),
            "default_audio_language": snippet.get("defaultAudioLanguage", ""),
            "tags": " | ".join(snippet.get("tags", []) or []),
            "topic_categories": " | ".join(topic_details.get("topicCategories", []) or []),
            "thumbnail_url": best_thumbnail(snippet.get("thumbnails") or item_snippet.get("thumbnails")),
            "is_stream": is_stream,
            "is_possible_short": is_possible_short,
        }

        row["shorts_classification_note"] = shorts_classification_note(row)
        rows.append(row)

    return rows


def classify_video_section(is_stream: bool, is_possible_short: bool) -> str:
    if is_stream:
        return "streams"

    if is_possible_short:
        return "shorts"

    return "videos"


def fetch_channel_playlists(api: YouTubeDataAPI, channel_id: str) -> list[dict[str, Any]]:
    return api.list_all(
        "playlists",
        {
            "part": PLAYLIST_PARTS,
            "channelId": channel_id,
        },
    )


def playlist_source_section(playlist_item: dict[str, Any]) -> str:
    podcast_status = playlist_item.get("status", {}).get("podcastStatus", "")
    return "podcasts" if podcast_status == "enabled" else "playlists"


def build_playlist_rows(context: ChannelContext, playlist_items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    for playlist_item in playlist_items:
        snippet = playlist_item.get("snippet", {})
        status = playlist_item.get("status", {})
        content_details = playlist_item.get("contentDetails", {})
        playlist_id = playlist_item.get("id", "")
        source_section = playlist_source_section(playlist_item)

        rows.append(
            {
                "channel": context.source,
                "channel_stem": context.channel_stem,
                "channel_id": context.channel_id,
                "channel_title": context.channel_title,
                "source_section": source_section,
                "playlist_id": playlist_id,
                "playlist_title": snippet.get("title", ""),
                "playlist_url": make_playlist_url(playlist_id) if playlist_id else "",
                "description": snippet.get("description", ""),
                "published_at": snippet.get("publishedAt", ""),
                "privacy_status": status.get("privacyStatus", ""),
                "podcast_status": status.get("podcastStatus", ""),
                "item_count": to_int(content_details.get("itemCount")),
                "thumbnail_url": best_thumbnail(snippet.get("thumbnails")),
            }
        )

    return rows


def fetch_playlist_video_rows(
    api: YouTubeDataAPI,
    context: ChannelContext,
    playlist_row: dict[str, Any],
) -> list[dict[str, Any]]:
    playlist_id = playlist_row["playlist_id"]
    items = api.list_all(
        "playlistItems",
        {
            "part": PLAYLIST_ITEM_PARTS,
            "playlistId": playlist_id,
        },
    )
    rows: list[dict[str, Any]] = []

    for fallback_position, item in enumerate(items):
        snippet = item.get("snippet", {})
        content_details = item.get("contentDetails", {})
        status = item.get("status", {})
        resource_id = snippet.get("resourceId", {})
        video_id = content_details.get("videoId") or resource_id.get("videoId") or ""

        rows.append(
            {
                "channel": context.source,
                "channel_stem": context.channel_stem,
                "channel_id": context.channel_id,
                "channel_title": context.channel_title,
                "source_section": playlist_row["source_section"],
                "playlist_id": playlist_id,
                "playlist_title": playlist_row["playlist_title"],
                "playlist_url": playlist_row["playlist_url"],
                "playlist_item_id": item.get("id", ""),
                "playlist_position": snippet.get("position", fallback_position) + 1,
                "playlist_item_published_at": snippet.get("publishedAt", ""),
                "video_id": video_id,
                "video_title": snippet.get("title", ""),
                "video_description": snippet.get("description", ""),
                "video_url": make_video_url(video_id) if video_id else "",
                "video_owner_channel_id": snippet.get("videoOwnerChannelId", ""),
                "video_owner_channel_title": snippet.get("videoOwnerChannelTitle", ""),
                "video_published_at": content_details.get("videoPublishedAt", ""),
                "privacy_status": status.get("privacyStatus", ""),
                "thumbnail_url": best_thumbnail(snippet.get("thumbnails")),
            }
        )

    return rows


def process_channel(
    api: YouTubeDataAPI,
    source: str,
    output_dir: Path,
    seconds_before_cutoff: int,
    seconds_after_cutoff: int,
    include_playlist_items: bool,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    failed_items: list[dict[str, Any]] = []
    empty_items: list[dict[str, Any]] = []
    skipped_items: list[dict[str, Any]] = []
    membership_rows: list[dict[str, Any]] = []

    context, channel_item = resolve_channel(api, source)
    channel_dir = output_dir

    print("\n==============================")
    print(f"Channel: {source}")
    print(f"Resolved: {context.channel_title} ({context.channel_id})")
    print(f"File stem: {context.channel_stem}")
    print("==============================")

    channel_metadata = [build_channel_metadata_row(context, channel_item)]
    save_rows(channel_metadata, channel_dir / f"{context.channel_stem}__channel", "channel")

    print("\nProcessing uploads playlist")
    upload_items = fetch_upload_playlist_items(api, context.uploads_playlist_id)
    video_ids = [
        item.get("contentDetails", {}).get("videoId")
        or item.get("snippet", {}).get("resourceId", {}).get("videoId")
        for item in upload_items
    ]
    video_ids = [video_id for video_id in video_ids if video_id]
    video_details = fetch_video_details(api, video_ids)
    upload_rows = build_upload_rows(
        context=context,
        playlist_items=upload_items,
        video_details=video_details,
        seconds_before_cutoff=seconds_before_cutoff,
        seconds_after_cutoff=seconds_after_cutoff,
    )

    if upload_rows:
        save_rows(upload_rows, channel_dir / f"{context.channel_stem}__uploads", "uploads")

        section_rows = {
            "videos": [row for row in upload_rows if row["api_section"] == "videos"],
            "streams": [row for row in upload_rows if row["api_section"] == "streams"],
            "shorts": [row for row in upload_rows if row["api_section"] == "shorts"],
        }

        for section, rows in section_rows.items():
            if rows:
                save_rows(rows, channel_dir / f"{context.channel_stem}__{section}", section)
                print(f"Saved {section}: {len(rows)} row(s)")
            else:
                empty_items.append(
                    {
                        "channel": source,
                        "section": section,
                        "url": context.channel_url.rstrip("/") + f"/{section}",
                        "reason": "No API-classified rows",
                    }
                )
    else:
        empty_items.append(
            {
                "channel": source,
                "section": "uploads",
                "url": make_playlist_url(context.uploads_playlist_id),
                "reason": "No public uploads returned by API",
            }
        )

    print("\nProcessing playlists and podcasts")
    playlist_api_items = fetch_channel_playlists(api, context.channel_id)
    playlist_rows = build_playlist_rows(context, playlist_api_items)

    rows_by_section = {
        "playlists": [row for row in playlist_rows if row["source_section"] == "playlists"],
        "podcasts": [row for row in playlist_rows if row["source_section"] == "podcasts"],
    }

    for section, rows in rows_by_section.items():
        if rows:
            save_rows(rows, channel_dir / f"{context.channel_stem}__{section}", section)
            print(f"Saved {section}: {len(rows)} row(s)")
        else:
            empty_items.append(
                {
                    "channel": source,
                    "section": section,
                    "url": context.channel_url.rstrip("/") + f"/{section}",
                    "reason": "No public API rows",
                }
            )

    if include_playlist_items:
        for playlist_row in playlist_rows:
            playlist_id = playlist_row["playlist_id"]
            playlist_title = playlist_row["playlist_title"]
            source_section = playlist_row["source_section"]
            playlist_part = safe_filename(playlist_id or playlist_title)
            base_path = channel_dir / f"{context.channel_stem}__{source_section}__{playlist_part}"

            print(f"  {source_section}: {playlist_title}")

            try:
                item_rows = fetch_playlist_video_rows(api, context, playlist_row)

                if item_rows:
                    save_rows(item_rows, base_path, "playlist_items")
                    membership_rows.extend(item_rows)
                else:
                    empty_items.append(
                        {
                            "channel": source,
                            "section": f"{source_section}_items",
                            "url": playlist_row["playlist_url"],
                            "reason": "No public playlist items returned by API",
                        }
                    )

            except Exception as exc:
                failed_items.append(
                    {
                        "channel": source,
                        "section": f"{source_section}_items",
                        "url": playlist_row["playlist_url"],
                        "error": str(exc),
                    }
                )

    skipped_items.append(
        {
            "channel": source,
            "section": "community",
            "url": context.channel_url.rstrip("/") + "/community",
            "reason": "YouTube Data API v3 has no public community posts endpoint",
        }
    )

    return failed_items, empty_items, skipped_items, membership_rows


def save_logs(
    output_dir: Path,
    failed_items: list[dict[str, Any]],
    empty_items: list[dict[str, Any]],
    skipped_items: list[dict[str, Any]],
    membership_rows: list[dict[str, Any]],
) -> None:
    if membership_rows:
        path = save_rows(
            membership_rows,
            output_dir / "playlist_video_membership",
            "membership",
        )
        print(f"\nPlaylist/podcast membership saved to: {path}")
        print(f"Membership rows: {len(membership_rows)}")

    if failed_items:
        path = save_rows(failed_items, output_dir / "failed_sections", "failed")
        print(f"\nFailure log saved to: {path}")

    if empty_items:
        path = save_rows(empty_items, output_dir / "empty_sections", "empty")
        print(f"Empty sections log saved to: {path}")

    if skipped_items:
        path = save_rows(skipped_items, output_dir / "skipped_sections", "skipped")
        print(f"Skipped sections log saved to: {path}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        default=DEFAULT_CHANNELS_FILE,
        help=f"Text file with channel references. Default: {DEFAULT_CHANNELS_FILE}",
    )
    parser.add_argument(
        "--output-dir",
        default=DEFAULT_OUTPUT_DIR,
        help=f"Directory for output files. Default: {DEFAULT_OUTPUT_DIR}",
    )
    parser.add_argument(
        "--api-key-file",
        default=DEFAULT_API_KEY_FILE,
        help=f"File containing YouTube Data API key. Default: {DEFAULT_API_KEY_FILE}",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.0,
        help="Optional delay between API requests in seconds.",
    )
    parser.add_argument(
        "--shorts-before-cutoff-seconds",
        type=int,
        default=DEFAULT_SHORTS_MAX_SECONDS_BEFORE_CUTOFF,
        help="Duration threshold for possible Shorts before 2024-10-15. Default: 60.",
    )
    parser.add_argument(
        "--shorts-after-cutoff-seconds",
        type=int,
        default=DEFAULT_SHORTS_MAX_SECONDS_AFTER_CUTOFF,
        help="Duration threshold for possible Shorts on/after 2024-10-15. Default: 180.",
    )
    parser.add_argument(
        "--skip-playlist-items",
        action="store_true",
        help="Only save playlist metadata, not the videos inside each playlist.",
    )
    parser.add_argument(
        "--limit-channels",
        type=int,
        default=0,
        help="Process only the first N channels. Useful for testing.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    if args.delay < 0 or args.limit_channels < 0:
        parser.error("--delay and --limit-channels must be nonnegative")
    if min(args.shorts_before_cutoff_seconds, args.shorts_after_cutoff_seconds) < 0:
        parser.error("Shorts duration thresholds must be nonnegative")
    if not Path(args.input).is_file():
        parser.error(f"Channels file not found: {args.input}. Copy channels.example.txt to channels.txt first.")

    api_key = read_api_key(args.api_key_file) or os.getenv("YOUTUBE_API_KEY")
    if not api_key:
        print(
            "YouTube API key not found. Expected key file "
            f"{args.api_key_file} or environment variable YOUTUBE_API_KEY.",
            file=sys.stderr,
        )
        return 2

    channels = load_channels(args.input)
    if args.limit_channels:
        channels = channels[: args.limit_channels]

    if not channels:
        print("No channels found in the channels file.")
        return 0

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    api = YouTubeDataAPI(api_key=api_key, delay=args.delay)

    print(f"Loaded {len(channels)} channel(s)")
    print(f"Output directory: {output_dir}")
    print("Using YouTube Data API v3")
    print("Sections: uploads -> videos/streams/possible shorts, playlists, podcasts")

    failed_items: list[dict[str, Any]] = []
    empty_items: list[dict[str, Any]] = []
    skipped_items: list[dict[str, Any]] = []
    membership_rows: list[dict[str, Any]] = []

    for channel in channels:
        try:
            failed, empty, skipped, membership = process_channel(
                api=api,
                source=channel,
                output_dir=output_dir,
                seconds_before_cutoff=args.shorts_before_cutoff_seconds,
                seconds_after_cutoff=args.shorts_after_cutoff_seconds,
                include_playlist_items=not args.skip_playlist_items,
            )
            failed_items.extend(failed)
            empty_items.extend(empty)
            skipped_items.extend(skipped)
            membership_rows.extend(membership)
        except Exception as exc:
            print(f"Unexpected error for channel '{channel}': {exc}")
            failed_items.append(
                {
                    "channel": channel,
                    "section": "channel",
                    "url": normalize_channel_source(channel),
                    "error": str(exc),
                }
            )

    save_logs(
        output_dir=output_dir,
        failed_items=failed_items,
        empty_items=empty_items,
        skipped_items=skipped_items,
        membership_rows=membership_rows,
    )

    print("\nProcessing complete.")
    return 1 if failed_items else 0


if __name__ == "__main__":
    raise SystemExit(main())
