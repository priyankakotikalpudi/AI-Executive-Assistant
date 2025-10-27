"""Pre-meeting preparation utilities built on top of Microsoft Graph."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta, timezone
from html.parser import HTMLParser
from typing import Iterable, Mapping, MutableMapping, Protocol, Sequence

from .azure import chat_messages_url, meeting_transcripts_url, user_messages_url
from .meeting import AgendaItem, Meeting, generate_agenda


class GraphClient(Protocol):
    """Protocol describing the subset of Microsoft Graph used by the assistant."""

    def get(
        self, url: str, params: Mapping[str, str] | None = None
    ) -> Mapping[str, object]:
        """Return the JSON body of a Graph GET request."""


@dataclass(frozen=True)
class MessageSnippet:
    """Lightweight representation of a message or transcript fragment."""

    id: str
    source: str
    subject: str
    summary: str
    owner: str | None = None
    url: str | None = None


@dataclass(frozen=True)
class PreMeetingBrief:
    """Structured pre-meeting preparation content."""

    meeting: Meeting
    agenda: Sequence[AgendaItem]
    highlights: Sequence[str]
    sources: Sequence[MessageSnippet]


class _HTMLStripper(HTMLParser):
    """Simple HTML stripper used to derive message previews."""

    def __init__(self) -> None:
        super().__init__()
        self._chunks: list[str] = []

    def handle_data(self, data: str) -> None:  # pragma: no cover - HTMLParser callback
        if data:
            self._chunks.append(data)

    def stripped(self) -> str:
        return "".join(self._chunks)


def collect_pre_meeting_brief(
    meeting: Meeting,
    graph_client: GraphClient,
    *,
    outlook_user: str = "me",
    chat_ids: Sequence[str] | None = None,
    meeting_ids: Sequence[str] | None = None,
    lookback_days: int = 14,
    max_items: int = 20,
) -> PreMeetingBrief:
    """Build a briefing document by aggregating data from Microsoft Graph."""

    agenda = generate_agenda(meeting)
    search_terms = _derive_search_terms(meeting)
    cutoff_date = meeting.meeting_date - timedelta(days=lookback_days)
    cutoff = datetime.combine(cutoff_date, time.min, tzinfo=timezone.utc)

    sources: list[MessageSnippet] = []

    sources.extend(
        _collect_outlook_messages(
            graph_client,
            outlook_user=outlook_user,
            search_terms=search_terms,
            cutoff=cutoff,
            max_items=max_items,
        )
    )

    if chat_ids:
        sources.extend(
            _collect_chat_messages(
                graph_client,
                chat_ids=chat_ids,
                cutoff=cutoff,
                max_items=max_items,
            )
        )

    if meeting_ids:
        sources.extend(
            _collect_transcripts(
                graph_client,
                meeting_ids=meeting_ids,
                cutoff=cutoff,
                max_items=max_items,
            )
        )

    unique_sources = _deduplicate_sources(sources)
    highlights = _summarise_highlights(unique_sources, meeting)

    return PreMeetingBrief(
        meeting=meeting,
        agenda=agenda,
        highlights=highlights,
        sources=tuple(unique_sources.values()),
    )


def _collect_outlook_messages(
    graph_client: GraphClient,
    *,
    outlook_user: str,
    search_terms: Sequence[str],
    cutoff: datetime,
    max_items: int,
) -> list[MessageSnippet]:
    snippets: list[MessageSnippet] = []
    endpoint = user_messages_url(outlook_user)

    for term in search_terms:
        if len(snippets) >= max_items:
            break
        params = {
            "$top": str(max_items),
            "$search": f'"{term}"',
        }
        payload = graph_client.get(endpoint, params=params)
        for item in payload.get("value", []):
            if len(snippets) >= max_items:
                break
            received = _parse_graph_datetime(item.get("receivedDateTime"))
            if received and received < cutoff:
                continue
            snippet = _graph_item_to_snippet(
                item,
                source="Outlook Mail",
                subject=item.get("subject") or term,
                owner=_extract_sender(item),
            )
            if snippet:
                snippets.append(snippet)
    return snippets


def _collect_chat_messages(
    graph_client: GraphClient,
    *,
    chat_ids: Sequence[str],
    cutoff: datetime,
    max_items: int,
) -> list[MessageSnippet]:
    snippets: list[MessageSnippet] = []
    for chat_id in chat_ids:
        if len(snippets) >= max_items:
            break
        payload = graph_client.get(
            chat_messages_url(chat_id), params={"$top": str(max_items)}
        )
        for item in payload.get("value", []):
            if len(snippets) >= max_items:
                break
            created = _parse_graph_datetime(item.get("createdDateTime"))
            if created and created < cutoff:
                continue
            snippet = _graph_item_to_snippet(
                item,
                source="Teams Chat",
                subject=item.get("summary") or "Teams message",
                owner=_extract_sender(item),
            )
            if snippet:
                snippets.append(snippet)
    return snippets


def _collect_transcripts(
    graph_client: GraphClient,
    *,
    meeting_ids: Sequence[str],
    cutoff: datetime,
    max_items: int,
) -> list[MessageSnippet]:
    snippets: list[MessageSnippet] = []
    for meeting_id in meeting_ids:
        if len(snippets) >= max_items:
            break
        payload = graph_client.get(
            meeting_transcripts_url(meeting_id), params={"$top": str(max_items)}
        )
        for item in payload.get("value", []):
            if len(snippets) >= max_items:
                break
            creation = _parse_graph_datetime(item.get("createdDateTime"))
            if creation and creation < cutoff:
                continue
            transcript = item.get("content") or item.get("transcriptContent")
            snippet = _graph_item_to_snippet(
                item,
                source="Teams Transcript",
                subject=item.get("title") or "Meeting transcript",
                owner=item.get("speaker"),
                body=transcript,
            )
            if snippet:
                snippets.append(snippet)
    return snippets


def _graph_item_to_snippet(
    item: Mapping[str, object],
    *,
    source: str,
    subject: str,
    owner: str | None,
    body: str | None = None,
) -> MessageSnippet | None:
    identifier = str(item.get("id")) if item.get("id") is not None else None
    if not identifier:
        return None

    preview = body or item.get("bodyPreview") or item.get("summary") or ""
    cleaned_preview = _strip_html(str(preview))
    summary = _truncate(cleaned_preview)

    web_link = item.get("webLink") or item.get("link")
    return MessageSnippet(
        id=identifier,
        source=source,
        subject=str(subject),
        summary=summary,
        owner=str(owner) if owner else None,
        url=str(web_link) if web_link else None,
    )


def _extract_sender(item: Mapping[str, object]) -> str | None:
    sender = item.get("from") or item.get("sender")
    if isinstance(sender, Mapping):
        email_address = sender.get("emailAddress")
        if isinstance(email_address, Mapping):
            return str(email_address.get("name") or email_address.get("address"))
        if sender.get("user") and isinstance(sender["user"], Mapping):
            return str(sender["user"].get("displayName"))
        if sender.get("application") and isinstance(sender["application"], Mapping):
            return str(sender["application"].get("displayName"))
    return None


def _parse_graph_datetime(value: object) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None


def _strip_html(value: str) -> str:
    if "<" not in value:
        return " ".join(value.split())
    parser = _HTMLStripper()
    parser.feed(value)
    parser.close()
    return " ".join(parser.stripped().split())


def _truncate(value: str, *, limit: int = 240) -> str:
    if len(value) <= limit:
        return value
    truncated = value[:limit].rsplit(" ", 1)[0]
    return truncated.rstrip(".,;") + "…"


def _deduplicate_sources(
    snippets: Iterable[MessageSnippet],
) -> MutableMapping[str, MessageSnippet]:
    unique: MutableMapping[str, MessageSnippet] = {}
    for snippet in snippets:
        key = f"{snippet.source}:{snippet.id}"
        if key not in unique:
            unique[key] = snippet
    return unique


def _summarise_highlights(
    sources: Mapping[str, MessageSnippet],
    meeting: Meeting,
) -> list[str]:
    highlights: list[str] = []
    if meeting.topics:
        highlights.append(
            "Focus topics: " + ", ".join(meeting.topics)
        )
    for snippet in sources.values():
        owner = f"{snippet.owner}: " if snippet.owner else ""
        highlights.append(f"{snippet.source} – {owner}{snippet.summary}")
    return highlights


def _derive_search_terms(meeting: Meeting) -> Sequence[str]:
    ordered_terms: list[str] = []
    seen: set[str] = set()

    def _add(term: str | None) -> None:
        if not term:
            return
        cleaned = term.strip()
        if not cleaned:
            return
        key = cleaned.lower()
        if key in seen:
            return
        seen.add(key)
        ordered_terms.append(cleaned)

    _add(meeting.title)
    for topic in meeting.topics:
        _add(topic)
    for participant in meeting.participants:
        _add(participant)
    return tuple(ordered_terms)

