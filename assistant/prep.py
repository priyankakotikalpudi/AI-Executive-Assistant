"""Pre-meeting preparation utilities built on top of Microsoft Graph."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta, timezone
from html.parser import HTMLParser
from typing import Iterable, Iterator, Mapping, MutableMapping, Protocol, Sequence

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


@dataclass(frozen=True)
class PostMeetingSummary:
    """Structured representation of a post-meeting follow-up package."""

    meeting: Meeting
    agenda: Sequence[AgendaItem]
    discussion_items: Sequence[str]
    action_items: Sequence[str]
    email_subject: str
    email_body: str


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

    sources: list[MessageSnippet]
    if max_items <= 0:
        sources = []
    else:
        sources = []

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


def generate_post_meeting_summary(
    meeting: Meeting,
    transcript: str,
    *,
    max_discussion_items: int = 5,
    max_action_items: int = 5,
    additional_notes: Sequence[str] | None = None,
) -> PostMeetingSummary:
    """Summarise a meeting transcript into discussion points and actions."""

    if max_discussion_items < 0:
        raise ValueError("max_discussion_items cannot be negative")
    if max_action_items < 0:
        raise ValueError("max_action_items cannot be negative")
    if not transcript or not transcript.strip():
        raise ValueError("transcript cannot be empty")

    agenda = generate_agenda(meeting)
    discussion_items, action_items = _analyse_transcript(
        transcript,
        meeting,
        max_discussion_items=max_discussion_items,
        max_action_items=max_action_items,
    )

    if additional_notes:
        for note in additional_notes:
            cleaned = note.strip()
            if cleaned:
                discussion_items.append(_normalise_statement(cleaned))

    subject = f"{meeting.title} – Summary & Actions ({meeting.meeting_date.isoformat()})"
    email_body = _build_summary_email(
        meeting,
        agenda,
        discussion_items,
        action_items,
    )

    return PostMeetingSummary(
        meeting=meeting,
        agenda=tuple(agenda),
        discussion_items=tuple(discussion_items),
        action_items=tuple(action_items),
        email_subject=subject,
        email_body=email_body,
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
        remaining = max_items - len(snippets)
        if remaining <= 0:
            break
        params = {
            "$top": str(remaining),
            "$search": f'"{term}"',
        }
        for item in _iterate_graph_collection(
            graph_client,
            endpoint,
            params=params,
            remaining=remaining,
        ):
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
        remaining = max_items - len(snippets)
        if remaining <= 0:
            break
        for item in _iterate_graph_collection(
            graph_client,
            chat_messages_url(chat_id),
            params={"$top": str(remaining)},
            remaining=remaining,
        ):
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
        remaining = max_items - len(snippets)
        if remaining <= 0:
            break
        for item in _iterate_graph_collection(
            graph_client,
            meeting_transcripts_url(meeting_id),
            params={"$top": str(remaining)},
            remaining=remaining,
        ):
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
        parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is None:
        return parsed.replace(tzinfo=timezone.utc)
    return parsed


def _iterate_graph_collection(
    graph_client: GraphClient,
    url: str,
    *,
    params: Mapping[str, str] | None,
    remaining: int,
) -> Iterator[Mapping[str, object]]:
    """Yield items from a Graph collection request, following next links."""

    next_url = url
    next_params: Mapping[str, str] | None = params
    remaining_items = remaining
    seen_links: set[str] = set()

    while remaining_items > 0 and next_url:
        payload = graph_client.get(next_url, params=next_params)
        items = payload.get("value") if isinstance(payload, Mapping) else None
        if not isinstance(items, list):
            items = []

        for item in items:
            if not isinstance(item, Mapping):
                continue
            yield item
            remaining_items -= 1
            if remaining_items <= 0:
                return

        next_link = payload.get("@odata.nextLink") if isinstance(payload, Mapping) else None
        if not next_link:
            break
        next_url = str(next_link)
        if next_url in seen_links:
            break
        seen_links.add(next_url)
        next_params = None


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


def _analyse_transcript(
    transcript: str,
    meeting: Meeting,
    *,
    max_discussion_items: int,
    max_action_items: int,
) -> tuple[list[str], list[str]]:
    sentences = _tokenise_transcript(transcript)
    topics = [topic.lower() for topic in meeting.topics]
    discussion_items: list[str] = []
    action_items: list[str] = []
    seen_discussion: set[str] = set()
    seen_actions: set[str] = set()

    action_keywords = (
        "action item",
        "action:",
        "todo",
        "to-do",
        "follow up",
        "follow-up",
        "next step",
        "next steps",
        "owner:",
        "due",
        "assign",
    )
    highlight_keywords = (
        "discussed",
        "decided",
        "agreed",
        "highlight",
        "reviewed",
        "noted",
        "update",
        "question",
    )

    for sentence in sentences:
        if len(action_items) >= max_action_items and len(discussion_items) >= max_discussion_items:
            break

        cleaned = _normalise_statement(sentence)
        if not cleaned:
            continue
        lowered = cleaned.lower()

        is_action = any(keyword in lowered for keyword in action_keywords)
        if is_action and len(action_items) < max_action_items:
            if cleaned not in seen_actions:
                action_items.append(cleaned)
                seen_actions.add(cleaned)
            continue

        topic_match = any(topic in lowered for topic in topics if topic)
        highlight_match = any(keyword in lowered for keyword in highlight_keywords)

        if (topic_match or highlight_match or len(cleaned.split()) >= 6) and len(discussion_items) < max_discussion_items:
            if cleaned not in seen_discussion:
                discussion_items.append(cleaned)
                seen_discussion.add(cleaned)

    return discussion_items, action_items


def _tokenise_transcript(transcript: str) -> list[str]:
    cleaned_text = transcript.replace("\r", "\n")
    segments = re.split(r"\n+", cleaned_text)
    sentences: list[str] = []
    for segment in segments:
        stripped = segment.strip(" -•\t")
        if not stripped:
            continue
        parts = re.split(r"(?<=[.!?])\s+", stripped)
        for part in parts:
            statement = part.strip(" -•\t")
            if statement:
                sentences.append(statement)
    return sentences


def _normalise_statement(value: str) -> str:
    compact = " ".join(value.split())
    if not compact:
        return ""

    if ":" in compact:
        speaker, remainder = compact.split(":", 1)
        speaker = speaker.strip()
        remainder = remainder.strip()
        if speaker and remainder and len(speaker.split()) <= 4:
            compact = f"{speaker.title()}: {remainder}"

    if compact.endswith(tuple(".!?")):
        return compact
    return compact + "."


def _build_summary_email(
    meeting: Meeting,
    agenda: Sequence[AgendaItem],
    discussion_items: Sequence[str],
    action_items: Sequence[str],
) -> str:
    lines: list[str] = []
    facilitator = meeting.primary_facilitator()
    meeting_date = meeting.meeting_date.strftime("%B %d, %Y")

    lines.append("Hi team,")
    lines.append("")
    lines.append(
        f"Thanks for joining {meeting.title} on {meeting_date}. Here's a quick recap and next steps."
    )
    lines.append("")

    lines.append("Agenda reviewed:")
    for item in agenda:
        duration = f" ({item.duration_minutes} min)" if item.duration_minutes else ""
        owner = f" – {item.owner}" if item.owner else ""
        lines.append(f"- {item.title}{owner}{duration}")

    discussion_section = list(discussion_items)
    if not discussion_section:
        discussion_section.append("No major discussion items were captured.")

    lines.append("")
    lines.append("Discussion highlights:")
    for highlight in discussion_section:
        lines.append(f"- {highlight}")

    action_section = list(action_items)
    if not action_section:
        action_section.append("No explicit action items recorded during the meeting.")

    lines.append("")
    lines.append("Action items:")
    for action in action_section:
        lines.append(f"- {action}")

    lines.append("")
    if facilitator:
        lines.append(f"Thanks,\n{facilitator}")
    else:
        lines.append("Thanks,\nAI Executive Assistant")

    return "\n".join(lines)

