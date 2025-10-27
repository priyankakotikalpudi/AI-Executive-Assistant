"""Meeting preparation utilities."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Iterable, List


@dataclass(frozen=True)
class AgendaItem:
    """Represents an individual agenda entry for a meeting."""

    title: str
    owner: str | None = None
    duration_minutes: int | None = None


@dataclass
class Meeting:
    """Data model capturing the details needed to prepare a meeting."""

    title: str
    meeting_date: date
    duration_minutes: int
    participants: List[str] = field(default_factory=list)
    topics: List[str] = field(default_factory=list)

    def __post_init__(self) -> None:
        if self.duration_minutes <= 0:
            raise ValueError("Meeting duration must be positive")
        if not self.title.strip():
            raise ValueError("Meeting title cannot be empty")
        # Normalise participant names to title case for consistency.
        self.participants = [participant.strip().title() for participant in self.participants if participant.strip()]
        self.topics = [topic.strip() for topic in self.topics if topic.strip()]

    @property
    def end_time_offset(self) -> timedelta:
        """Return the duration as a ``timedelta`` for convenience."""

        return timedelta(minutes=self.duration_minutes)

    def primary_facilitator(self) -> str | None:
        """Return the first participant to serve as the default facilitator."""

        return self.participants[0] if self.participants else None


def generate_agenda(meeting: Meeting, *, buffer_minutes: int = 5) -> List[AgendaItem]:
    """Generate a structured agenda for a meeting.

    Args:
        meeting: The meeting configuration to use.
        buffer_minutes: Minutes reserved at the start and end for administration.

    Returns:
        A list of :class:`AgendaItem` entries ordered as they should appear in the meeting.

    The function allocates time evenly between topics while reserving optional
    buffers for introductions and wrap-up. If no topics are provided a default
    catch-all agenda item is produced.
    """

    if meeting.duration_minutes <= buffer_minutes * 2:
        raise ValueError("Meeting duration too short to allocate buffers")

    agenda: List[AgendaItem] = []

    facilitator = meeting.primary_facilitator()
    if facilitator:
        agenda.append(
            AgendaItem(
                title="Introductions & Objectives",
                owner=facilitator,
                duration_minutes=buffer_minutes,
            )
        )
    elif buffer_minutes:
        agenda.append(AgendaItem(title="Introductions & Objectives", duration_minutes=buffer_minutes))

    core_topics = meeting.topics or ["General Discussion"]
    available_minutes = meeting.duration_minutes - buffer_minutes * 2
    topic_duration = max(5, available_minutes // len(core_topics))

    for topic in core_topics:
        owner = _find_topic_owner(topic, meeting.participants)
        agenda.append(AgendaItem(title=topic, owner=owner, duration_minutes=topic_duration))

    agenda.append(AgendaItem(title="Wrap-up & Next Steps", owner=facilitator, duration_minutes=buffer_minutes))
    return agenda


def _find_topic_owner(topic: str, participants: Iterable[str]) -> str | None:
    """Derive an agenda owner for the topic.

    The heuristic selects the first participant whose name appears in the
    topic description; otherwise the primary facilitator is used.
    """

    topic_lower = topic.lower()
    for participant in participants:
        if participant.lower() in topic_lower:
            return participant
    return next(iter(participants), None)
