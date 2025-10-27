from datetime import date

import pytest

from assistant.meeting import AgendaItem, Meeting, generate_agenda


def test_meeting_validation():
    with pytest.raises(ValueError):
        Meeting(title="", meeting_date=date.today(), duration_minutes=30)

    with pytest.raises(ValueError):
        Meeting(title="Weekly Sync", meeting_date=date.today(), duration_minutes=0)


def test_generate_agenda_with_topics_and_participants():
    meeting = Meeting(
        title="Product Planning",
        meeting_date=date(2024, 8, 1),
        duration_minutes=60,
        participants=["alex johnson", "Jordan"],
        topics=["Alex Johnson - roadmap review", "Budget alignment"],
    )

    agenda = generate_agenda(meeting)

    assert agenda[0] == AgendaItem(
        title="Introductions & Objectives", owner="Alex Johnson", duration_minutes=5
    )
    assert agenda[-1] == AgendaItem(
        title="Wrap-up & Next Steps", owner="Alex Johnson", duration_minutes=5
    )
    core_items = agenda[1:-1]
    assert len(core_items) == 2
    assert all(item.duration_minutes >= 5 for item in core_items)
    assert core_items[0].owner == "Alex Johnson"
    assert core_items[1].owner == "Alex Johnson"


def test_generate_agenda_without_topics():
    meeting = Meeting(
        title="Team Catch-up",
        meeting_date=date(2024, 8, 2),
        duration_minutes=45,
        participants=[],
        topics=[],
    )

    agenda = generate_agenda(meeting, buffer_minutes=5)

    assert agenda[0].title == "Introductions & Objectives"
    assert agenda[0].owner is None
    assert agenda[1].title == "General Discussion"
    assert agenda[-1].title == "Wrap-up & Next Steps"


def test_generate_agenda_raises_when_duration_too_short():
    meeting = Meeting(
        title="Quick Sync",
        meeting_date=date(2024, 8, 3),
        duration_minutes=10,
        participants=["Chris"],
        topics=["Update"],
    )

    with pytest.raises(ValueError):
        generate_agenda(meeting, buffer_minutes=6)
