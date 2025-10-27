from datetime import date

import pytest

from assistant.meeting import Meeting
from assistant.prep import generate_post_meeting_summary


def test_generate_post_meeting_summary_extracts_actions_and_discussions():
    meeting = Meeting(
        title="Launch Readout",
        meeting_date=date(2024, 9, 1),
        duration_minutes=60,
        participants=["alex", "Jordan", "Sam"],
        topics=["Launch plan", "Beta feedback"],
    )

    transcript = """
    Alex: Discussed launch plan timeline adjustments and dependencies.
    Jordan: Action item - send updated spec by Thursday.
    Sam: We agreed to capture beta feedback insights for the leadership deck.
    Jordan: Next steps include coordinating with analytics on adoption metrics.
    """

    summary = generate_post_meeting_summary(meeting, transcript)

    assert summary.discussion_items
    assert any("Alex" in item for item in summary.discussion_items)
    assert any("beta feedback" in item.lower() for item in summary.discussion_items)
    assert any("Action item" in item for item in summary.action_items)
    assert any("Next steps" in item for item in summary.action_items)
    assert meeting.title in summary.email_subject
    assert "Agenda reviewed:" in summary.email_body
    assert "Discussion highlights:" in summary.email_body
    assert "Action items:" in summary.email_body


def test_generate_post_meeting_summary_handles_missing_actions():
    meeting = Meeting(
        title="Weekly Sync",
        meeting_date=date(2024, 7, 15),
        duration_minutes=45,
        participants=["Taylor"],
        topics=["Status"],
    )

    transcript = "Taylor: Discussed current sprint progress and blockers."

    summary = generate_post_meeting_summary(
        meeting,
        transcript,
        max_action_items=0,
        additional_notes=["Stakeholders will receive slides tomorrow"],
    )

    assert summary.action_items == ()
    assert any("slides tomorrow" in item.lower() for item in summary.discussion_items)
    assert "No explicit action items" in summary.email_body


def test_generate_post_meeting_summary_validates_inputs():
    meeting = Meeting(
        title="Review",
        meeting_date=date(2024, 5, 20),
        duration_minutes=30,
    )

    with pytest.raises(ValueError):
        generate_post_meeting_summary(meeting, "   ")

    with pytest.raises(ValueError):
        generate_post_meeting_summary(meeting, "Notes", max_discussion_items=-1)

    with pytest.raises(ValueError):
        generate_post_meeting_summary(meeting, "Notes", max_action_items=-2)
