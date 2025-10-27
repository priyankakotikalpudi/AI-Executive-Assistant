from datetime import date

from assistant.meeting import Meeting
from assistant.prep import collect_pre_meeting_brief


class FakeGraphClient:
    def __init__(self, responses):
        self.responses = responses
        self.requests = []

    def get(self, url, params=None):
        key = (url, tuple(sorted((params or {}).items())))
        self.requests.append(key)
        return self.responses.get(key, {"value": []})


def build_key(url, **params):
    return (url, tuple(sorted(params.items())))


def test_collect_pre_meeting_brief_aggregates_sources():
    meeting = Meeting(
        title="Product Launch Sync",
        meeting_date=date(2024, 8, 20),
        duration_minutes=60,
        participants=["alex johnson", "Jordan"],
        topics=["Launch risks", "Marketing timeline"],
    )

    outlook_url = "https://graph.microsoft.com/v1.0/users/me/messages"
    chat_url = "https://graph.microsoft.com/v1.0/chats/chat-id/messages"
    transcript_url = (
        "https://graph.microsoft.com/v1.0/communications/onlineMeetings/meeting-id/transcripts"
    )

    responses = {
        build_key(
            outlook_url,
            **{"$search": '"Product Launch Sync"', "$top": "5"},
        ): {
            "value": [
                {
                    "id": "mail-1",
                    "subject": "Launch risk review",
                    "receivedDateTime": "2024-08-15T10:00:00Z",
                    "bodyPreview": "Finalize mitigation plan before executive review.",
                    "from": {"emailAddress": {"name": "Jordan"}},
                    "webLink": "https://outlook.office.com/mail-1",
                }
            ]
        },
        build_key(chat_url, **{"$top": "5"}): {
            "value": [
                {
                    "id": "chat-1",
                    "summary": "Alex posted updated launch checklist",
                    "createdDateTime": "2024-08-16T08:30:00Z",
                    "bodyPreview": "Checklist now tracks marketing approvals.",
                    "from": {"user": {"displayName": "Alex Johnson"}},
                }
            ]
        },
        build_key(transcript_url, **{"$top": "5"}): {
            "value": [
                {
                    "id": "transcript-1",
                    "title": "Notes from beta feedback readout",
                    "createdDateTime": "2024-08-12T17:00:00Z",
                    "content": "<p>Focus on customer onboarding messaging improvements.</p>",
                    "speaker": "Jordan",
                }
            ]
        },
    }

    client = FakeGraphClient(responses)

    brief = collect_pre_meeting_brief(
        meeting,
        client,
        chat_ids=["chat-id"],
        meeting_ids=["meeting-id"],
        lookback_days=10,
        max_items=5,
    )

    assert len(brief.agenda) == 4
    assert any(source.source == "Outlook Mail" for source in brief.sources)
    assert any(source.source == "Teams Chat" for source in brief.sources)
    assert any(source.source == "Teams Transcript" for source in brief.sources)
    assert brief.highlights[0].startswith("Focus topics: Launch risks, Marketing timeline")
    assert any("Jordan" in highlight for highlight in brief.highlights)


def test_collect_pre_meeting_brief_filters_old_items():
    meeting = Meeting(
        title="Finance Review",
        meeting_date=date(2024, 7, 15),
        duration_minutes=45,
        participants=["Morgan"],
        topics=["Budget variance"],
    )

    outlook_url = "https://graph.microsoft.com/v1.0/users/me/messages"

    responses = {
        build_key(
            outlook_url,
            **{"$search": '"Finance Review"', "$top": "10"},
        ): {
            "value": [
                {
                    "id": "mail-recent",
                    "subject": "Budget variance summary",
                    "receivedDateTime": "2024-07-10T09:00:00Z",
                    "bodyPreview": "Variance now within acceptable tolerance thanks to new controls.",
                    "from": {"emailAddress": {"name": "Morgan"}},
                },
                {
                    "id": "mail-old",
                    "subject": "Old update",
                    "receivedDateTime": "2024-06-01T09:00:00Z",
                    "bodyPreview": "Superseded information.",
                    "from": {"emailAddress": {"name": "Morgan"}},
                },
            ]
        }
    }

    client = FakeGraphClient(responses)

    brief = collect_pre_meeting_brief(
        meeting,
        client,
        lookback_days=20,
        max_items=10,
    )

    assert len(brief.sources) == 1
    assert brief.sources[0].id == "mail-recent"
    assert "Variance now within acceptable tolerance" in brief.highlights[1]


def test_collect_pre_meeting_brief_deduplicates_by_source():
    meeting = Meeting(
        title="Design Review",
        meeting_date=date(2024, 6, 20),
        duration_minutes=30,
        participants=["Casey"],
        topics=["UI polish"],
    )

    outlook_url = "https://graph.microsoft.com/v1.0/users/me/messages"

    responses = {
        build_key(
            outlook_url,
            **{"$search": '"Design Review"', "$top": "2"},
        ): {
            "value": [
                {
                    "id": "shared-id",
                    "subject": "Latest mock-ups",
                    "receivedDateTime": "2024-06-18T09:00:00Z",
                    "bodyPreview": "Mock-ups ready for sign-off.",
                    "from": {"emailAddress": {"name": "Casey"}},
                }
            ]
        },
        build_key(
            outlook_url,
            **{"$search": '"Casey"', "$top": "2"},
        ): {
            "value": [
                {
                    "id": "shared-id",
                    "subject": "Reminder",
                    "receivedDateTime": "2024-06-18T11:00:00Z",
                    "bodyPreview": "Same message surfaced by another search term.",
                    "from": {"emailAddress": {"name": "Casey"}},
                }
            ]
        },
    }

    client = FakeGraphClient(responses)

    brief = collect_pre_meeting_brief(
        meeting,
        client,
        lookback_days=5,
        max_items=2,
    )

    assert len(brief.sources) == 1
    assert brief.sources[0].subject == "Latest mock-ups"


def test_collect_pre_meeting_brief_respects_max_items():
    meeting = Meeting(
        title="Roadmap Alignment",
        meeting_date=date(2024, 9, 5),
        duration_minutes=45,
        participants=["Taylor"],
        topics=["Dependencies"],
    )

    outlook_url = "https://graph.microsoft.com/v1.0/users/me/messages"

    responses = {
        build_key(
            outlook_url,
            **{"$search": '"Roadmap Alignment"', "$top": "3"},
        ): {
            "value": [
                {
                    "id": f"mail-{idx}",
                    "subject": f"Thread {idx}",
                    "receivedDateTime": "2024-09-01T09:00:00Z",
                    "bodyPreview": "Update",
                    "from": {"emailAddress": {"name": "Taylor"}},
                }
                for idx in range(5)
            ]
        }
    }

    client = FakeGraphClient(responses)

    brief = collect_pre_meeting_brief(
        meeting,
        client,
        lookback_days=10,
        max_items=3,
    )

    assert len(brief.sources) == 3
    assert all(source.id in {"mail-0", "mail-1", "mail-2"} for source in brief.sources)
