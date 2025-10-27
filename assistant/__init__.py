"""Core package for AI Executive Assistant."""

from .azure import (
    AzureAppConfig,
    build_authorization_url,
    build_token_request_payload,
    calendar_events_url,
    chat_messages_url,
    default_graph_scopes,
    graph_request_headers,
    meeting_transcripts_url,
    teams_chat_message_url,
    teams_online_meetings_url,
    user_messages_url,
)
from .meeting import AgendaItem, Meeting, generate_agenda
from .prep import (
    MessageSnippet,
    PreMeetingBrief,
    collect_pre_meeting_brief,
    generate_post_meeting_summary,
    PostMeetingSummary,
)

__all__ = [
    "AgendaItem",
    "AzureAppConfig",
    "MessageSnippet",
    "Meeting",
    "PreMeetingBrief",
    "build_authorization_url",
    "build_token_request_payload",
    "calendar_events_url",
    "chat_messages_url",
    "default_graph_scopes",
    "collect_pre_meeting_brief",
    "generate_agenda",
    "generate_post_meeting_summary",
    "graph_request_headers",
    "meeting_transcripts_url",
    "PostMeetingSummary",
    "teams_chat_message_url",
    "teams_online_meetings_url",
    "user_messages_url",
]
