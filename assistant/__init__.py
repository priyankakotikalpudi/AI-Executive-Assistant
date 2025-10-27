"""Core package for AI Executive Assistant."""

from .azure import (
    AzureAppConfig,
    build_authorization_url,
    build_token_request_payload,
    calendar_events_url,
    default_graph_scopes,
    graph_request_headers,
    teams_chat_message_url,
    teams_online_meetings_url,
)
from .meeting import AgendaItem, Meeting, generate_agenda

__all__ = [
    "AgendaItem",
    "AzureAppConfig",
    "Meeting",
    "build_authorization_url",
    "build_token_request_payload",
    "calendar_events_url",
    "default_graph_scopes",
    "generate_agenda",
    "graph_request_headers",
    "teams_chat_message_url",
    "teams_online_meetings_url",
]
