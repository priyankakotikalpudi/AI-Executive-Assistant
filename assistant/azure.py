"""Utilities for configuring Azure AD and Microsoft Graph access."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Mapping, Sequence
from urllib.parse import urlencode

_GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


@dataclass(frozen=True)
class AzureAppConfig:
    """Configuration values for an Azure AD application registration."""

    tenant_id: str
    client_id: str
    client_secret: str | None = None
    redirect_uri: str | None = None
    scopes: Sequence[str] = field(default_factory=tuple)

    def authority(self) -> str:
        """Return the OAuth authority URL for the configured tenant."""

        tenant = self.tenant_id.strip()
        if not tenant:
            raise ValueError("tenant_id cannot be empty")
        return f"https://login.microsoftonline.com/{tenant}"

    def normalised_scopes(self) -> tuple[str, ...]:
        """Return scopes as a tuple without duplicates while preserving order."""

        seen: set[str] = set()
        ordered: list[str] = []
        for scope in self.scopes:
            cleaned = scope.strip()
            if cleaned and cleaned not in seen:
                ordered.append(cleaned)
                seen.add(cleaned)
        return tuple(ordered)


def default_graph_scopes(*, include_offline_access: bool = True) -> tuple[str, ...]:
    """Return the baseline Microsoft Graph scopes for Outlook and Teams."""

    scopes = [
        "https://graph.microsoft.com/User.Read",
        "https://graph.microsoft.com/Calendars.ReadWrite",
        "https://graph.microsoft.com/Mail.ReadWrite",
        "https://graph.microsoft.com/OnlineMeetings.ReadWrite",
        "https://graph.microsoft.com/ChannelMessage.Send",
    ]
    if include_offline_access:
        scopes.append("offline_access")
    return tuple(scopes)


def build_authorization_url(
    config: AzureAppConfig,
    *,
    state: str,
    prompt: str | None = None,
) -> str:
    """Construct the interactive authorization URL for the given configuration."""

    params: dict[str, str] = {
        "client_id": config.client_id,
        "response_type": "code",
        "redirect_uri": _require_redirect_uri(config),
        "scope": " ".join(config.normalised_scopes()),
        "state": state,
    }
    if prompt:
        params["prompt"] = prompt
    return f"{config.authority()}/oauth2/v2.0/authorize?{urlencode(params)}"


def build_token_request_payload(
    config: AzureAppConfig,
    *,
    authorization_code: str,
    code_verifier: str | None = None,
) -> Mapping[str, str]:
    """Return the payload required to exchange an auth code for tokens."""

    payload: dict[str, str] = {
        "client_id": config.client_id,
        "grant_type": "authorization_code",
        "code": authorization_code,
        "redirect_uri": _require_redirect_uri(config),
        "scope": " ".join(config.normalised_scopes()),
    }
    if config.client_secret:
        payload["client_secret"] = config.client_secret
    if code_verifier:
        payload["code_verifier"] = code_verifier
    return payload


def graph_request_headers(access_token: str) -> Mapping[str, str]:
    """Return the HTTP headers for calling Microsoft Graph with a bearer token."""

    token = access_token.strip()
    if not token:
        raise ValueError("access_token cannot be empty")
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


def calendar_events_url(user: str = "me") -> str:
    """Return the Graph API URL for listing calendar events."""

    return f"{_GRAPH_BASE_URL}/users/{user}/events"


def teams_online_meetings_url(user: str = "me") -> str:
    """Return the Graph API URL for creating or listing Teams online meetings."""

    return f"{_GRAPH_BASE_URL}/users/{user}/onlineMeetings"


def teams_chat_message_url(team_id: str, channel_id: str) -> str:
    """Return the Graph API URL for sending a message to a Teams channel."""

    _require_value(team_id, "team_id")
    _require_value(channel_id, "channel_id")
    return (
        f"{_GRAPH_BASE_URL}/teams/{team_id}/channels/{channel_id}/messages"
    )


def user_messages_url(user: str = "me") -> str:
    """Return the Graph API URL for listing Outlook messages."""

    return f"{_GRAPH_BASE_URL}/users/{user}/messages"


def chat_messages_url(chat_id: str) -> str:
    """Return the Graph API URL for retrieving Teams chat messages."""

    _require_value(chat_id, "chat_id")
    return f"{_GRAPH_BASE_URL}/chats/{chat_id}/messages"


def meeting_transcripts_url(meeting_id: str) -> str:
    """Return the Graph API URL for retrieving transcripts for a Teams meeting."""

    _require_value(meeting_id, "meeting_id")
    return f"{_GRAPH_BASE_URL}/communications/onlineMeetings/{meeting_id}/transcripts"


def _require_redirect_uri(config: AzureAppConfig) -> str:
    if not config.redirect_uri:
        raise ValueError("redirect_uri must be provided for interactive flows")
    return config.redirect_uri


def _require_value(value: str, field_name: str) -> None:
    if not value.strip():
        raise ValueError(f"{field_name} cannot be empty")
