"""Tests for the Azure configuration helpers."""

from assistant.azure import (
    AzureAppConfig,
    build_authorization_url,
    build_token_request_payload,
    calendar_events_url,
    default_graph_scopes,
    graph_request_headers,
    teams_chat_message_url,
    teams_online_meetings_url,
)


def test_default_graph_scopes_include_expected_permissions() -> None:
    scopes = default_graph_scopes()
    assert "https://graph.microsoft.com/Calendars.ReadWrite" in scopes
    assert "https://graph.microsoft.com/OnlineMeetings.ReadWrite" in scopes
    assert scopes[-1] == "offline_access"


def test_authorization_url_includes_normalised_scopes() -> None:
    config = AzureAppConfig(
        tenant_id="contoso.onmicrosoft.com",
        client_id="client-id",
        redirect_uri="https://localhost/auth",
        scopes=(
            " https://graph.microsoft.com/User.Read ",
            "https://graph.microsoft.com/User.Read",
            "https://graph.microsoft.com/Calendars.ReadWrite",
        ),
    )

    url = build_authorization_url(config, state="12345", prompt="select_account")

    assert url.startswith(
        "https://login.microsoftonline.com/contoso.onmicrosoft.com/oauth2/v2.0/authorize"
    )
    assert "scope=https%3A%2F%2Fgraph.microsoft.com%2FUser.Read+https%3A%2F%2Fgraph.microsoft.com%2FCalendars.ReadWrite" in url
    assert "prompt=select_account" in url
    assert "state=12345" in url


def test_token_request_payload_handles_client_secret_and_pkce() -> None:
    config = AzureAppConfig(
        tenant_id="contoso.onmicrosoft.com",
        client_id="client-id",
        client_secret="secret",
        redirect_uri="https://localhost/auth",
        scopes=("https://graph.microsoft.com/User.Read",),
    )

    payload = build_token_request_payload(
        config,
        authorization_code="auth-code",
        code_verifier="code-verifier",
    )

    assert payload["client_secret"] == "secret"
    assert payload["code_verifier"] == "code-verifier"
    assert payload["scope"] == "https://graph.microsoft.com/User.Read"


def test_graph_request_headers_requires_token() -> None:
    headers = graph_request_headers("token-value")
    assert headers["Authorization"] == "Bearer token-value"
    assert headers["Content-Type"] == "application/json"


def test_graph_urls_cover_outlook_and_teams_endpoints() -> None:
    assert calendar_events_url() == "https://graph.microsoft.com/v1.0/users/me/events"
    assert (
        teams_online_meetings_url("user@contoso.com")
        == "https://graph.microsoft.com/v1.0/users/user@contoso.com/onlineMeetings"
    )
    assert (
        teams_chat_message_url("team-id", "channel-id")
        == "https://graph.microsoft.com/v1.0/teams/team-id/channels/channel-id/messages"
    )
