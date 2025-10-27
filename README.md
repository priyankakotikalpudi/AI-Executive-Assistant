# AI Executive Assistant

AI Executive Assistant is an intelligent assistant designed to help users prepare for and follow up on meetings. It streamlines the process of scheduling, organizing, and managing meeting tasks, ensuring that users stay on top of their commitments and responsibilities.

## Features
- **Meeting Preparation**: Automatically generate agendas and reminders for upcoming meetings.
- **Follow-Up Tasks**: Create follow-up tasks based on meeting outcomes and decisions.
- **Integration**: Seamlessly integrates with popular calendar and task management tools.
- **User-Friendly Interface**: Intuitive design for easy navigation and use.

## Prerequisites

- Python 3.10 or newer
- `pip` for managing Python packages

## Installation Instructions
1. Clone the repository:
   ```bash
   git clone https://github.com/priyankakotikalpudi/AI-Executive-Assistant.git
   ```
2. Navigate to the project directory:
   ```bash
   cd AI-Executive-Assistant
   ```
3. (Optional) Create and activate a virtual environment:
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows use: .venv\Scripts\activate
   ```
4. Install the package in editable mode with development dependencies:
   ```bash
   pip install -e .[dev]
   ```

## Usage
The core meeting preparation utilities are provided as a Python package. You can import the meeting helpers in your own scripts or notebooks:

```python
from datetime import date

from assistant import Meeting, generate_agenda

meeting = Meeting(
    title="Weekly Sync",
    meeting_date=date(2024, 5, 1),
    duration_minutes=45,
    participants=["Alex", "Priya"],
    topics=["Roadmap review", "Team metrics"],
)

agenda = generate_agenda(meeting)
for item in agenda:
    print(item)
```

### Building a pre-meeting brief

With Microsoft Graph access configured you can combine recent Outlook messages,
Teams chats, and meeting transcripts into a single briefing:

```python
from assistant import collect_pre_meeting_brief

# graph_client should expose a .get(url, params=None) method that returns JSON.
brief = collect_pre_meeting_brief(
    meeting,
    graph_client,
    chat_ids=["19:meetingChatId"],
    meeting_ids=["MSpfx-generated-meeting-id"],
    max_items=5,  # caps how many Outlook/Teams snippets are pulled into the brief
)

for line in brief.highlights:
    print("-", line)
```

The helper deduplicates items that surface through multiple queries and respects the
`max_items` limit across Outlook, Teams chat, and transcript sources so the brief stays
focused on the most recent material.

### Running Tests

To validate the meeting preparation feature set, execute the unit tests with `pytest`:

```bash
pytest
```

## Connecting to Outlook and Teams via Azure

To enable calendar and Teams integration you must register the assistant as an application within your Azure Active Directory tenant and grant Microsoft Graph permissions.

1. **Create an app registration**
   - Sign in to the [Azure Portal](https://portal.azure.com).
   - Navigate to *Azure Active Directory → App registrations → New registration*.
   - Provide a name (for example `AI Executive Assistant`) and select the supported account type that matches your organisation.
   - Add a web redirect URI such as `https://localhost/auth` for local testing.

2. **Configure client credentials**
   - Note the generated *Application (client) ID* and *Directory (tenant) ID*.
   - If you plan to use the authorization code flow, create a **client secret** under *Certificates & secrets* and store it securely.

3. **Grant Microsoft Graph permissions**
   - Under *API permissions*, add delegated permissions for:
     - `Calendars.ReadWrite` and `Mail.ReadWrite` (Outlook calendar and mail access)
     - `OnlineMeetings.ReadWrite` (Teams meeting creation)
     - `ChannelMessage.Send` (posting to Teams channels)
     - `User.Read` (basic profile access)
   - Click *Grant admin consent* so the permissions are available to users.

4. **Configure the assistant**
   - Provide your Azure details in code using the new helpers:

    ```python
    from assistant import (
        AzureAppConfig,
        build_authorization_url,
        build_token_request_payload,
        calendar_events_url,
        chat_messages_url,
        default_graph_scopes,
        graph_request_headers,
        meeting_transcripts_url,
        teams_online_meetings_url,
        user_messages_url,
    )

     config = AzureAppConfig(
         tenant_id="contoso.onmicrosoft.com",
         client_id="<CLIENT_ID>",
         client_secret="<CLIENT_SECRET>",
         redirect_uri="https://localhost/auth",
         scopes=default_graph_scopes(),
     )

     authorization_url = build_authorization_url(config, state="state-token")
     print("Navigate to:", authorization_url)
     ```

   - After the user completes the sign-in and you receive an authorization code at your redirect URI, exchange it for tokens using `build_token_request_payload` with the HTTP client of your choice. Use `graph_request_headers` alongside `user_messages_url`, `chat_messages_url`, `meeting_transcripts_url`, `calendar_events_url`, or `teams_online_meetings_url` to call Microsoft Graph.

5. **Recommended tooling**
   - Use `msal` or `azure-identity` to handle the OAuth 2.0 flow and token caching.
   - Microsoft Graph requests can be performed with `requests` or `msgraph-core` once you supply the bearer token headers generated above.

These steps provide the Azure prerequisites needed for integrating the assistant with Outlook and Microsoft Teams.

## Contributing
1. Fork the repository.
2. Create a new branch:
   ```bash
   git checkout -b feature/YourFeature
   ```
3. Make your changes and commit them:
   ```bash
   git commit -m "Add some feature"
   ```
4. Push to the branch:
   ```bash
   git push origin feature/YourFeature
   ```
5. Open a Pull Request.

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

