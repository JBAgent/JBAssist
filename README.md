# JBAssist - Microsoft Graph MCP Server

A Windows-compatible implementation of the Model Context Protocol (MCP) server that connects Claude Desktop to Microsoft Graph API.

## Overview

This project provides a simple MCP server that allows Claude to interact with Microsoft 365 services through Microsoft Graph API. The server offers a variety of tools to access both standard and beta Graph API features.

### Standard API Tools

1. **get-profile**: Get the current user's profile information
2. **get-emails**: Retrieve recent emails from the user's inbox
3. **get-calendar-events**: Get upcoming calendar events for the specified number of days
4. **search-users**: Search for users in the organization

### Beta API Tools (Extended User Features)

1. **get-enhanced-profile**: Get detailed user profile info including skills, interests, schools, and more
2. **get-direct-reports**: Get a user's direct reports with detailed information
3. **get-manager**: Get a user's manager information
4. **get-presence**: Get a user's presence information (availability status)
5. **get-teamwork**: Get a user's Microsoft Teams and teamwork activity
6. **advanced-user-search**: Search for users with advanced filtering by skills, department, job title

## Prerequisites

- [Node.js](https://nodejs.org/) version 18 or higher
- [Claude Desktop](https://claude.ai/download) with MCP support
- Windows 10 or 11
- Microsoft 365 account with appropriate permissions
- Azure AD App Registration with appropriate permissions

## Setup

### 1. Create an Azure AD App Registration

1. Go to the [Azure Portal](https://portal.azure.com) and navigate to Azure Active Directory
2. Select "App registrations" and click "New registration"
3. Enter a name for your application (e.g., "JBAssist")
4. Select "Accounts in this organizational directory only" for Supported account types
5. Click "Register"
6. Note the "Application (client) ID" and "Directory (tenant) ID" values
7. Under "Certificates & secrets", create a new client secret and note its value

### 2. Configure API Permissions

1. In your App Registration, go to "API permissions"
2. Click "Add a permission" and select "Microsoft Graph"
3. Choose "Application permissions" (for daemon/service access) or "Delegated permissions" (for user access)
4. Add the following permissions:
   - User.Read
   - Mail.Read
   - Calendars.Read
   - User.Read.All (for user search)
   - Presence.Read.All (for presence information)
   - TeamMember.Read.All (for Teams information)
5. Click "Add permissions"
6. Click "Grant admin consent for [your organization]"

### 3. Clone and Configure the Repository

1. Clone this repository:
   ```
   git clone https://github.com/JBAgent/JBAssist.git
   cd JBAssist
   ```

2. Create a `.env` file based on the `.env.example` template:
   ```
   cp .env.example .env
   ```

3. Edit the `.env` file with your Azure AD app details:
   ```
   TENANT_ID=your-tenant-id
   CLIENT_ID=your-client-id
   CLIENT_SECRET=your-client-secret
   SCOPES=User.Read,Mail.Read,Calendars.Read,User.Read.All,Presence.Read.All,TeamMember.Read.All
   ```

4. Install dependencies:
   ```
   npm install
   ```

5. Build the project:
   ```
   npm run build
   ```

## Running the Microsoft Graph Server with Claude Desktop

1. First, start Claude Desktop

2. Run the server using one of the following methods:

   - PowerShell:
     ```
     .\run-graph-server.ps1
     ```

   - Command Prompt:
     ```
     run-graph-server.bat
     ```

   - Or directly with Node.js:
     ```
     npm start
     ```
   
3. In Claude Desktop, click the puzzle piece icon (🧩) in the top-right corner to open the MCP server selector

4. Select "Add MCP Server" and enter `msgraph` when prompted for the server name

5. You should see confirmation that the Microsoft Graph server has connected successfully

## Using the Microsoft Graph Tools

Once connected, you can ask Claude to use the MS Graph tools with natural language, for example:

### Standard API Examples

- "Can you show me my profile information?"
- "What are my recent emails?"
- "Show me my calendar events for the next week."
- "Can you search for users with the name John?"

### Beta API Examples

- "Get my enhanced profile with skills and interests."
- "Who are my direct reports?"
- "Who is my manager?"
- "What's my current presence status?"
- "Show me my Teams information."
- "Search for users in the Marketing department who have SQL skills."

## Beta API Features

The Microsoft Graph beta API provides access to features that are still in development. These features are subject to change and may not be stable. The beta API provides access to:

### Enhanced User Profiles

The beta API offers expanded user profile information, including:

- Skills and expertise
- Interests
- Past projects
- Education history
- Responsibilities
- About me information
- Birthday
- Preferred languages
- Alternative email addresses

### Organizational Relationships

- Direct reports hierarchy
- Manager information
- Team memberships
- Team presence and status

### Presence Information

- User availability status
- Activity information
- Status messages
- Teams presence information

### Advanced Search

- Search by skills
- Search by department 
- Search by job title
- Combined search with multiple filters

## Troubleshooting

If you encounter issues:

1. **Authentication Errors**: Double-check your Azure AD app credentials in the `.env` file

2. **Permission Errors**: Ensure your Azure AD app has the appropriate permissions granted

3. **Connection Errors**: Make sure you've clicked the "Allow" button when Claude Desktop prompts for MCP permission

4. **Log Files**: Check the `graph-server.log` file for detailed error information

5. **Path Issues**: Ensure Node.js and npm are in your system's PATH variable

6. **Permission Errors in Windows**: Try running Command Prompt or PowerShell as Administrator

7. **Beta API Errors**: Remember that beta API endpoints may change or behave differently than documented

## Development

To modify the server or add new functionality:

1. Edit the files in the `src` directory
2. Add new tools in `src/index.ts` following the existing patterns
3. Rebuild with `npm run build`
4. Restart the server

### Adding New Beta API Features

When adding new beta API features:

1. Reference the [Microsoft Graph beta API documentation](https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-beta)
2. Use the `graphBetaClient` instance for beta API calls
3. Test thoroughly as beta APIs may change
4. Add appropriate error handling for beta features
5. Document any limitations or known issues

## Additional Microsoft Graph Endpoints

This server provides basic functionality, but Microsoft Graph offers many more endpoints. Here are some examples you could add:

- OneNote access
- SharePoint document access
- Teams message access
- OneDrive file access
- Task management

See the [Microsoft Graph documentation](https://learn.microsoft.com/en-us/graph/api/overview) for more endpoints.

## License

MIT

## Acknowledgements

This project is based on the [Windows MCP Weather example](https://github.com/JBAgent/windows-mcp-weather-example) and adapted to work with Microsoft Graph.
