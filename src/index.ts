#!/usr/bin/env node

// Import required dependencies
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";
import "isomorphic-fetch"; // Required for MS Graph client
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import dotenv from "dotenv";

// Load environment variables
dotenv.config();

// Get the directory name of the current module
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Log file for debugging
const logFile = path.join(__dirname, "..", "graph-server.log");

// Helper function for logging (useful for Windows debugging)
function logToFile(message: string): void {
  const timestamp = new Date().toISOString();
  const logMessage = `${timestamp} - ${message}\n`;
  
  try {
    fs.appendFileSync(logFile, logMessage, { encoding: "utf8" });
  } catch (error) {
    // Don't throw errors if logging fails
    console.error("Failed to write to log file:", error);
  }
}

// Initialize MS Graph Client
function initializeGraphClient() {
  try {
    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;
    const scopes = (process.env.SCOPES || "User.Read").split(",");

    if (!tenantId || !clientId || !clientSecret) {
      throw new Error("Missing required environment variables for authentication");
    }

    // Create the credential object
    const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
    
    // Create an authentication provider
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: scopes,
    });

    // Initialize the Graph client
    const graphClient = Client.initWithMiddleware({
      authProvider: authProvider,
    });

    logToFile("MS Graph client initialized successfully");
    return graphClient;
  } catch (error) {
    logToFile(`Error initializing MS Graph client: ${error}`);
    throw error;
  }
}

// Create server instance with error handling for Windows
try {
  logToFile("Initializing MCP server");
  
  const server = new McpServer({
    name: "msgraph",
    version: "1.0.0",
  });

  // Initialize MS Graph client
  const graphClient = initializeGraphClient();

  // Register MS Graph tools
  server.tool(
    "get-profile",
    "Get the current user's profile information",
    {},
    async () => {
      logToFile("Get profile tool called");
      
      try {
        const user = await graphClient.api('/me').get();
        
        const userInfo = `
User Profile Information:
Display Name: ${user.displayName || 'N/A'}
Email: ${user.mail || user.userPrincipalName || 'N/A'}
Job Title: ${user.jobTitle || 'N/A'}
Department: ${user.department || 'N/A'}
Office Location: ${user.officeLocation || 'N/A'}
User ID: ${user.id}
`;

        return {
          content: [
            {
              type: "text",
              text: userInfo,
            },
          ],
        };
      } catch (error) {
        logToFile(`Error fetching user profile: ${error}`);
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve user profile: ${error}`,
            },
          ],
        };
      }
    }
  );

  // Get recent emails
  server.tool(
    "get-emails",
    "Get recent emails from the user's inbox",
    {
      count: z.number().min(1).max(50).default(10).describe("Number of emails to retrieve"),
    },
    async ({ count }) => {
      logToFile(`Get emails tool called, requesting ${count} emails`);
      
      try {
        const messages = await graphClient
          .api('/me/messages')
          .top(count)
          .orderby('receivedDateTime DESC')
          .select('subject,from,receivedDateTime,bodyPreview')
          .get();

        if (!messages.value || messages.value.length === 0) {
          return {
            content: [
              {
                type: "text",
                text: "No emails found in the inbox.",
              },
            ],
          };
        }

        const emailList = messages.value.map((message: any, index: number) => {
          const from = message.from.emailAddress.name || message.from.emailAddress.address;
          const date = new Date(message.receivedDateTime).toLocaleString();
          
          return `Email ${index + 1}:
From: ${from}
Date: ${date}
Subject: ${message.subject}
Preview: ${message.bodyPreview}
---`;
        }).join('\n\n');

        return {
          content: [
            {
              type: "text",
              text: `Recent ${count} emails:\n\n${emailList}`,
            },
          ],
        };
      } catch (error) {
        logToFile(`Error fetching emails: ${error}`);
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve emails: ${error}`,
            },
          ],
        };
      }
    }
  );

  // Get calendar events
  server.tool(
    "get-calendar-events",
    "Get upcoming calendar events",
    {
      days: z.number().min(1).max(30).default(7).describe("Number of days to look ahead"),
    },
    async ({ days }) => {
      logToFile(`Get calendar events tool called, looking ahead ${days} days`);
      
      try {
        const now = new Date();
        const future = new Date();
        future.setDate(future.getDate() + days);
        
        const nowIso = now.toISOString();
        const futureIso = future.toISOString();

        const events = await graphClient
          .api('/me/calendarView')
          .query({
            startDateTime: nowIso,
            endDateTime: futureIso,
          })
          .select('subject,organizer,start,end,location')
          .orderby('start/dateTime')
          .top(50)
          .get();

        if (!events.value || events.value.length === 0) {
          return {
            content: [
              {
                type: "text",
                text: `No calendar events found in the next ${days} days.`,
              },
            ],
          };
        }

        const eventList = events.value.map((event: any, index: number) => {
          const organizer = event.organizer.emailAddress.name || event.organizer.emailAddress.address;
          const startTime = new Date(event.start.dateTime + 'Z').toLocaleString();
          const endTime = new Date(event.end.dateTime + 'Z').toLocaleString();
          const location = event.location?.displayName || 'No location specified';
          
          return `Event ${index + 1}:
Subject: ${event.subject}
Organizer: ${organizer}
Start: ${startTime}
End: ${endTime}
Location: ${location}
---`;
        }).join('\n\n');

        return {
          content: [
            {
              type: "text",
              text: `Calendar events for the next ${days} days:\n\n${eventList}`,
            },
          ],
        };
      } catch (error) {
        logToFile(`Error fetching calendar events: ${error}`);
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve calendar events: ${error}`,
            },
          ],
        };
      }
    }
  );

  // Search for users
  server.tool(
    "search-users",
    "Search for users in the organization",
    {
      query: z.string().min(3).describe("Search query for users (name, email, etc.)"),
      limit: z.number().min(1).max(20).default(5).describe("Maximum number of results to return"),
    },
    async ({ query, limit }) => {
      logToFile(`Search users tool called with query: ${query}, limit: ${limit}`);
      
      try {
        const users = await graphClient
          .api('/users')
          .filter(`startswith(displayName,'${query}') or startswith(mail,'${query}') or startswith(userPrincipalName,'${query}')`)
          .select('displayName,mail,jobTitle,department')
          .top(limit)
          .get();

        if (!users.value || users.value.length === 0) {
          return {
            content: [
              {
                type: "text",
                text: `No users found matching '${query}'.`,
              },
            ],
          };
        }

        const userList = users.value.map((user: any, index: number) => {
          return `User ${index + 1}:
Name: ${user.displayName || 'N/A'}
Email: ${user.mail || user.userPrincipalName || 'N/A'}
Title: ${user.jobTitle || 'N/A'}
Department: ${user.department || 'N/A'}
---`;
        }).join('\n\n');

        return {
          content: [
            {
              type: "text",
              text: `Users matching '${query}':\n\n${userList}`,
            },
          ],
        };
      } catch (error) {
        logToFile(`Error searching users: ${error}`);
        return {
          content: [
            {
              type: "text",
              text: `Failed to search users: ${error}`,
            },
          ],
        };
      }
    }
  );

  // Start the server with Windows-specific error handling
  async function main() {
    try {
      logToFile("Creating stdio transport");
      const transport = new StdioServerTransport();
      
      logToFile("Connecting server to transport");
      await server.connect(transport);
      
      logToFile("MS Graph MCP Server running on stdio");
      console.error("MS Graph MCP Server running on stdio");
    } catch (error) {
      logToFile(`Fatal error initializing server: ${error}`);
      console.error("Fatal error initializing server:", error);
      process.exit(1);
    }
  }

  main().catch((error) => {
    logToFile(`Fatal error in main(): ${error}`);
    console.error("Fatal error in main():", error);
    process.exit(1);
  });

} catch (error) {
  logToFile(`Fatal error during setup: ${error}`);
  console.error("Fatal error during setup:", error);
  process.exit(1);
}
