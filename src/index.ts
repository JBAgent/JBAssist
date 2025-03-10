#!/usr/bin/env node

// Import required dependencies
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { fileURLToPath } from "url";
import path from "path";
import "isomorphic-fetch"; // Required for MS Graph client
import dotenv from "dotenv";
import { initializeGraphClient, initializeGraphBetaClient, logToFile } from "./auth.js";

// Load environment variables
dotenv.config();

// Get the directory name of the current module
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Create server instance with error handling for Windows
try {
  logToFile("Initializing MCP server");
  
  const server = new McpServer({
    name: "msgraph",
    version: "1.0.0",
  });

  // Get environment variables
  const tenantId = process.env.TENANT_ID || "demo-tenant-id";
  const clientId = process.env.CLIENT_ID || "demo-client-id";
  const clientSecret = process.env.CLIENT_SECRET || "demo-client-secret";
  const scopes = (process.env.SCOPES || "User.Read").split(",");

  // Check if real auth is configured or we're using placeholders
  const isUsingRealAuth = process.env.TENANT_ID && process.env.CLIENT_ID && process.env.CLIENT_SECRET;
  
  if (!isUsingRealAuth) {
    logToFile("WARNING: Using demo credentials - Graph API functionality will be limited");
    console.error("WARNING: Using demo credentials - Graph API functionality will be limited");
  }

  // Initialize MS Graph clients with real or demo credentials
  const graphClient = initializeGraphClient(tenantId, clientId, clientSecret, scopes);
  const graphBetaClient = initializeGraphBetaClient(tenantId, clientId, clientSecret, scopes);

  // Register MS Graph tools
  server.tool(
    "get-profile",
    "Get the current user's profile information",
    {},
    async () => {
      logToFile("Get profile tool called");
      
      try {
        if (!isUsingRealAuth) {
          return {
            content: [
              {
                type: "text",
                text: "Authentication configuration is required to use this feature. Please set up TENANT_ID, CLIENT_ID, and CLIENT_SECRET environment variables.",
              },
            ],
          };
        }
        
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
        if (!isUsingRealAuth) {
          return {
            content: [
              {
                type: "text",
                text: "Authentication configuration is required to use this feature. Please set up TENANT_ID, CLIENT_ID, and CLIENT_SECRET environment variables.",
              },
            ],
          };
        }

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

  // Add check-auth tool to help with setup
  server.tool(
    "check-auth",
    "Check authentication status and get setup help",
    {},
    async () => {
      return {
        content: [
          {
            type: "text",
            text: `
Authentication Status: ${isUsingRealAuth ? "✅ Configured" : "❌ Not Configured"}

${!isUsingRealAuth ? `To use JBAssist with Microsoft Graph, you need to set up authentication:

1. Create a .env file in the root directory with:
   TENANT_ID=your-tenant-id
   CLIENT_ID=your-client-id
   CLIENT_SECRET=your-client-secret
   SCOPES=User.Read,Mail.Read,Calendars.Read,User.Read.All

2. Get these values from Azure Portal:
   - Register an app in Azure Active Directory
   - Note the Application (client) ID and Directory (tenant) ID
   - Create a client secret
   - Grant appropriate permissions

See the README.md for detailed instructions.` : "Authentication is properly configured. All Graph API features should be available."}
`,
          },
        ],
      };
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
