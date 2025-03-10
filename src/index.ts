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
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const scopes = (process.env.SCOPES || "User.Read").split(",");

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Missing required environment variables for authentication");
  }

  // Initialize MS Graph clients
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

  // NEW BETA API TOOLS

  // Get enhanced user profile using beta API
  server.tool(
    "get-enhanced-profile",
    "Get enhanced user profile information using the beta API",
    {},
    async () => {
      logToFile("Get enhanced profile tool called (beta API)");
      
      try {
        const user = await graphBetaClient.api('/me')
          .select('id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,aboutMe,skills,interests,pastProjects,responsibilities,schools,preferredLanguage,otherMails,businessPhones,mobilePhone,birthday')
          .get();
        
        // Build enhanced profile information
        let profileInfo = `
Enhanced User Profile (Beta API):
Display Name: ${user.displayName || 'N/A'}
Email: ${user.mail || user.userPrincipalName || 'N/A'}
Job Title: ${user.jobTitle || 'N/A'}
Department: ${user.department || 'N/A'}
Office Location: ${user.officeLocation || 'N/A'}
`;

        // Add phone information if available
        if (user.businessPhones && user.businessPhones.length > 0) {
          profileInfo += `Business Phone: ${user.businessPhones.join(', ')}\n`;
        }
        
        if (user.mobilePhone) {
          profileInfo += `Mobile Phone: ${user.mobilePhone}\n`;
        }

        // Add alternative emails if available
        if (user.otherMails && user.otherMails.length > 0) {
          profileInfo += `Alternative Emails: ${user.otherMails.join(', ')}\n`;
        }
        
        // Add preferred language if available
        if (user.preferredLanguage) {
          profileInfo += `Preferred Language: ${user.preferredLanguage}\n`;
        }
        
        // Add birthday if available
        if (user.birthday) {
          const birthday = new Date(user.birthday).toLocaleDateString();
          profileInfo += `Birthday: ${birthday}\n`;
        }

        // Add about me if available
        if (user.aboutMe) {
          profileInfo += `\nAbout Me:\n${user.aboutMe}\n`;
        }
        
        // Add skills if available
        if (user.skills && user.skills.length > 0) {
          profileInfo += `\nSkills:\n${user.skills.join('\n- ')}\n`;
        }
        
        // Add interests if available
        if (user.interests && user.interests.length > 0) {
          profileInfo += `\nInterests:\n${user.interests.join('\n- ')}\n`;
        }
        
        // Add past projects if available
        if (user.pastProjects && user.pastProjects.length > 0) {
          profileInfo += `\nPast Projects:\n${user.pastProjects.join('\n- ')}\n`;
        }
        
        // Add responsibilities if available
        if (user.responsibilities && user.responsibilities.length > 0) {
          profileInfo += `\nResponsibilities:\n${user.responsibilities.join('\n- ')}\n`;
        }
        
        // Add schools if available
        if (user.schools && user.schools.length > 0) {
          profileInfo += `\nEducation:\n`;
          user.schools.forEach((school: any) => {
            profileInfo += `- ${school.description || 'Unnamed School'}`;
            if (school.graduationYear) {
              profileInfo += ` (Graduation: ${school.graduationYear})`;
            }
            profileInfo += '\n';
          });
        }

        return {
          content: [
            {
              type: "text",
              text: profileInfo,
            },
          ],
        };
      } catch (error) {
        logToFile(`Error fetching enhanced user profile: ${error}`);
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve enhanced user profile: ${error}`,
            },
          ],
        };
      }
    }
  );

  // Get user's direct reports using beta API
  server.tool(
    "get-direct-reports",
    "Get user's direct reports using beta API",
    {
      userId: z.string().optional().describe("User ID to get direct reports for (default: current user)"),
    },
    async ({ userId }) => {
      logToFile(`Get direct reports tool called for user: ${userId || 'current user'} (beta API)`);
      
      try {
        const endpoint = userId ? `/users/${userId}/directReports` : '/me/directReports';
        
        const directReports = await graphBetaClient.api(endpoint)
          .select('id,displayName,mail,jobTitle,department,officeLocation')
          .get();

        if (!directReports.value || directReports.value.length === 0) {
          return {
            content: [
              {
                type: "text",
                text: "No direct reports found.",
              },
            ],
          };
        }

        const reportsList = directReports.value.map((user: any, index: number) => {
          return `Direct Report ${index + 1}:
Name: ${user.displayName || 'N/A'}
Email: ${user.mail || user.userPrincipalName || 'N/A'}
Title: ${user.jobTitle || 'N/A'}
Department: ${user.department || 'N/A'}
Office: ${user.officeLocation || 'N/A'}
---`;
        }).join('\n\n');

        return {
          content: [
            {
              type: "text",
              text: `Direct Reports:\n\n${reportsList}`,
            },
          ],
        };
      } catch (error) {
        logToFile(`Error fetching direct reports: ${error}`);
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve direct reports: ${error}`,
            },
          ],
        };
      }
    }
  );

  // Get user's manager using beta API
  server.tool(
    "get-manager",
    "Get user's manager information using beta API",
    {
      userId: z.string().optional().describe("User ID to get manager for (default: current user)"),
    },
    async ({ userId }) => {
      logToFile(`Get manager tool called for user: ${userId || 'current user'} (beta API)`);
      
      try {
        const endpoint = userId ? `/users/${userId}/manager` : '/me/manager';
        
        const manager = await graphBetaClient.api(endpoint)
          .select('id,displayName,mail,jobTitle,department,officeLocation,businessPhones,mobilePhone')
          .get();

        if (!manager) {
          return {
            content: [
              {
                type: "text",
                text: "No manager information found.",
              },
            ],
          };
        }

        const managerInfo = `
Manager Information:
Name: ${manager.displayName || 'N/A'}
Email: ${manager.mail || manager.userPrincipalName || 'N/A'}
Title: ${manager.jobTitle || 'N/A'}
Department: ${manager.department || 'N/A'}
Office: ${manager.officeLocation || 'N/A'}
${manager.businessPhones && manager.businessPhones.length > 0 ? `Business Phone: ${manager.businessPhones[0]}\n` : ''}${manager.mobilePhone ? `Mobile Phone: ${manager.mobilePhone}\n` : ''}`;

        return {
          content: [
            {
              type: "text",
              text: managerInfo,
            },
          ],
        };
      } catch (error) {
        logToFile(`Error fetching manager information: ${error}`);
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve manager information: ${error}`,
            },
          ],
        };
      }
    }
  );

  // Get user's presence information using beta API
  server.tool(
    "get-presence",
    "Get user's presence information using beta API",
    {
      userId: z.string().optional().describe("User ID to get presence for (default: current user)"),
    },
    async ({ userId }) => {
      logToFile(`Get presence tool called for user: ${userId || 'current user'} (beta API)`);
      
      try {
        let presence;
        
        if (userId) {
          // For other users
          presence = await graphBetaClient.api(`/users/${userId}/presence`)
            .get();
        } else {
          // For current user
          presence = await graphBetaClient.api('/me/presence')
            .get();
        }

        if (!presence) {
          return {
            content: [
              {
                type: "text",
                text: "No presence information found.",
              },
            ],
          };
        }

        const presenceInfo = `
Presence Information:
Availability: ${presence.availability || 'N/A'}
Activity: ${presence.activity || 'N/A'}
Status: ${presence.status || 'N/A'}
${presence.statusMessage ? `Status Message: ${presence.statusMessage}\n` : ''}
Last Modified: ${presence.lastModifiedDateTime ? new Date(presence.lastModifiedDateTime).toLocaleString() : 'N/A'}
`;

        return {
          content: [
            {
              type: "text",
              text: presenceInfo,
            },
          ],
        };
      } catch (error) {
        logToFile(`Error fetching presence information: ${error}`);
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve presence information: ${error}`,
            },
          ],
        };
      }
    }
  );

  // Get user's teamwork information using beta API
  server.tool(
    "get-teamwork",
    "Get user's Microsoft Teams and teamwork activity using beta API",
    {},
    async () => {
      logToFile("Get teamwork information tool called (beta API)");
      
      try {
        // Get joined teams
        const joinedTeams = await graphBetaClient.api('/me/joinedTeams')
          .select('id,displayName,description,visibility')
          .get();

        // Get team presence
        const presence = await graphBetaClient.api('/me/presence')
          .get();

        let teamworkInfo = `
Microsoft Teams Information:
`;

        // Add presence information
        if (presence) {
          teamworkInfo += `
Presence:
Status: ${presence.availability || 'N/A'}
Activity: ${presence.activity || 'N/A'}
`;
        }

        // Add teams information
        if (joinedTeams && joinedTeams.value && joinedTeams.value.length > 0) {
          teamworkInfo += `\nJoined Teams (${joinedTeams.value.length}):\n`;
          
          joinedTeams.value.forEach((team: any, index: number) => {
            teamworkInfo += `
Team ${index + 1}:
Name: ${team.displayName || 'N/A'}
Description: ${team.description || 'N/A'}
Visibility: ${team.visibility || 'N/A'}
ID: ${team.id}
`;
          });
        } else {
          teamworkInfo += "\nNo teams found for this user.";
        }

        return {
          content: [
            {
              type: "text",
              text: teamworkInfo,
            },
          ],
        };
      } catch (error) {
        logToFile(`Error fetching teamwork information: ${error}`);
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve teamwork information: ${error}`,
            },
          ],
        };
      }
    }
  );

  // Advanced user search with beta API
  server.tool(
    "advanced-user-search",
    "Advanced search for users using beta API with more filtering options",
    {
      query: z.string().min(3).describe("Search query for users"),
      department: z.string().optional().describe("Filter by department"),
      jobTitle: z.string().optional().describe("Filter by job title"),
      skills: z.string().optional().describe("Filter by skills (comma-separated)"),
      limit: z.number().min(1).max(20).default(5).describe("Maximum number of results"),
    },
    async ({ query, department, jobTitle, skills, limit }) => {
      logToFile(`Advanced user search called with query: ${query}, department: ${department}, jobTitle: ${jobTitle}, skills: ${skills} (beta API)`);
      
      try {
        // Build filter string
        let filter = `startswith(displayName,'${query}') or startswith(mail,'${query}') or startswith(userPrincipalName,'${query}')`;
        
        if (department) {
          filter += ` and department eq '${department}'`;
        }
        
        if (jobTitle) {
          filter += ` and jobTitle eq '${jobTitle}'`;
        }
        
        const users = await graphBetaClient.api('/users')
          .filter(filter)
          .select('id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,skills,businessPhones,mobilePhone')
          .top(limit)
          .get();

        if (!users.value || users.value.length === 0) {
          return {
            content: [
              {
                type: "text",
                text: `No users found matching the criteria.`,
              },
            ],
          };
        }

        // If skills filter is provided, do client-side filtering
        let filteredUsers = users.value;
        if (skills) {
          const skillsArray = skills.split(',').map(s => s.trim().toLowerCase());
          filteredUsers = users.value.filter((user: any) => {
            if (!user.skills) return false;
            
            // Check if user has any of the requested skills
            return user.skills.some((skill: string) => 
              skillsArray.includes(skill.toLowerCase())
            );
          });
        }

        if (filteredUsers.length === 0) {
          return {
            content: [
              {
                type: "text",
                text: `No users found with the specified skills.`,
              },
            ],
          };
        }

        const userList = filteredUsers.map((user: any, index: number) => {
          let userInfo = `User ${index + 1}:
Name: ${user.displayName || 'N/A'}
Email: ${user.mail || user.userPrincipalName || 'N/A'}
Title: ${user.jobTitle || 'N/A'}
Department: ${user.department || 'N/A'}
Office: ${user.officeLocation || 'N/A'}
`;

          // Add phone information if available
          if (user.businessPhones && user.businessPhones.length > 0) {
            userInfo += `Business Phone: ${user.businessPhones.join(', ')}\n`;
          }
          
          if (user.mobilePhone) {
            userInfo += `Mobile Phone: ${user.mobilePhone}\n`;
          }

          // Add skills if available
          if (user.skills && user.skills.length > 0) {
            userInfo += `Skills: ${user.skills.join(', ')}\n`;
          }

          userInfo += '---';
          return userInfo;
        }).join('\n\n');

        return {
          content: [
            {
              type: "text",
              text: `Advanced User Search Results:\n\n${userList}`,
            },
          ],
        };
      } catch (error) {
        logToFile(`Error in advanced user search: ${error}`);
        return {
          content: [
            {
              type: "text",
              text: `Failed to perform advanced user search: ${error}`,
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
