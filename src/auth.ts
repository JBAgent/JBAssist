// Custom authentication solution for MS Graph

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

// Get the directory name of the current module
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Log file for debugging
const logFile = path.join(__dirname, "..", "graph-server.log");

// Helper function for logging (useful for Windows debugging)
export function logToFile(message: string): void {
  const timestamp = new Date().toISOString();
  const logMessage = `${timestamp} - ${message}\n`;
  
  try {
    fs.appendFileSync(logFile, logMessage, { encoding: "utf8" });
  } catch (error) {
    // Don't throw errors if logging fails
    console.error("Failed to write to log file:", error);
  }
}

// Custom implementation of the authentication provider
class CustomAuthProvider {
  private credential: ClientSecretCredential;
  private scopes: string[];
  
  constructor(tenantId: string, clientId: string, clientSecret: string, scopes: string[]) {
    this.credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
    this.scopes = scopes;
  }
  
  // This is the method the Graph client will call to get the access token
  async getAccessToken(): Promise<string> {
    try {
      const response = await this.credential.getToken(this.scopes);
      return response.token;
    } catch (error) {
      logToFile(`Error getting access token: ${error}`);
      throw error;
    }
  }
}

// Initialize standard MS Graph Client
export function initializeGraphClient(tenantId: string, clientId: string, clientSecret: string, scopes: string[]) {
  try {
    if (!tenantId || !clientId || !clientSecret) {
      throw new Error("Missing required environment variables for authentication");
    }

    // Create a custom auth provider that doesn't rely on TokenCredentialAuthenticationProvider
    const authProvider = new CustomAuthProvider(tenantId, clientId, clientSecret, scopes);

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

// Initialize MS Graph Beta Client
export function initializeGraphBetaClient(tenantId: string, clientId: string, clientSecret: string, scopes: string[]) {
  try {
    if (!tenantId || !clientId || !clientSecret) {
      throw new Error("Missing required environment variables for authentication");
    }

    // Create a custom auth provider that doesn't rely on TokenCredentialAuthenticationProvider
    const authProvider = new CustomAuthProvider(tenantId, clientId, clientSecret, scopes);

    // Initialize the Graph client with beta endpoint
    const graphBetaClient = Client.initWithMiddleware({
      authProvider: authProvider,
      baseUrl: "https://graph.microsoft.com/beta",
    });

    logToFile("MS Graph Beta client initialized successfully");
    return graphBetaClient;
  } catch (error) {
    logToFile(`Error initializing MS Graph Beta client: ${error}`);
    throw error;
  }
}
