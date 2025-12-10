# Google Apps Script Project

This project is a Google Apps Script project developed locally using `@google/clasp`.

## Setup

1. **Install `clasp`:**
   ```bash
   npm install -g @google/clasp
   ```

2. **Log in to your Google account:**
   ```bash
   clasp login
   ```

3. **Clone this project:**
   If you haven't already, clone this project to your local machine.

4. **Create a new Google Apps Script project or use an existing one:**
   - To create a new project:
     ```bash
     clasp create --title "My New Project"
     ```
   - To use an existing project, you'll need the Script ID. You can find it in the project's URL. Then run:
     ```bash
     clasp clone <scriptId>
     ```

5. **Push the code to your Google Apps Script project:**
   ```bash
   clasp push
   ```

## Usage

- `clasp pull`: Pulls the code from your Google Apps Script project to your local machine.
- `clasp push`: Pushes the code from your local machine to your Google Apps Script project.
- `clasp open`: Opens the Google Apps Script project in the browser.

For more information on `clasp`, see the [official documentation](https://github.com/google/clasp).

## Project Overview

This Google Apps Script (`Code.js`) is the server-side logic for a web application that generates demo scripts for Google Workspace. 

### Core Functionality

The script has two primary responsibilities:

1.  **Demo User Management:** It creates, manages, and deletes user accounts within a Google Workspace domain (`cymbal.se`). This is handled through the Google Admin SDK Directory API.
2.  **Demo Script Generation:** It leverages a generative AI model (Gemini 1.5 Flash) to create customized, step-by-step demo scripts based on user-provided context.

### Generative AI Integration (Gemini)

*   **`callGeminiAPI(prompt, modelName)`:** This is the central function for interacting with the Gemini model. It takes a prompt and a model name, then sends a request to the Google Generative Language API. The function is configured to receive a JSON object as a response, which is crucial for structured data retrieval.
*   **`constructDemoScriptPrompt_(...)`:** This helper function is a prime example of **prompt engineering**. It builds a detailed prompt for the AI, specifying:
    *   The persona for the AI ("expert demo script writer").
    *   A strict JSON schema for the output, including fields like `summary`, `title`, `steps`, and `presenter_script`.
    *   Critical rules for the output, such as returning *only* a valid JSON object.
    *   Contextual information, including the demo user's name, email, and the customer's needs.
*   **`generateDemoScript(demoContext)`:** This function orchestrates the script generation process. It retrieves the current demo user's information, calls `constructDemoScriptPrompt_` to create the prompt, and then executes `callGeminiAPI`. It includes robust error handling to parse and validate the JSON returned by the AI, ensuring the data is in the expected format.

### Google Workspace Integration (Admin SDK & OAuth)

*   **Authentication:** The script uses a service account and the OAuth2 for Apps Script library to authenticate with Google APIs.
    *   `getOAuthService_` and `getService_` create and manage an OAuth2 service object, handling the token exchange process.
    *   `getAdminAccessToken` retrieves the access token needed to make authorized calls to the Admin SDK.
*   **User Management:**
    *   **`getInitialAppState()`**: When the web app loads, this function checks if a demo user already exists for the person using the app. It does this by making a `GET` request to the Admin SDK's `users` endpoint.
    *   **`createUser(nameInfo)`**: If a demo user doesn't exist, this function is called. It sends a `POST` request to the Admin SDK to provision a new user with a random password.
    *   **`resetPasswordForDemoUser()`**: This allows the user to get a new password for their demo account.
    *   **`deleteDemoAccount(userEmail)`**: This function cleans up by sending a `DELETE` request to the Admin SDK to remove the demo user.

### How It Works in Practice

1.  A user opens the web application in their browser.
2.  The app's front-end (in `Index.html`) calls `getInitialAppState`. The script checks if a demo user exists for them in the `cymbal.se` domain.
3.  If no user exists, the UI will prompt them to create one. The front-end captures the user's name and calls the `createUser` function.
4.  Once a user is provisioned, the user can input the context for a demo (e.g., "Show how a sales team can collaborate on a proposal using Google Drive and Docs").
5.  This context is sent to the `generateDemoScript` function.
6.  The script constructs a detailed prompt and queries the Gemini API.
7.  Gemini returns a structured JSON object containing a complete, step-by-step demo script.
8.  The script parses this JSON and sends it back to the front-end to be displayed to the user.

In essence, this is a sophisticated application that combines the power of generative AI for content creation with the administrative capabilities of the Google Workspace platform, all orchestrated within the serverless environment of Google Apps Script.
