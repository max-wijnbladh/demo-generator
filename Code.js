/**
 * Code.gs - CONSOLIDATED Server-Side Logic for the Demo Script Generator
 * This version uses a robust JSON-based architecture with a two-step UI flow.
 *
 * REQUIRED SCRIPT PROPERTIES:
 * - GEMINI_API_KEY: API key for Google Gemini.
 * - PRIVATE_KEY: Private key for the Service Account.
 * - CLIENT_EMAIL: Email address of the Service Account.
 * - IMPERSONATED_USER: Email address of the user to impersonate (e.g., admin@domain.com).
 *
 * OPTIONAL SCRIPT PROPERTIES (defaults used if missing):
 * - PRIMARY_DOMAIN: The domain for creating demo users (default: cymbal.se).
 * - GEMINI_MODEL: The Gemini model to use (default: models/gemini-1.5-flash-latest).
 * - LOGGING_SHEET_ID: ID of the Google Sheet for logging (default: hardcoded fallback).
 */

// =====================================================================================
// Configuration Helpers
// =====================================================================================

function getPrimaryDomain() {
  return PropertiesService.getScriptProperties().getProperty('PRIMARY_DOMAIN') || 'cymbal.se';
}

function getGeminiModel() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'models/gemini-1.5-flash-latest';
}

function getLoggingSheetId() {
  return PropertiesService.getScriptProperties().getProperty('LOGGING_SHEET_ID') || '154d_7tRcMxkDrfPtmVFApY6W42a-Vv6RrnuzMUMxIHI';
}

/**
 * Serves the HTML for the web app.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Google Workspace Demo Generator')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}


// =====================================================================================
// Client-Callable Functions
// =====================================================================================

/**
 * Checks if a demo user already exists in the directory for the current app user.
 * If so, it returns their data. Otherwise, it returns a failure status.
 * This is called on initial app load.
 */
function getInitialAppState() {
  const requestingUserEmail = Session.getActiveUser().getEmail();
  if (!requestingUserEmail) {
    // Cannot identify the user running the script.
    return { success: false, error: "Could not identify the requesting user." };
  }

  // Construct the potential demo email address based on the app user's email
  const requestingUserPrefix = requestingUserEmail.split('@')[0].toLowerCase().replace(/[^a-z0-9]/g, '');
  const primaryEmail = `${requestingUserPrefix}@${getPrimaryDomain()}`;

  const token = getAdminAccessToken();
  if (!token) {
     return { success: false, error: "Authentication failed. Could not get admin token." };
  }

  try {
    const checkUserUrl = `https://admin.googleapis.com/admin/directory/v1/users/${primaryEmail}`;
    const checkOptions = {
      method: "GET",
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    };
    const checkResponse = UrlFetchApp.fetch(checkUserUrl, checkOptions);

    if (checkResponse.getResponseCode() === 200) {
      // User EXISTS. Load their data.
      console.log(`Initial check: User ${primaryEmail} already exists. Loading their data.`);
      const existingUser = JSON.parse(checkResponse.getContentText());

      const provisionResult = {
        email: primaryEmail,
        password: null, // IMPORTANT: Never send the password back on initial load for security.
        firstName: existingUser.name.givenName,
        lastName: existingUser.name.familyName
      };

      // Save the found user's data to properties so other functions can use it
      saveDemoState(provisionResult, null);

      // Retrieve any existing script for this user
      const savedState = loadDemoState();

      return {
        success: true,
        provisionResult: provisionResult,
        demoScript: savedState ? savedState.demoScript : null
      };
    } else {
      // User does NOT exist. Tell the UI to show the creation form.
      console.log(`Initial check: User ${primaryEmail} does not exist.`);
      clearDemoState(); // Clear any potentially old/stale data
      return { success: false };
    }
  } catch (e) {
    console.error(`Exception during initial app state check for ${primaryEmail}: ${e.message}`);
    return { success: false, error: "A network exception occurred during the initial check: " + e.message };
  }
}


/**
 * Creates a new user or retrieves an existing one.
 * @param {object} nameInfo An object with {firstName, lastName}.
 * @returns {object} The result of the provisioning, including success status, credentials, and error message if any.
 */
function createUser(nameInfo) {
  const {
    firstName,
    lastName
  } = nameInfo;
  const requestingUserEmail = Session.getActiveUser().getEmail();

  if (!requestingUserEmail) {
    return {
      success: false,
      error: "Could not identify the requesting user."
    };
  }

  const provisionResult = createDemoUserInRoot(requestingUserEmail, firstName, lastName);

  if (provisionResult.success) {
    // Clear any old script data and save the new user data
    clearDemoState();
    saveDemoState(provisionResult, null);
  }

  return provisionResult;
}

/**
 * Resets the password for the currently saved demo user.
 * @returns {object} An object with the new password or an error.
 */
function resetPasswordForDemoUser() {
  const savedState = loadDemoState();
  if (!savedState || !savedState.provisionResult || !savedState.provisionResult.email) {
    return {
      success: false,
      error: "No active demo user found in the session to reset the password for."
    };
  }

  const token = getAdminAccessToken();
  const primaryEmail = savedState.provisionResult.email;

  if (!token) {
    return { success: false, error: "Authentication failed. Could not get token to reset password." };
  }

  console.log(`Attempting to reset password for ${primaryEmail}.`);
  const newPassword = generateRandomPassword_(14);

  const updateUserPayload = {
    password: newPassword,
    changePasswordAtNextLogin: false
  };

  const updateUserUrl = `https://admin.googleapis.com/admin/directory/v1/users/${primaryEmail}`;
  const updateOptions = {
    method: "PUT",
    contentType: "application/json",
    headers: { "Authorization": "Bearer " + token },
    payload: JSON.stringify(updateUserPayload),
    muteHttpExceptions: true
  };

  try {
    const updateResponse = UrlFetchApp.fetch(updateUserUrl, updateOptions);
    const updateResponseCode = updateResponse.getResponseCode();

    if (updateResponseCode === 200) {
      console.log(`Successfully updated password for ${primaryEmail}.`);
      // Update the saved state with the new password
      savedState.provisionResult.password = newPassword;
      saveDemoState(savedState.provisionResult, savedState.demoScript);
      return { success: true, password: newPassword };
    } else {
      const errorBody = updateResponse.getContentText();
      console.error(`Failed to update password for user ${primaryEmail}. Code: ${updateResponseCode}, Body: ${errorBody}`);
      return { success: false, error: `Failed to reset password. Admin API responded with code ${updateResponseCode}.` };
    }
  } catch (e) {
      console.error(`Exception during password reset for ${primaryEmail}: ${e.message}`);
      return { success: false, error: "A network exception occurred during password reset." };
  }
}


/**
 * Generates a demo script for the currently saved user.
 * @param {string} demoContext The user-provided context for the demo.
 * @returns {object} The result of the script generation, including the script object.
 */
function generateDemoScript(demoContext) {
  const savedState = loadDemoState();
  if (!savedState || !savedState.provisionResult) {
    return {
      success: false,
      error: "No demo user found. Please create a demo account first."
    };
  }

  const {
    email: newUserEmail,
    firstName: newUserFirstName,
    lastName: newUserLastName
  } = savedState.provisionResult;
  const demoScriptPrompt = constructDemoScriptPrompt_(demoContext, newUserFirstName, newUserLastName, newUserEmail);
  const scriptGenerationResult = callGeminiAPI(demoScriptPrompt, getGeminiModel());

  if (!scriptGenerationResult.success || !scriptGenerationResult.text) {
    console.error('[Code.gs] Demo script generation failed:', scriptGenerationResult.error);
    // Log the prompt and failed output
    logPromptAndOutput(demoScriptPrompt, "Script generation failed: " + scriptGenerationResult.error);
    return {
      success: false,
      error: `Demo script generation failed: ${scriptGenerationResult.error}`
    };
  }

  try {
    const scriptObject = JSON.parse(scriptGenerationResult.text);

    // Validate the script object
    if (!scriptObject.title || !scriptObject.steps || !Array.isArray(scriptObject.steps)) {
      throw new Error("AI response was missing required 'title' or 'steps' fields.");
    }

    // Save the new script along with the existing user data
    saveDemoState(savedState.provisionResult, scriptObject);
    // Log the prompt and successful output
    logPromptAndOutput(demoScriptPrompt, JSON.stringify(scriptObject));
    return {
      success: true,
      demoScript: scriptObject
    };
  } catch (e) {
    console.error(`[Code.gs] Failed to parse or validate demo script JSON. Error: ${e.message}. Raw Text: ${scriptGenerationResult.text}`);
    // Log the prompt and the raw, malformed output
    logPromptAndOutput(demoScriptPrompt, "Malformed AI response: " + scriptGenerationResult.text);
    return {
      success: false,
      error: `The AI returned malformed data that could not be read. Please try again.`
    };
  }
}

/**
 * Deletes the demo account user and clears the state.
 * @param {string} userEmail The email of the user to delete.
 * @returns {object} The result of the deletion attempt.
 */
function deleteDemoAccount(userEmail) {
  if (!userEmail) {
    return { success: false, error: "No user email provided for deletion." };
  }

  const token = getAdminAccessToken();
  if (!token) {
    return { success: false, error: "Authentication failed. Could not get token to delete user." };
  }

  const url = `https://admin.googleapis.com/admin/directory/v1/users/${userEmail}`;
  const options = {
    method: "DELETE",
    headers: { "Authorization": "Bearer " + token },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 204 || responseCode === 404) {
      console.log(`Successfully processed deletion for user ${userEmail} (Status: ${responseCode}).`);
      clearDemoState();
      return { success: true };
    } else {
      const errorBody = response.getContentText();
      console.error(`Failed to delete user ${userEmail}. Code: ${responseCode}, Body: ${errorBody}`);
      return { success: false, error: `Failed to delete user. Admin API responded with code ${responseCode}.` };
    }
  } catch (e) {
    console.error(`Exception during user deletion for ${userEmail}: ${e.message}`);
    return { success: false, error: "A network exception occurred during deletion." };
  }
}

// =====================================================================================
// State Management Functions
// =====================================================================================

function getNamespacedKey_(key) {
  const userEmail = Session.getActiveUser().getEmail();
  // Fallback to a generic key if email isn't available (shouldn't happen in authenticated web app)
  if (!userEmail) {
    console.warn("No active user email found. Using generic key.");
    return key;
  }
  return `${key}_${userEmail}`;
}

function saveDemoState(provisionResult, demoScript) {
  try {
    const userProperties = PropertiesService.getUserProperties();

    if (provisionResult) {
      userProperties.setProperty(getNamespacedKey_('savedProvisionResult'), JSON.stringify(provisionResult));
    }
    if (demoScript) {
      userProperties.setProperty(getNamespacedKey_('savedDemoScript'), JSON.stringify(demoScript));
    }
  } catch(e) {
    console.error("Error saving state to UserProperties: " + e.toString());
  }
}

/**
 * Loads the currently saved state from UserProperties.
 * This is used by functions that run after the initial load.
 */
function loadDemoState() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const savedProvisionResult = userProperties.getProperty(getNamespacedKey_('savedProvisionResult'));

    if (savedProvisionResult) {
      const savedDemoScript = userProperties.getProperty(getNamespacedKey_('savedDemoScript'));
      const result = {
        provisionResult: JSON.parse(savedProvisionResult),
        demoScript: savedDemoScript ? JSON.parse(savedDemoScript) : null
      };
      return result;
    }
  } catch(e) {
     console.error("Error loading or parsing state from UserProperties: " + e.toString());
     // Clear potentially corrupted properties
     clearDemoState();
  }
  return null;
}

/**
 * Clears all demo-related UserProperties.
 */
function clearDemoState() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty(getNamespacedKey_('savedProvisionResult'));
  userProperties.deleteProperty(getNamespacedKey_('savedDemoScript'));
  console.log('Internal call: Demo state cleared.');
}

/**
 * Clears only the saved demo script from UserProperties.
 */
function clearDemoScriptOnly() {
  PropertiesService.getUserProperties().deleteProperty(getNamespacedKey_('savedDemoScript'));
  console.log('Demo script data cleared.');
}

/**
 * Logs the prompt and output to a specified Google Sheet.
 * @param {string} prompt The user's prompt.
 * @param {string} output The generated output.
 */
function logPromptAndOutput(prompt, output) {
  try {
    // Attempt to open the spreadsheet by ID
    const spreadsheet = SpreadsheetApp.openById(getLoggingSheetId());
    // Get the first sheet
    const sheet = spreadsheet.getSheets()[0];
    // Append the log entry
    sheet.appendRow([new Date(), prompt, output]);
  } catch (e) {
    // Fail silently but log to console
    console.error("Error logging prompt and output to Google Sheet:", e.message);
  }
}

// =====================================================================================
// Helper & Core Logic Functions
// =====================================================================================
function constructDemoScriptPrompt_(demoContext, demoUserFirstName, demoUserLastName, demoUserEmail) {
  return `
    You are an expert demo script writer for Google Workspace. Your task is to generate a step-by-step demo script based on user context.

    Return ONLY a valid JSON object. Do not include any other text, Markdown formatting, or code fences like \`\`\`json.

    The JSON Schema must be:
    {
      "summary": "A concise, one-paragraph summary of the demo flow, scenario, and goals based on the user context.",
      "title": "A creative and professional title for the demo.",
      "introduction": "A brief, welcoming presenter talking script to kick off the demo. Use placeholders like \${demoUserFirstName}.",
      "prerequisites": ["A clear, concise instruction for a prerequisite step. For example, 'Ensure an email with the subject 'Q3 Project Proposal' is in the user's inbox.'"],
      "steps": [
        {
          "step_title": "Concise title for this step (e.g., 'Step 1: Drafting a Response in Gmail')",
          "action": "A direct, imperative instruction for the presenter to perform (e.g., 'Open a new Incognito browser window.').",
          "ui_interaction": "Detailed, step-by-step instructions on how to perform the actions in the UI (e.g., 'Click the 'Compose' button. In the 'To' field, enter the recipient's email. Use Gemini to help you draft the email body.').",
          "presenter_script": "The presenter's talking points for this step. This should explain the 'why' behind the actions."
        }
      ]
    }

    CRITICAL RULES:
    - The entire output MUST be a single, valid JSON object.
    - Do NOT use markdown like asterisks or backticks.
    - If there are no specific prerequisites, provide one generic one like "Practice the script before the live demo."
    - The "prerequisites" field MUST be a top-level key and its value must be an array of strings.

    Demo User & Context:
    * First Name: ${demoUserFirstName}
    * Last Name: ${demoUserLastName}
    * Email: ${demoUserEmail}
    * Customer Needs: "${demoContext}"

    Generate the JSON object now.
  `;
}

function createDemoUserInRoot(requestingUserEmail, firstName, lastName) {
  const token = getAdminAccessToken();
  if (!token) return { success: false, error: "Could not get authentication token to manage users." };

  const requestingUserPrefix = requestingUserEmail.split('@')[0].toLowerCase().replace(/[^a-z0-9]/g, '');
  const primaryEmail = `${requestingUserPrefix}@${getPrimaryDomain()}`;

  try {
    const checkUserUrl = `https://admin.googleapis.com/admin/directory/v1/users/${primaryEmail}`;
    const checkOptions = {
      method: "GET",
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    };
    const checkResponse = UrlFetchApp.fetch(checkUserUrl, checkOptions);

    if (checkResponse.getResponseCode() === 200) {
      console.log(`User ${primaryEmail} already exists. Returning their information.`);
      const existingUser = JSON.parse(checkResponse.getContentText());
      return {
        success: true,
        email: primaryEmail,
        password: null, // Explicitly return no password
        firstName: existingUser.name.givenName,
        lastName: existingUser.name.familyName
      };
    }

    if (checkResponse.getResponseCode() === 404) {
      console.log("User does not exist, creating:", primaryEmail);
      const password = generateRandomPassword_(14);
      const userPayload = {
        primaryEmail,
        password,
        name: { givenName: firstName, familyName: lastName },
        orgUnitPath: "/",
        changePasswordAtNextLogin: false
      };
      const createUserUrl = "https://admin.googleapis.com/admin/directory/v1/users";
      const createOptions = {
        method: "POST",
        contentType: "application/json",
        headers: { "Authorization": "Bearer " + token },
        payload: JSON.stringify(userPayload),
        muteHttpExceptions: true
      };
      const createResponse = UrlFetchApp.fetch(createUserUrl, createOptions);
      if (createResponse.getResponseCode() === 200 || createResponse.getResponseCode() === 201) {
        return { success: true, email: primaryEmail, password: password, firstName, lastName };
      } else {
        return { success: false, error: `Failed to create user. Admin API responded with code ${createResponse.getResponseCode()}.` };
      }
    } else {
      return { success: false, error: `An unexpected error occurred. Code: ${checkResponse.getResponseCode()}` };
    }
  } catch (e) {
    console.error("Exception during user provisioning:", e.message);
    return { success: false, error: "A network exception occurred: " + e.message };
  }
}

function callGeminiAPI(prompt, modelName) {
  let apiKey;
  try {
    apiKey = getScriptProperty_('GEMINI_API_KEY');
  } catch (e) {
    return { success: false, error: e.message };
  }
  const url = `https://generativelanguage.googleapis.com/v1beta/${modelName}:generateContent?key=${apiKey}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: "application/json" }
  };
  const options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      const aiText = jsonResponse.candidates?.[0]?.content?.parts?.[0]?.text;
      if (aiText) return { success: true, text: aiText };
      return { success: false, error: "AI content generation failed: No valid content received." };
    } else {
      return { success: false, error: `AI API request failed with status ${responseCode}. Response: ${responseBody.substring(0, 500)}` };
    }
  } catch (e) {
    return { success: false, error: "Exception during AI API call: " + e.message };
  }
}

function generateRandomPassword_(length = 14) {
  const upper = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  const lower = "abcdefghijklmnopqrstuvwxyz";
  const digits = "0123456789";
  const symbols = "!@#$%^&*()";
  const all = upper + lower + digits + symbols;
  let password = "";
  password += upper[Math.floor(Math.random() * upper.length)];
  password += lower[Math.floor(Math.random() * lower.length)];
  password += digits[Math.floor(Math.random() * digits.length)];
  password += symbols[Math.floor(Math.random() * symbols.length)];
  for (let i = 4; i < length; i++) {
    password += all[Math.floor(Math.random() * all.length)];
  }
  return password.split('').sort(() => 0.5 - Math.random()).join('');
}

// =====================================================================================
// Authentication Functions
// =====================================================================================

/**
 * Creates and configures the OAuth2 service.
 * Retrieves credentials securely from Script Properties.
 * @returns {OAuth2.Service} The configured OAuth2 service.
 */
function getOAuthService_() {
  // Retrieve credentials from Script Properties
  const privateKey = getScriptProperty_('PRIVATE_KEY');
  const clientEmail = getScriptProperty_('CLIENT_EMAIL');
  const impersonatedUser = getScriptProperty_('IMPERSONATED_USER');

  return OAuth2.createService('DemoGeneratorAdminService')
    .setTokenUrl('https://accounts.google.com/o/oauth2/token')
    .setPrivateKey(privateKey)
    .setIssuer(clientEmail)
    .setSubject(impersonatedUser)
    .setPropertyStore(PropertiesService.getUserProperties())
    .setCache(CacheService.getUserCache())
    .setLock(LockService.getUserLock())
    .setScope("https://www.googleapis.com/auth/admin.directory.user");
}

/**
 * Gets the OAuth2 service and checks if it has access.
 * @returns {OAuth2.Service|null} The service object if access is granted, null otherwise.
 */
function getService_() {
  // No arguments needed, credentials are internal
  var service = getOAuthService_();

  if (service && service.hasAccess()) {
    return service;
  } else {
    Logger.log('getService_(): Access denied. Last Error: ' + (service ? service.getLastError() : "Service object was null."));
    return null;
  }
}

/**
 * Retrieves the access token for the admin user.
 * @returns {string} The access token.
 * @throws {Error} If authentication fails.
 */
function getAdminAccessToken() {
  var service = getService_();
  if (service) {
    return service.getAccessToken();
  } else {
    // This provides a clearer error back to the client if auth fails.
    throw new Error("Could not get a valid OAuth2 service object. Check server logs for details.");
  }
}

/**
 * Resets the authentication service.
 * Useful for clearing cached tokens or fixing auth issues.
 */
function resetAuthentication() {
  // Create the service (credentials needed to identify which service to reset)
  var serviceToReset = getOAuthService_();
  serviceToReset.reset();
  Logger.log("Authentication has been reset.");
}

/**
 * Helper function to retrieve script properties safely.
 * @param {string} key The property key.
 * @returns {string} The property value.
 * @throws {Error} If the property is missing.
 */
function getScriptProperty_(key) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const value = scriptProperties.getProperty(key);
  if (!value) {
    throw new Error('Script property ' + key + ' is missing. Please set it in Project Settings > Script Properties.');
  }
  return value;
}
