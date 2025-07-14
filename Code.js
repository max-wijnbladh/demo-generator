/**
 * Code.gs - CONSOLIDATED Server-Side Logic for the Demo Script Generator
 * This version uses a robust JSON-based architecture with a two-step UI flow.
 *
 * This version handles existing users by showing their info without a password
 * and providing a separate function to reset the password on demand.
 */

// --- Configuration Constants ---
const PRIMARY_DOMAIN_CONFIG = 'cymbal.se';
const DEMO_SCRIPT_GEMINI_MODEL_CONFIG = 'models/gemini-1.5-flash-latest';

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
  const scriptGenerationResult = callGeminiAPI(demoScriptPrompt, DEMO_SCRIPT_GEMINI_MODEL_CONFIG);

  if (!scriptGenerationResult.success || !scriptGenerationResult.text) {
    console.error('[Code.gs] Demo script generation failed:', scriptGenerationResult.error);
    return {
      success: false,
      error: `Demo script generation failed: ${scriptGenerationResult.error}`
    };
  }

  try {
    const scriptObject = JSON.parse(scriptGenerationResult.text);

    // --- NEW: VALIDATION BLOCK ---
    // Check if the parsed object has the keys required by the front-end.
    if (!scriptObject.title || !scriptObject.steps || !Array.isArray(scriptObject.steps)) {
      console.error('[Code.gs] AI returned valid JSON but with missing required keys (title/steps).', JSON.stringify(scriptObject));
      return { 
        success: false, 
        error: "The AI's response was missing the required 'title' or 'steps' fields. Please try generating the script again." 
      };
    }
    // --- END: VALIDATION BLOCK ---

    // Save the new script along with the existing user data
    saveDemoState(savedState.provisionResult, scriptObject);
    return {
      success: true,
      demoScript: scriptObject
    };
  } catch (e) {
    console.error(`[Code.gs] Failed to parse demo script JSON. Error: ${e.message}. Raw Text: ${scriptGenerationResult.text}`);
    return {
      success: false,
      error: `The AI returned malformed data that could not be read. Please try again.`
    };
  }
}

// =====================================================================================
// State Management Functions
// =====================================================================================
function saveDemoState(provisionResult, demoScript) {
  const userProperties = PropertiesService.getUserProperties();
  if (provisionResult) {
    userProperties.setProperty('savedProvisionResult', JSON.stringify(provisionResult));
  }
  if (demoScript) {
    userProperties.setProperty('savedDemoScript', JSON.stringify(demoScript));
  }
}

function loadDemoState() {
  const userProperties = PropertiesService.getUserProperties();
  const savedProvisionResult = userProperties.getProperty('savedProvisionResult');
  const savedDemoScript = userProperties.getProperty('savedDemoScript');

  if (savedProvisionResult) {
    const result = {
      provisionResult: JSON.parse(savedProvisionResult),
      demoScript: savedDemoScript ? JSON.parse(savedDemoScript) : null
    };
    // Never send the password back on initial load for security.
    if(result.provisionResult.password) {
      delete result.provisionResult.password;
    }
    return result;
  }
  return null;
}

/**
 * Internal function to clear UserProperties.
 */
function clearDemoState() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty('savedProvisionResult');
  userProperties.deleteProperty('savedDemoScript');
  console.log('Internal call: Demo state cleared.');
}

/**
 * Clears all user properties for the current user.
 * This is used to reset the client-side state without deleting the user account.
 */
function clearUserData() {
  PropertiesService.getUserProperties().deleteAllProperties();
  console.log('All user properties cleared.');
}

/**
 * [FIXED] Deletes the demo account user.
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

  console.log(`Attempting to delete user: ${userEmail}`);
  const url = `https://admin.googleapis.com/admin/directory/v1/users/${userEmail}`;
  const options = {
    method: "DELETE",
    headers: { "Authorization": "Bearer " + token },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    // 204 No Content is success. 404 means it was already gone. Both are "success" for our purposes.
    if (responseCode === 204 || responseCode === 404) {
      console.log(`Successfully processed deletion for user ${userEmail} (Status: ${responseCode}).`);
      clearDemoState(); // Clear state since the user is gone.
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
// Helper & Core Logic Functions
// =====================================================================================
function constructDemoScriptPrompt_(demoContext, demoUserFirstName, demoUserLastName, demoUserEmail) {
  return `
    You are an expert demo script writer for Google Workspace. Your task is to generate a step-by-step demo script based on user context.
    
    **Output Format:**
    Return ONLY a valid JSON object. Do not include any other text, Markdown formatting, or code fences like \`\`\`json.
    
    **JSON Schema:**
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
    
    **CRITICAL RULES:**
    - The entire output MUST be a single, valid JSON object.
    - Do NOT include any unescaped quotation marks (single or double) within string values.
    - Do NOT use markdown like asterisks or backticks. Provide only raw text.
    - If there are no specific prerequisites for the demo, provide one generic one like "Practice the script before the live demo."
    - The "prerequisites" field must be a top-level key and its value must be an array of strings.

    **Demo User & Context:**
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

  let requestingUserPrefix = requestingUserEmail.split('@')[0].toLowerCase().replace(/[^a-z0-9]/g, '');
  if (requestingUserEmail === "mawi@google.com") {
    requestingUserPrefix = "mawi" + Math.floor(Math.random() * 1000);
  }
  const primaryEmail = `${requestingUserPrefix}@${PRIMARY_DOMAIN_CONFIG}`;

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
    apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) throw new Error('GEMINI_API_KEY not found in Script Properties.');
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

// --- Authentication Functions (No changes) ---
function getOAuthService_(){const e="service@cymbal-workshops.iam.gserviceaccount.com",t="-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDVbFypp0IRz/Bc\nnPn+i6afegr9Z6ZmDlPqF98i81GEBwirebfaqxuY2bUwihWVDZtGbfvOsspitfjl\nZOi4V5rdDv/1ZZZmEROjuMdhZvBZXxcGiImpstR90FsF9C+7jH5F6d4LUrPrk0gd\n1Mdc+Bpiu5Ww0Rci4e4VPuBYAA4VBFc1LlpICkVkpCG3d3hRo6c9DZ1kHWpMUsrc\n65Zp/MCG9fE5hXsUrY7Ufe6jqvJB2LATZIb7bbons28acsCooIGtiyXWUTudR12v\nbYuICUSNamh0SUcUDvTkdOrVQCmYrFpEOCInbpg/OavO7ryuffQo+GrgM9gfJVFU\nYJiBihtXAgMBAAECggEAJ2rEEnFZuoB1HCXB5klUlM+th+/Ew8SRqwKNq57Ux1Wl\nPEZWtoQzrJ9I35YhNk41B2T4xMwwpNqHBZcFhEZpy7ohe+kvRdqRjgNqj4q7iUYO\nsp41DqqApFv+87KNvk3MZI00/VJg+HlTMG9EAt+vv9x1YRq88yxXFIVwWdBoyWiV\neNuoZAY4NoRppCB7KCUGsmz1pIPgYG9/cSmLOXQBxcjiXUTk332l334anuB1T38Q\n+aN8sBSyOgzIWCjv7W8cC/cRDMq67yn6JbRotW4PzwNwbdn96R8gaefZzTTFOQyZ\nN35qCSFHBU/c98dQRR2vnllPDKg9VrADnUwRTo2awQKBgQDtON2z0197g8WROyBV\nGyQecl/Zdoz1gjRXMm2p4bipxZJ/kcpJjf9pyG0EDlUwEoZKxZElctuG1o0OM0b/\n9peuB5FIo2IIJL02jP7/+3QuQyUkK0U+nJ5XEPhSzYmhXo09p/VgTFa2JsZhZM2K\nOtWs0uAcaO4hVNSnW1oBQsZU0wKBgQDmUTrMBmDCKgkHwq7MZmDDr1w1mXCLCGMm\nU9crPgDZIO4vXiURSQm84VfvjdM3cQxWOvM53Xr4oYp+k+Wzs5Pgx+tgP2AQPLVF\nbZZXOjU1fB2vh+g0ZI1ZG+YYr0BhBMjw3sOnzwKPjjzCjejaNGtvhUg+xfcpszDO\nPQ0MyK+c7QKBgQDeJlvQNEj9hTg2OjWcHY+kh51lK9TzcNyNL+dsqLpjGmeH2cKj\nQTwIFy6oFrgGDcL/MKctd7NHQZLU0oZR297Nlb6jVIXQdH9RH5cJp7R0QmL8zRzK\ndqb9iCHUgTC7Eq2YKLrsVHD7obIzsM+e/FvvvYcsc8NVKXj/xNezyJGtCwKBgQDW\n1xzepnBpjiaAS7UcS7+lqhV8lhXqSzeZ0Aldd+f4ooQsQUiYeCYSP630crp89AIL\nCdBKwPPtq1pyOmnBmBiwTCyeyl9Epix9h/z+fviVXKKgU0liXg2P+rtHeWq3VWxP\na6zdAvgjiw3YeeGkcdNp4s0CaU3mYxV6vG5I54cQ/QKBgHf4sKDXQTTeQa5RqYap\ns1u3g+X2b2JjqQxFJVIpQSPiJxVx/MvkXUsAtxgXJAPljBVqWvLohGbpWxIxMzO0\nGEqe+kfaVyhgXgsgpcTxvV2Viv8L/EYA6a2JMHtijMMUrwX+yLBuy5WCSpp5ViNi\nfi9VSNqDV0jEVGRlyVmm2Pq9\n-----END PRIVATE KEY-----\n",o="mawi@google.com",r=["https://www.googleapis.com/auth/admin.directory.user"];return OAuth2.createService("DemoGeneratorService_V5").setTokenUrl("https://accounts.google.com/o/oauth2/token").setPrivateKey(t).setIssuer(e).setSubject(o).setPropertyStore(PropertiesService.getScriptProperties()).setCache(CacheService.getScriptCache()).setScope(r.join(" "))}
function getService_(){const e=getOAuthService_();return e&&e.hasAccess()?e:(console.error("Failed to get a valid OAuth2 service or the service has no access."),null)}
function getAdminAccessToken(){try{const e=getService_();if(e)return e.getAccessToken();throw new Error("Could not get a valid OAuth2 service object.")}catch(e){return console.error("Exception in getAdminAccessToken:",e.toString()),null}}
function resetAuthentication(){getOAuthService_().reset()}