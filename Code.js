/**
 * Code.gs - CONSOLIDATED Server-Side Logic for the Demo Script Generator
 * This version uses a robust JSON-based architecture with a two-step UI flow.
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
    
    // Validate the script object
    if (!scriptObject.title || !scriptObject.steps || !Array.isArray(scriptObject.steps)) {
      throw new Error("AI response was missing required 'title' or 'steps' fields.");
    }

    // Save the new script along with the existing user data
    saveDemoState(savedState.provisionResult, scriptObject);
    return {
      success: true,
      demoScript: scriptObject
    };
  } catch (e) {
    console.error(`[Code.gs] Failed to parse or validate demo script JSON. Error: ${e.message}. Raw Text: ${scriptGenerationResult.text}`);
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
  try {
    const userProperties = PropertiesService.getUserProperties();
    if (provisionResult) {
      userProperties.setProperty('savedProvisionResult', JSON.stringify(provisionResult));
    }
    if (demoScript) {
      userProperties.setProperty('savedDemoScript', JSON.stringify(demoScript));
    }
  } catch(e) {
    console.error("Error saving state to UserProperties: " + e.toString());
  }
}

function loadDemoState() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const savedProvisionResult = userProperties.getProperty('savedProvisionResult');
    
    if (savedProvisionResult) {
      const savedDemoScript = userProperties.getProperty('savedDemoScript');
      const result = {
        provisionResult: JSON.parse(savedProvisionResult),
        demoScript: savedDemoScript ? JSON.parse(savedDemoScript) : null
      };
      // Never send the password back on initial load for security.
      if (result.provisionResult.password) {
        delete result.provisionResult.password;
      }
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
 * Internal function to clear UserProperties.
 */
function clearDemoState() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty('savedProvisionResult');
  userProperties.deleteProperty('savedDemoScript');
  console.log('Internal call: Demo state cleared.');
}

/**
 * Clears only the saved demo script from UserProperties.
 * Accessible from the client side for the "Start Over" button (for script only).
 */
function clearDemoScriptOnly() {
  PropertiesService.getUserProperties().deleteProperty('savedDemoScript');
  console.log('Demo script data cleared.');
}

/**
 * [FIXED] Clears only the application-specific properties for the current user.
 * This is used for the "Start Over" button.
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

// --- Authentication Functions ---
function getOAuthService_(privateKey, clientEmail) {
  const impersonatedUser = "mawi@cymbal.se";

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

function getService_() {
  const private_key = "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDsphzrkBjIePMg\nyYGBGOPkewPKw2T4fbbfA4QQYnqP+GxN2kuve+E5mcDytGyESBUa9eM2q8LMbQ7U\nzqkntCG2Ug7CY+9iiVm8bHwrKXjzeDEvOlN1++gpnzQH15up8Uq8fq/JWtztYi0L\n1U4NK6B/aCDyhYO+oX1OQyj4DvZS35pzBqAbw/CmG3pHmh/CmG1u/lk3diAbmMiW\nwvDNnKxKODGPiVI/rp1w6DfdjzcTwg8hiJ3+lWubcuTZ3Z2FmRluRd8XFwDwqbuc\nKfosxsQgHw13heCoALfZILnew4yOR0YsR43kyaHofagieL94OGgR8MTwo0ym6lHV\nSbxeXaOVAgMBAAECggEAAaC0Tl0QTW791I6L/EfW57zsBPg8qw93c1VkpKfsLpwM\nDGDBXztrjGOFteEXrtmI76EJx0W0ds/vOdMS8BLneVDtmzRDe+hiYPzFL032FKl/\nVtMuPMi7lG9s+WiMm+nana0oi/PNHjm+SyYvq3eqI+Gi+hgTJvu9hfrJkSPOWYMj\njgdo8p3s6FXM5U6V2jt0rQ9/m9q8KQyNbhhA8eEiaUZ/jwM8DIwoFPW+icMH2NGW\nBz7pT3feJ6BqVI3a86Rf+kxo18MCGe7VhYTs9sRCtefiMuPoHTmMIXz6imoRoTcm\nYSqoSG14uFD9j2S9SBRN5kPVZ3DTkMFlEXJ9t118JwKBgQD4nhOpRrPVl/xxkwq2\n7re3dr2174QTyMAsLmKiEyt5LlBVecktZAVNZJ2rpnRk8ul9ZqzxdZt5LVAUagO4\nENsYO2/qErONu3EcQjqWkStKeJrOW2NJNMnZgT6ta0PQxxsaWUG6PAcOp4+TMBwj\nchxLmQG3szIrRzeVoj1o1SaXdwKBgQDzrQ3Us99VcPBgOOQZD36ERPtW3FYm5OCx\n8d2BpEoslr6egKH4R7qkJv/99BgpteqgRClIIE0IlNW58PFCBlgUh8KHHOSFiST5\nhwgkiw0mFDcn6fqDAyAIJQjaIZ50wEezaZ3Emiy6jxlkrjTMPGBaIYxNyLByiWFN\nLx7P0yu4UwKBgBY586IHix5GVzBEKAoQr2X8fJteTV2DbgLFJtY8hn9v74ikuaKQ\nNZUksJ/e4rr/qHYojr+LdxnPPkCE9c4n255/+dJgV6MNJeCT3y8EzWz7+UMHkonB\n6WXDkznnxAlPM5IYdrLSmQLrYf+TpoBYvETZ6fhlUc/irwp2lazgmXGjAoGBAIBN\nTyn+p4oqVDal3dwgH2Jvm9MpYqdJ/dFT42ieY3vEx4tXeXDr+6bw7fr+KjbUFTzb\nhsz2TPlGvJ4R8kXsZzYwIUnY+a4h/vjvk2cCXCL/o+b9OK0A2T3Qmi+YYgFhOJ+L\n7ckV0JVOQXWUkDI1XBo47dIK6HT2RuhH9jZBHxUHAoGAAunSPHqU2XU1HP0X57eo\n1DVKkR8ZbUwZPl2Km9KZWsOym26ChsDPOeoDqgFBXjEPFsYdcJ79dEkjByMIJImK\nKf0btVWkgVAWeXAbCoWDieMlWe+G5Q0cFyHZquFZMJu+2jRwOTduqCGAFOT3OwzE\nWhR+uwkARi6nRqaxpXbwitY=\n-----END PRIVATE KEY-----\n";
  const client_email = "ai-demos@cymbal-workshops.iam.gserviceaccount.com";

  var service = getOAuthService_(private_key, client_email);

  if (service && service.hasAccess()) {
    return service;
  } else {
    Logger.log('getService_(): Access denied. Last Error: ' + (service ? service.getLastError() : "Service object was null."));
    return null;
  }
}

function getAdminAccessToken() {
  var service = getService_();
  if (service) {
    return service.getAccessToken();
  } else {
    // This provides a clearer error back to the client if auth fails.
    throw new Error("Could not get a valid OAuth2 service object. Check server logs for details.");
  }
}

function resetAuthentication() {
  const private_key = "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDsphzrkBjIePMg\nyYGBGOPkewPKw2T4fbbfA4QQYnqP+GxN2kuve+E5mcDytGyESBUa9eM2q8LMbQ7U\nzqkntCG2Ug7CY+9iiVm8bHwrKXjzeDEvOlN1++gpnzQH15up8Uq8fq/JWtztYi0L\n1U4NK6B/aCDyhYO+oX1OQyj4DvZS35pzBqAbw/CmG3pHmh/CmG1u/lk3diAbmMiW\nwvDNnKxKODGPiVI/rp1w6DfdjzcTwg8hiJ3+lWubcuTZ3Z2FmRluRd8XFwDwqbuc\nKfosxsQgHw13heCoALfZILnew4yOR0YsR43kyaHofagieL94OGgR8MTwo0ym6lHV\nSbxeXaOVAgMBAAECggEAAaC0Tl0QTW791I6L/EfW57zsBPg8qw93c1VVkpKfsLpwM\nDGDBXztrjGOFteEXrtmI76EJx0W0ds/vOdMS8BLneVDtmzRDe+hiYPzFL032FKl/\nVtMuPMi7lG9s+WiMm+nana0oi/PNHjm+SyYvq3eqI+Gi+hgTJvu9hfrJkSPOWYMj\njgdo8p3s6FXM5U6V2jt0rQ9/m9q8KQyNbhhA8eEiaUZ/jwM8DIwoFPW+icMH2NGW\nBz7pT3feJ6BqVI3a86Rf+kxo18MCGe7VhYTs9sRCtefiMuPoHTmMIXz6imoRoTcm\nYSqoSG14uFD9j2S9SBRN5kPVZ3DTkMFlEXJ9t118JwKBgQD4nhOpRrPVl/xxkwq2\n7re3dr2174QTyMAsLmKiEyt5LlBVecktZAVNZJ2rpnRk8ul9ZqzxdZt5LVAUagO4\nENsYO2/qErONu3EcQjqWkStKeJrOW2NJNMnZgT6ta0PQxxsaWUG6PAcOp4+TMBwj\nchxLmQG3szIrRzeVoj1o1SaXdwKBgQDzrQ3Us99VcPBgOOQZD36ERPtW3FYm5OCx\n8d2BpEoslr6egKH4R7qkJv/99BgpteqgRClIIE0IlNW58PFCBlgUh8KHHOSFiST5\nhwgkiw0mFDcn6fqDAyAIJQjaIZ50wEezaZ3Emiy6jxlkrjTMPGBaIYxNyLByiWFN\nLx7P0yu4UwKBgBY586IHix5GVzBEKAoQr2X8fJteTV2DbgLFJtY8hn9v74ikuaKQ\nNZUksJ/e4rr/qHYojr+LdxnPPkCE9c4n255/+dJgV6MNJeCT3y8EzWz7+UMHkonB\n6WXDkznnxAlPM5IYdrLSmQLrYf+TpoBYvETZ6fhlUc/irwp2lazgmXGjAoGBAIBN\nTyn+p4oqVDal3dwgH2Jvm9MpYqdJ/dFT42ieY3vEx4tXeXDr+6bw7fr+KjbUFTzb\nhsz2TPlGvJ4R8kXsZzYwIUnY+a4h/vjvk2cCXCL/o+b9OK0A2T3Qmi+YYgFhOJ+L\n7ckV0JVOQXWUkDI1XBo47dIK6HT2RuhH9jZBHxUHAoGAAunSPHqU2XU1HP0X57eo\n1DVKkR8ZbUwZPl2Km9KZWsOym26ChsDPOeoDqgFBXjEPFsYdcJ79dEkjByMIJImK\nKf0btVWkgVAWeXAbCoWDieMlWe+G5Q0cFyHZquFZMJu+2jRwOTduqCGAFOT3OwzE\nWhR+uwkARi6nRqaxpXbwitY=\n-----END PRIVATE KEY-----\n";
  const client_email = "ai-demos@cymbal-workshops.iam.gserviceaccount.com";

  var serviceToReset = getOAuthService_(private_key, client_email);
  serviceToReset.reset();
  Logger.log("Authentication has been reset.");
}