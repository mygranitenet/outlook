/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, fetch, DOMParser */

const GEMINI_API_KEY_SETTING = "geminiApiKey";

// This is the entry point for the add-in.
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    run();
  }
});

function run() {
  // --- Get references to all our DOM elements ---
  const apiKeySection = document.getElementById("apiKeySection");
  const mainContent = document.getElementById("mainContent");
  const apiKeyInput = document.getElementById("apiKeyInput");
  const saveKeyButton = document.getElementById("saveKeyButton");
  const changeKeyButton = document.getElementById("changeKeyButton");
  const apiKeyError = document.getElementById("apiKeyError");

  const summarizeButton = document.getElementById("summarizeButton");
  const actionsButton = document.getElementById("actionsButton");
  const replyButton = document.getElementById("replyButton");
  const customPromptButton = document.getElementById("customPromptButton");
  const customPrompt = document.getElementById("customPrompt");

  const spinner = document.getElementById("spinner");
  const generalError = document.getElementById("generalError");
  const resultCard = document.getElementById("resultCard");
  const resultText = document.getElementById("resultText");
  
  // --- State management ---
  let currentApiKey = null;

  // --- UI Update Functions ---
  function showView(isKeySaved) {
    if (isKeySaved) {
      apiKeySection.style.display = "none";
      mainContent.style.display = "flex";
    } else {
      apiKeySection.style.display = "flex";
      mainContent.style.display = "none";
    }
  }

  function showLoading(isLoading) {
    spinner.style.display = isLoading ? "flex" : "none";
    if (isLoading) {
      clearStatus(); // Clear previous results/errors
    }
  }

  function displayError(message) {
      generalError.innerText = message ? `Error: ${message}` : "";
      generalError.style.display = message ? "block" : "none";
  }

  function displayResult(text) {
      resultText.innerText = text;
      resultCard.style.display = "block";
  }
  
  function clearStatus() {
      displayError(null);
      resultCard.style.display = "none";
      resultText.innerText = "";
  }

  // --- Core Logic Functions ---
  function loadApiKey() {
    currentApiKey = Office.context.roamingSettings.get(GEMINI_API_KEY_SETTING);
    if (currentApiKey) {
      apiKeyInput.value = currentApiKey; // for display, though it's a password field
      showView(true);
    } else {
      showView(false);
    }
  }

  function handleSaveKey() {
    const key = apiKeyInput.value;
    if (!key || key.trim() === "") {
        apiKeyError.innerText = "API Key cannot be empty.";
        return;
    }
    apiKeyError.innerText = "";
    Office.context.roamingSettings.set(GEMINI_API_KEY_SETTING, key);
    Office.context.roamingSettings.saveAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        currentApiKey = key;
        showView(true);
      } else {
        displayError("Could not save API key. " + asyncResult.error.message);
      }
    });
  }
  
  function handleChangeKey() {
      currentApiKey = null;
      apiKeyInput.value = "";
      Office.context.roamingSettings.remove(GEMINI_API_KEY_SETTING);
      Office.context.roamingSettings.saveAsync(() => showView(false));
  }

  async function getEmailThread() {
    return new Promise((resolve, reject) => {
      const conversation = Office.context.mailbox.item.conversation;
      conversation.getItemIdsAsync(async (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          return reject(new Error(asyncResult.error.message));
        }
        
        const itemIds = asyncResult.value;
        if (!itemIds || itemIds.length === 0) {
          return resolve("This email is not part of a conversation.");
        }

        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
          if (result.status === "succeeded") {
            const accessToken = result.value;
            let formattedThread = "";
            
            // We fetch emails one by one to preserve order (Promise.all doesn't guarantee order)
            for (const itemId of itemIds) {
                const message = await getEmailBody(itemId, accessToken);
                if (message) {
                    const plainTextBody = new DOMParser().parseFromString(message.body.content, "text/html").documentElement.textContent;
                    formattedThread += `--- Email ---\nFrom: ${message.from.emailAddress.name} (${message.from.emailAddress.address})\nSubject: ${message.subject}\nDate: ${new Date(message.receivedDateTime).toLocaleString()}\n\n${plainTextBody.trim()}\n\n`;
                }
            }
            resolve(formattedThread);

          } else {
            reject(new Error("Could not get access token."));
          }
        });
      });
    });
  }

  async function getEmailBody(itemId, accessToken) {
    const restUrl = `${Office.context.mailbox.restUrl}/v2.0/me/messages/${itemId}?$select=subject,from,receivedDateTime,body`;
    try {
        const response = await fetch(restUrl, {
            headers: { Authorization: `Bearer ${accessToken}` }
        });
        if (!response.ok) return null;
        return await response.json();
    } catch (error) {
        console.error(`Error fetching message ${itemId}:`, error);
        return null;
    }
  }

  async function callGeminiApi(prompt, threadContent) {
    const model = "gemini-pro";
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${currentApiKey}`;

    const fullPrompt = `Based on the following email thread, please perform this task: "${prompt}"\n\nHere is the email thread:\n===========================\n${threadContent}\n===========================`;
    
    const requestBody = { contents: [{ parts: [{ text: fullPrompt }] }] };

    const response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(requestBody),
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error?.message || "Failed to call Gemini API");
    }

    const data = await response.json();
    try {
      return data.candidates[0].content.parts[0].text;
    } catch (e) {
      throw new Error("Could not parse Gemini's response. The response may have been blocked or the format is unexpected.");
    }
  }

  async function runAnalysis(prompt) {
    if (!currentApiKey) {
        displayError("API Key is not set.");
        return;
    }
    if (!prompt) {
        displayError("The request prompt cannot be empty.");
        return;
    }
    
    showLoading(true);

    try {
        const threadContent = await getEmailThread();
        const geminiResponse = await callGeminiApi(prompt, threadContent);
        displayResult(geminiResponse);
    } catch (err) {
        displayError(err.message);
    } finally {
        showLoading(false);
    }
  }
  
  // --- Attach Event Listeners ---
  saveKeyButton.addEventListener("click", handleSaveKey);
  changeKeyButton.addEventListener("click", handleChangeKey);
  
  summarizeButton.addEventListener("click", () => runAnalysis("Summarize this thread in a few bullet points."));
  actionsButton.addEventListener("click", () => runAnalysis("Extract all action items from this thread. For each item, list who is responsible and the due date if mentioned."));
  replyButton.addEventListener("click", () => runAnalysis("Draft a polite and professional reply that acknowledges the last email and confirms we will look into their request."));
  customPromptButton.addEventListener("click", () => runAnalysis(customPrompt.value));

  // --- Initial Load ---
  loadApiKey();
}
