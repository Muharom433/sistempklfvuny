// This file now handles syncing data via Google Apps Script (GAS) Web App
// rather than direct Sheets API with OAuth, which means users don't need to sign in to Google.

export async function checkAndCreateSheets(spreadsheetId: string) {
  // Handled automatically by the GAS Web App's setupDatabase() function.
  return;
}

export async function saveDataToSheet(gasWebAppUrl: string, stateName: string, jsonData: any) {
  const finalUrl = gasWebAppUrl || import.meta.env.VITE_GAS_WEB_APP_URL;
  if (!finalUrl) {
    console.warn("GAS_WEB_APP_URL is not set. Data will only be saved locally.");
    return;
  }

  try {
    const response = await fetch(finalUrl, {
      method: 'POST',
      // We use text/plain to avoid CORS preflight issues with Apps Script
      headers: { 'Content-Type': 'text/plain' }, 
      body: JSON.stringify({ stateToSave: jsonData }),
      mode: 'cors',
      redirect: 'follow'
    });

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
  } catch (err) {
    console.error("Save to Database failed:", err);
    throw err;
  }
}

export async function loadDataFromSheet(gasWebAppUrl: string, stateName: string) {
  const finalUrl = gasWebAppUrl || import.meta.env.VITE_GAS_WEB_APP_URL;
  if (!finalUrl) {
    console.warn("GAS_WEB_APP_URL is not set. App will use initial local state.");
    return null;
  }

  try {
    const response = await fetch(finalUrl, {
      method: 'GET',
      mode: 'cors',
      redirect: 'follow'
    });

    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const data = await response.json();
    return data;
  } catch (err) {
    console.error("Load from Database failed:", err);
    return null;
  }
}
