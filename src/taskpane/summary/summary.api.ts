import { CONFIG } from "../utils/config";

// api.ts
const baseUrl = CONFIG.dataUrl // Set your actual base URL

export async function getSummaryTagsByWorkbenchId(workbenchId: number | string, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/summary-tag/workbench/${workbenchId}`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    }
  });

  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }

  return await response.json();
}


export async function activateSummaryMode(payload: any, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/summary-tag/activate-mode`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }

  return await response.json(); // if API returns JSON
}

export async function refreshSummaryMode(payload: { WorkbenchID: number; RefreshSummaryTag: boolean; ActiveDocument: string }, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/summary-tag/refresh-mode`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }

  return await response.json();
}

export async function getSummaryTagHistory(workbenchSummaryTagID: number | string, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/summarytag/history/${workbenchSummaryTagID}`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    }
  });

  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }

  return await response.json();
}

export async function getSummaryTagStatus(workbenchId: number | string, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/summarytag/status/${workbenchId}`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    }
  });

  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }

  return await response.json();
}

export async function addSummaryHistory(payload: any, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/summary-history/add`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }

  return await response.json();
}

export async function updateSummaryHistory(payload: any, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/summary-history/update`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }

  return await response.json();
}

export async function addSummaryTag(payload: any, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/summary-tag/add`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }

  return await response.json();
}

export async function updateSummaryTagPrompt(payload: { Name: string; Prompt: string }, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/summarytag/update-prompt`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }

  return await response.json();
}
