import { CONFIG } from "../utils/config";

// api.ts
const baseUrl = CONFIG.dataUrl // Set your actual base URL

export async function getSummaryTagsByReportHeadId(reportHeadId: number | string, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/summarytag/reportHead/${reportHeadId}`, {
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
  const response = await fetch(`${baseUrl}/api/report/activate-summarymode`, {
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
export async function refreshSummaryMode(payload: { ReportHeadID: number; RefreshSummaryTag: boolean; ActiveDocument: string }, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/report/refresh-summarymode`, {
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

export async function getSummaryTagHistory(reportHeadSummaryTagID: number | string, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/summarytag/history/${reportHeadSummaryTagID}`, {
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

export async function getSummaryTagStatus(reportHeadSummaryTagID: number | string, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/summarytag/status/${reportHeadSummaryTagID}`, {
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
  const response = await fetch(`${baseUrl}/api/report/summary-history/add`, {
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
  const response = await fetch(`${baseUrl}/api/report/summary-history/update`, {
    method: 'PUT',
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
  const response = await fetch(`${baseUrl}/api/report/summary-tag/add`, {
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
