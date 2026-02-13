import { CONFIG } from "../utils/config";

// api.ts
const baseUrl = CONFIG.dataUrl // Set your actual base URL

export async function loginUser(organization: string, username: string, password: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/login`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      CompanyName: organization,
      Username: username,
      Password: password,
      "ApplicationCode": "LINK"
    })
  });

  if (!response.ok) {
    throw new Error('Network response was not ok');
  }

  const data: any = await response.json();
  return data;
}


// api.ts

export async function getReportById(WorkbenchID: string, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/workbench/${WorkbenchID}`, {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${jwt}`
    }
  });

  if (!response.ok) {
    throw new Error('Network response was not ok');
  }

  const data = await response.json();
  return data;
}


export async function getAllSponsors(jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/sponsor/all`, {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${jwt}`
    }
  });

  if (!response.ok) {
    throw new Error('Network response was not ok');
  }

  const data: any = await response.json();
  return data;
}


export async function getAiHistory(tagId: string, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/workbench/ai-history/${tagId}`, {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${jwt}`
    }
  });

  if (!response.ok) {
    throw new Error('Network response was not ok');
  }

  const data: any = await response.json();
  return data;
}

export async function updateGroupKey(tag: any, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/workbench/group-key`, {
    method: 'PUT',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(tag)
  });

  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }

  const data: any = await response.json();
  return data;
}


export async function addAiHistory(payload: any, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/ai-history-add`, {
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

  const data: any = await response.json();
  return data;
}


export async function updateAiHistory(payload: any, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/ai-history-update`, {
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

  const data: any = await response.json();
  return data;
}

export async function fetchGlossaryTemplate(sponsorID: string, bodyText: any, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/glossary-client/${sponsorID}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(bodyText)
  });

  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }

  const data: any = await response.json();
  return data;
}


export async function addGroupKey(payload: any, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/group-key/add`, {
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

  const data: any = await response.json();
  return data;
}


export async function getAllPromptTemplates(jwt): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/prompt-builders/all`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
  });

  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }

  const data: any = await response.json();
  return data;
}


export async function getPromptTemplateById(id: string, jwt): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/prompt-builders/${id}/data`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
  });

  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }

  const data: any = await response.json();
  return data;
}

export async function updatePromptTemplate(payload: any, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/groupkey/update-prompt`, {
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

  const data: any = await response.json();
  return data;
}

export async function getAllCustomTables(jwt): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/custom-table/all`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
  });

  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }

  const data: any = await response.json();
  return data;
}

export async function getGeneralImages(jwt): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/image/general`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    }
  });

  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }

  const data: any = await response.json();
  return data;
}

export async function getReportHeadImageById(id: string, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/addin/image/workbench/${id}`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    }
  });

  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  const data: any = await response.json();
  return data;
};