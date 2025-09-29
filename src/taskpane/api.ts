import { dataUrl } from "./data";

// api.ts
const baseUrl = dataUrl // Set your actual base URL

export async function loginUser(organization: string, username: string, password: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/user/login`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      ClientName: organization,
      Username: username,
      Password: password
    })
  });

  if (!response.ok) {
    throw new Error('Network response was not ok');
  }

  const data: any = await response.json();
  return data;
}


// api.ts

export async function getReportById(documentID: string, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/report/id/${documentID}`, {
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


export async function getAllClients(userId: string, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/client/all/${userId}`, {
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
  const response = await fetch(`${baseUrl}/api/report/ai-history/${tagId}`, {
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
  const response = await fetch(`${baseUrl}/api/report/head/groupkey`, {
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
  const response = await fetch(`${baseUrl}/api/report/ai-history/add`, {
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
  const response = await fetch(`${baseUrl}/api/report/ai-history/update`, {
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

export async function fetchGlossaryTemplate(clientId: string, bodyText: any, jwt: string): Promise<any> {
  const response = await fetch(`${baseUrl}/api/glossary-template/client-id/${clientId}`, {
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
  const response = await fetch(`${baseUrl}/api/report/group-key/add`, {
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
  const response = await fetch(`${baseUrl}/api/prompt-template/all`, {
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
  const response = await fetch(`${baseUrl}/api/prompt-template/${id}/data`, {
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
  const response = await fetch(`${baseUrl}/api/groupkey/update-prompt`, {
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
  const response = await fetch(`${baseUrl}/api/custom-table/all`, {
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