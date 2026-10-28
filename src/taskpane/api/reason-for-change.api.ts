import { CONFIG } from "../utils/config";

const baseUrl = {
  toString: () => CONFIG.dataUrl
} as any;

export interface ReasonForChangeModule {
  ID: number;
  GroupName: string;
  ModuleName: string;
  IsRequired: number; // 1 = Enable, 0 = Disable
  Status: number; // 1 = Active, 0 = Inactive
}

/** Reason for change configuration of all modules (same API as the LINK AI web app). */
export async function getReasonForChangeModules(jwt: string): Promise<ReasonForChangeModule[]> {
  const response = await fetch(`${baseUrl}/api/reason-for-change/all`, {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${jwt}`
    }
  });

  if (!response.ok) {
    throw new Error('Network response was not ok');
  }

  const data = await response.json();
  if (data?.Status !== true) {
    throw new Error(data?.Message || 'Request failed');
  }

  // Normalize API values (number / string / boolean) to 1 or 0
  return (Array.isArray(data.Data) ? data.Data : []).map((module: any) => ({
    ...module,
    IsRequired: Number(module.IsRequired) ? 1 : 0,
    Status: module.Status === undefined || module.Status === null ? 1 : Number(module.Status)
  }));
}
