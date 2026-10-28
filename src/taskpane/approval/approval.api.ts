import { CONFIG } from "../utils/config";

const baseUrl = {
  toString: () => CONFIG.dataUrl
} as any;

export type ApprovalApiAction = 'APPROVED' | 'UNAPPROVED';

/** Current approval state of a report (also returned by /set). */
export interface ApprovalStatusData {
  ReportHeadID: number;
  IsApproved: boolean;
  CreatedBy: string;
  CreatedDate: string;
  DocumentName: string;
  DocumentType: string;
  DocumentTemplate: string;
}

export interface ApprovalHistoryData {
  ID: number;
  ReportHeadID: number;
  Action: ApprovalApiAction;
  CreatedBy: string;
  CreatedDate: string;
  DocumentName: string;
  DocumentType: string;
  DocumentTemplate: string;
  ReasonForChange: string;
}

/** Calls a report-approval endpoint and returns its Data, throwing with the API Message when Status is false. */
async function approvalRequest<T>(path: string, jwt: string, body?: any): Promise<T> {
  const response = await fetch(`${baseUrl}${path}`, {
    method: body ? 'POST' : 'GET',
    headers: {
      'Authorization': `Bearer ${jwt}`,
      ...(body ? { 'Content-Type': 'application/json' } : {})
    },
    ...(body ? { body: JSON.stringify(body) } : {})
  });

  if (!response.ok) {
    throw new Error('Network response was not ok');
  }

  const data = await response.json();
  if (data?.Status !== true) {
    throw new Error(data?.Message || 'Request failed');
  }
  return data.Data as T;
}

export async function getApprovalStatus(reportHeadId: number, jwt: string): Promise<ApprovalStatusData> {
  return approvalRequest<ApprovalStatusData>(`/api/report-approval/status/${reportHeadId}`, jwt);
}

export async function getApprovalHistory(reportHeadId: number, jwt: string): Promise<ApprovalHistoryData[]> {
  const data = await approvalRequest<ApprovalHistoryData[]>(`/api/report-approval/history/${reportHeadId}`, jwt);
  return Array.isArray(data) ? data : [];
}

export async function setApproval(reportHeadId: number, action: ApprovalApiAction, jwt: string, reasonForChange: string | null = null): Promise<ApprovalStatusData> {
  return approvalRequest<ApprovalStatusData>('/api/report-approval/set', jwt, {
    ReportHeadID: reportHeadId,
    Action: action,
    ReasonForChange: reasonForChange
  });
}
