
/**
 * Common types for pole data processing
 */

export interface PoleData {
  operationNumber: number;
  attachmentAction: string;
  poleOwner: string;
  poleNumber: string;
  poleStructure: string;
  proposedRiser: string;
  proposedGuy: string;
  pla: string;
  constructionGrade: string;
  heightLowestCom: string;
  heightLowestCpsElectrical: string;
  spans: SpanData[];
  fromPole: string;
  toPole: string;
}

export interface SpanData {
  spanHeader: string;
  attachments: AttachmentData[];
}

export interface AttachmentData {
  description: string;
  existingHeight: string;
  proposedHeight: string;
  midSpanProposed: string;
}

export interface ProcessingError {
  code: string;
  message: string;
  details?: string;
}

// Internal type for Katapult midspan data
export interface KatapultMidspanData {
  com: string;
  electrical: string;
  spans: Map<string, Map<string, string>>;
}
