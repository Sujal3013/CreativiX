
export interface KeywordItem {
  id: string;
  phrase: string;
}

export interface AdConfigContext {
  platform: string;
  persona: string;
  image_style: string;
  do_not_include: string[];
  aspect_ratio: string;
  brand_tone: string;
  background_preference: string;
  lighting_style: string;
  visual_trends: string;
  composition: string;
  differentiation_strategy: string;
  subject_prominence: string;
  include_keyword_text: boolean; // New flag for text inclusion
}

export interface ImageVariant {
  id: string;
  url: string;
  variantType: 'Product-Focused' | 'Lifestyle-Context';
  status: 'pending' | 'generating' | 'completed' | 'error';
  errorMessage?: string;
}

export interface KeywordResult {
  keyword: KeywordItem;
  variants: ImageVariant[];
}

export interface LogEntry {
  id: string;
  timestamp: string;
  message: string;
  type: 'info' | 'success' | 'warning' | 'error';
}

export enum AppStatus {
  IDLE = 'IDLE',
  PROCESSING = 'PROCESSING',
  COMPLETED = 'COMPLETED',
  QUOTA_EXCEEDED = 'QUOTA_EXCEEDED'
}
