// Office.js specific type definitions and extensions

export interface OfficeContext {
  host: Office.HostType
  platform: Office.PlatformType
  isReady: boolean
  requirements: Office.RequirementSetSupport
  version?: string
}

export interface AuthConfig {
  clientId: string
  tenantId: string
  apiScope: string
  backendUrl: string
  environment: 'development' | 'staging' | 'production'
}

export interface UserInfo {
  user: string | null
  email: string | null
  oid: string
  tenant: string
  upn?: string
  roles: string[]
  authenticated_at: string
}

export interface AuthError {
  error: string
  error_description: string
  error_code?: string
  timestamp: string
}

// Office.js error codes
export enum OfficeErrorCodes {
  USER_CANCELLED = '13001',
  USER_NOT_SIGNED_IN = '13002',
  CONSENT_REQUIRED = '13003',
  TOKEN_FAILED = '13004',
  NETWORK_ERROR = '13005',
  INVALID_GRANT = '13006',
  UNKNOWN_ERROR = '13000'
}

// Extended Office API types
export interface DocumentProperties extends Office.FileProperties {
  lastModified?: Date
  author?: string
  version?: string
}

export interface InsertOptions {
  coercionType?: Office.CoercionType
  cellFormat?: Office.TableData | Office.TableBinding
  imageLeft?: number
  imageTop?: number
  imageWidth?: number
  imageHeight?: number
}

// Host-specific feature detection
export interface HostCapabilities {
  supportsSso: boolean
  supportsRibbon: boolean
  supportsTaskpane: boolean
  supportsContentApps: boolean
  supportsDialog: boolean
  version: string
}

// Event types for Office.js
export interface OfficeEvent<T = unknown> {
  type: string
  data: T
}

export type OfficeEventHandler<T = unknown> = (event: OfficeEvent<T>) => void

// Utility types
export type OfficeHost = keyof typeof Office.HostType
export type OfficePlatform = keyof typeof Office.PlatformType
export type AsyncResult<T> = Office.AsyncResult<T>
