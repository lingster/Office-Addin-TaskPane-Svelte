// Svelte store type definitions
import type { Writable, Readable } from 'svelte/store'
import type { OfficeContext, UserInfo, AuthError } from './office'

// Store interfaces
export interface OfficeStore extends Readable<OfficeContext | null> {
  initialize: () => Promise<void>
  reset: () => void
}

export interface AuthStore extends Writable<UserInfo | null> {
  authenticate: () => Promise<UserInfo>
  signOut: () => Promise<void>
  refresh: () => Promise<UserInfo>
}

export interface ErrorStore extends Writable<AuthError | null> {
  setError: (error: string | Error, code?: string) => void
  clearError: () => void
}

// State management types
export interface AppState {
  office: OfficeContext | null
  user: UserInfo | null
  error: AuthError | null
  isLoading: boolean
  isInitialized: boolean
}

export type StoreUpdater<T> = (value: T) => T
export type StoreSubscriber<T> = (value: T) => void
