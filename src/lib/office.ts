import { writable, derived, type Writable, type Readable } from 'svelte/store'
import type {
  OfficeContext,
  AuthConfig,
  UserInfo,
  AuthError,
  DocumentProperties,
  InsertOptions
} from '../types/office'

// Svelte stores for reactive state management
export const officeContext: Writable<OfficeContext | null> = writable(null)
export const isOfficeReady: Writable<boolean> = writable(false)
export const currentUser: Writable<UserInfo | null> = writable(null)
export const authError: Writable<AuthError | null> = writable(null)
export const isLoading: Writable<boolean> = writable(false)

// Derived stores
export const isAuthenticated: Readable<boolean> = derived(
  currentUser,
  ($currentUser) => $currentUser !== null
)

export const hostType: Readable<Office.HostType | null> = derived(
  officeContext,
  ($context) => $context?.host ?? null
)

export const isWordHost: Readable<boolean> = derived(
  hostType,
  ($hostType) => $hostType === Office.HostType.Word
)

export const isExcelHost: Readable<boolean> = derived(
  hostType,
  ($hostType) => $hostType === Office.HostType.Excel
)

export const isPowerPointHost: Readable<boolean> = derived(
  hostType,
  ($hostType) => $hostType === Office.HostType.PowerPoint
)

class OfficeService {
  private config: AuthConfig
  private initialized = false
  private initializationPromise: Promise<void> | null = null

  constructor() {
    this.config = {
      clientId: import.meta.env.VITE_CLIENT_ID,
      tenantId: import.meta.env.VITE_TENANT_ID,
      apiScope: import.meta.env.VITE_API_SCOPE,
      backendUrl: import.meta.env.VITE_BACKEND_URL,
      environment: import.meta.env.VITE_ENVIRONMENT || 'development'
    }

    // Validate configuration
    this.validateConfig()

    // Initialize Office.js
    this.initialize()
  }

  private validateConfig(): void {
    const required = ['clientId', 'tenantId', 'apiScope', 'backendUrl']
    const missing = required.filter(key => !this.config[key as keyof AuthConfig])

    if (missing.length > 0) {
      throw new Error(`Missing required configuration: ${missing.join(', ')}`)
    }
  }

  async initialize(): Promise<void> {
    if (this.initializationPromise) {
      return this.initializationPromise
    }

    this.initializationPromise = this.doInitialize()
    return this.initializationPromise
  }

  private async doInitialize(): Promise<void> {
    if (this.initialized) return

    try {
      isLoading.set(true)

      // Wait for Office.js to be available
      await this.waitForOffice()

      // Initialize Office.js
      await new Promise<void>((resolve) => {
        Office.onReady((info) => {
          console.log('Office.js initialized:', info)

          const context: OfficeContext = {
            host: info.host,
            platform: info.platform,
            isReady: true,
            requirements: Office.context.requirements,
            version: Office.context.displayLanguage // Changed from host.version
          }

          // Update stores
          officeContext.set(context)
          isOfficeReady.set(true)

          this.initialized = true
          resolve()
        })
      })

      // Attempt auto-authentication
      await this.attemptAutoAuth()

    } catch (error) {
      console.error('Office.js initialization failed:', error)
      this.setError('Failed to initialize Office.js', error)
    } finally {
      isLoading.set(false)
    }
  }

  private waitForOffice(): Promise<void> {
    return new Promise((resolve, reject) => {
      let attempts = 0
      const maxAttempts = 50 // 5 seconds timeout

      const checkOffice = () => {
        if (typeof Office !== 'undefined' && typeof Office.onReady === 'function') {
          resolve()
        } else if (attempts >= maxAttempts) {
          reject(new Error('Office.js not available after timeout'))
        } else {
          attempts++
          setTimeout(checkOffice, 100)
        }
      }

      checkOffice()
    })
  }

  // Authentication methods
  async authenticateWithSSO(): Promise<UserInfo> {
    try {
      this.clearError()
      isLoading.set(true)

      if (!this.initialized) {
        await this.initialize()
      }

      // Check SSO support
      if (!this.isSsoSupported()) {
        throw new Error('SSO not supported in this Office version')
      }

      console.log('Attempting SSO authentication...')

      // Get access token from Office
      const token = await this.getAccessToken()

      // Validate with backend
      const userData = await this.validateTokenWithBackend(token)

      // Update stores
      currentUser.set(userData)
      localStorage.removeItem('userSignedOut')

      console.log('Authentication successful:', userData.user)
      return userData

    } catch (error) {
      console.error('SSO authentication failed:', error)
      this.handleAuthError(error)
      throw error
    } finally {
      isLoading.set(false)
    }
  }

  private async getAccessToken(): Promise<string> {
    return new Promise((resolve, reject) => {
      Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: false
      })
      .then(resolve)
      .catch(reject)
    })
  }

  private async validateTokenWithBackend(token: string): Promise<UserInfo> {
    const response = await fetch(`${this.config.backendUrl}/api/auth/microsoft`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      body: JSON.stringify({ token })
    })

    if (!response.ok) {
      const errorData = await response.json()
      throw new Error(errorData.error_description || 'Backend validation failed')
    }

    return response.json()
  }

  async signOut(): Promise<void> {
    currentUser.set(null)
    this.clearError()
    localStorage.setItem('userSignedOut', 'true')
    console.log('User signed out')
  }

  private async attemptAutoAuth(): Promise<void> {
    try {
      const userSignedOut = localStorage.getItem('userSignedOut')
      if (!userSignedOut && this.isSsoSupported()) {
        await this.authenticateWithSSO()
      }
    } catch (error) {
      // Silently fail auto-auth
      console.log('Auto-authentication not available:', error)
    }
  }

  // Office.js API methods
  async insertText(text: string, options?: InsertOptions): Promise<void> {
    if (!this.initialized) throw new Error('Office not ready')

    return new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(
        text,
        {
          coercionType: options?.coercionType || Office.CoercionType.Text,
        },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve()
          } else {
            reject(new Error(result.error?.message || 'Failed to insert text'))
          }
        }
      )
    })
  }

  async getSelectedText(): Promise<string> {
    if (!this.initialized) throw new Error('Office not ready')

    return new Promise((resolve, reject) => {
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value as string)
          } else {
            reject(new Error(result.error?.message || 'Failed to get selected text'))
          }
        }
      )
    })
  }

  async getDocumentProperties(): Promise<DocumentProperties> {
    if (!this.initialized) throw new Error('Office not ready')

    return new Promise((resolve, reject) => {
      Office.context.document.getFilePropertiesAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value as DocumentProperties)
        } else {
          reject(new Error(result.error?.message || 'Failed to get document properties'))
        }
      })
    })
  }

  async showNotification(message: string, type: 'info' | 'error' = 'info'): Promise<void> {
    if (!this.initialized) {
      console.log(`Notification (${type}): ${message}`)
      return
    }

    try {
      // Use Office notification API if available
      if (Office.context.ui) {
        Office.context.ui.displayDialogAsync(`https://appsforoffice.microsoft.com/fabric/dialog.html?_host_Info=HOST$notification$message=${message}&type=${type}`)
      } else {
        console.log(`Office Notification (${type}): ${message}`)
      }
    } catch (error) {
      console.log(`Notification (${type}): ${message}`)
    }
  }

  // Utility methods
  isSsoSupported(): boolean {
    return !!(Office.auth && Office.auth.getAccessToken)
  }

  isHostSupported(host: Office.HostType): boolean {
    const context = this.getCurrentContext()
    return context?.host === host
  }

  isRequirementSupported(requirement: string, version?: string): boolean {
    if (!this.initialized) return false
    return Office.context.requirements.isSetSupported(requirement, version)
  }

  getCurrentContext(): OfficeContext | null {
    let context: OfficeContext | null = null
    officeContext.subscribe(value => context = value)()
    return context
  }

  getCurrentUser(): UserInfo | null {
    let user: UserInfo | null = null
    currentUser.subscribe(value => user = value)()
    return user
  }

  // Error handling
  private handleAuthError(error: unknown): void {
    if (error instanceof Error) {
      const errorCode = this.extractErrorCode(error.message)
      const userMessage = this.getUserFriendlyErrorMessage(errorCode, error.message)
      this.setError(userMessage, errorCode)
    } else {
      this.setError('Unknown authentication error')
    }
  }

  private extractErrorCode(errorMessage: string): string {
    const codeMatch = errorMessage.match(/13\d{3}/)
    return codeMatch ? codeMatch[0] : '13000'
  }

  private getUserFriendlyErrorMessage(code: string, originalMessage: string): string {
    const errorMessages: Record<string, string> = {
      '13001': 'Sign-in was cancelled. Please try again.',
      '13002': 'Please sign in to Office 365 first.',
      '13003': 'Additional permissions required. Please contact your administrator.',
      '13004': 'Unable to get authentication token. Please try again.',
      '13005': 'Network error. Please check your connection.',
      '13006': 'Invalid credentials. Please sign out and try again.',
      '13000': 'Authentication failed. Please try again.'
    }

    return errorMessages[code] || originalMessage || 'Unknown error occurred'
  }

  private setError(message: string, error?: unknown): void {
    const authErrorObj: AuthError = {
      error: 'authentication_failed',
      error_description: message,
      timestamp: new Date().toISOString()
    }
    if (error instanceof Error) {
      authErrorObj.error_code = this.extractErrorCode(error.message)
    }

    authError.set(authErrorObj)
  }

  private clearError(): void {
    authError.set(null)
  }
}

// Export singleton instance
export const officeService = new OfficeService()

// Export convenient store getters
export function useOfficeReady(): Readable<boolean> {
  return isOfficeReady
}

export function useCurrentUser(): Writable<UserInfo | null> {
  return currentUser
}

export function useAuthError(): Writable<AuthError | null> {
  return authError
}

export function useOfficeContext(): Writable<OfficeContext | null> {
  return officeContext
}

export function useIsAuthenticated(): Readable<boolean> {
  return isAuthenticated
}

export function useHostType(): Readable<Office.HostType | null> {
  return hostType
}
