declare const OfficeRuntime: any;

interface AuthConfig {
  clientId: string
  backendUrl: string
  scope: string
}

export interface UserInfo {
  user: string | null
  email: string | null
  oid: string
  tenant: string
  authenticated_at: string
}

interface AuthError {
  error: string
  error_description: string
  timestamp: string
}

class AuthService {
  private config: AuthConfig
  private currentUser: UserInfo | null = null

  constructor(config: AuthConfig) {
    this.config = config
  }

  async authenticateWithSSO(): Promise<UserInfo> {
    try {
      // Attempt to get SSO token
      const token = await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: false
      })

      console.log("Successfully obtained SSO token")

      // Send token to backend for validation
      const response = await fetch(`${this.config.backendUrl}/api/auth/microsoft`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Accept": "application/json"
        },
        body: JSON.stringify({ token }),
      })

      if (!response.ok) {
        const errorData: AuthError = await response.json()
        throw new Error(`Authentication failed: ${errorData.error_description}`)
      }

      const userData: UserInfo = await response.json()
      this.currentUser = userData

      console.log("User authenticated successfully:", userData.user)

      return userData

    } catch (error) {
      console.error("SSO authentication failed:", error)

      // Handle specific error cases
      if (error.message.includes("13001")) {
        // User cancelled sign-in
        throw new Error("Sign-in was cancelled. Please try again.")
      } else if (error.message.includes("13002")) {
        // User not signed in
        throw new Error("Please sign in to Office 365 first.")
      } else if (error.message.includes("13003")) {
        // User consent required
        throw new Error("Additional permissions required. Please contact your administrator.")
      }

      // Re-throw the error to be handled by the caller
      throw error
    }
  }

  getCurrentUser(): UserInfo | null {
    return this.currentUser
  }

  async signOut(): Promise<void> {
    this.currentUser = null
    // In a real app, you might also want to clear any locally stored tokens or session info.
    localStorage.removeItem("userSignedIn"); // Example
    console.log("User signed out")
  }
}

// Initialize authentication
// Note: Vite uses `import.meta.env` instead of `process.env`
const authConfig: AuthConfig = {
  clientId: import.meta.env.VITE_CLIENT_ID!,
  backendUrl: import.meta.env.VITE_BACKEND_URL!,
  scope: import.meta.env.VITE_API_SCOPE!
}

export const authService = new AuthService(authConfig)
