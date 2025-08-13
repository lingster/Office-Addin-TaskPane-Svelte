/// <reference types="@microsoft/office-js" />

interface AuthConfig {
  clientId: string
  backendUrl: string
  scope: string
}

interface UserInfo {
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
      // Show loading state
      this.showLoadingState()

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

      // Update UI with user info
      this.updateUIWithUser(userData)

      return userData

    } catch (error) {
      console.error("SSO authentication failed:", error)

      // Handle specific error cases
      if (error instanceof Error) {
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
      }

      // Fallback to manual login
      await this.fallbackToManualLogin()
      throw error
    } finally {
      this.hideLoadingState()
    }
  }

  private showLoadingState(): void {
    const button = document.getElementById("auth-button")
    if (button) {
      button.textContent = "Signing in..."
      button.setAttribute("disabled", "true")
    }
  }

  private hideLoadingState(): void {
    const button = document.getElementById("auth-button")
    if (button) {
      button.textContent = "Sign in"
      button.removeAttribute("disabled")
    }
  }

  private updateUIWithUser(user: UserInfo): void {
    const userDisplay = document.getElementById("user-display")
    if (userDisplay) {
      userDisplay.innerHTML = `
        <div class="user-info">
          <div class="user-name">${user.user || 'Unknown User'}</div>
          <div class="user-email">${user.email || 'No email available'}</div>
        </div>
      `
    }

    // Hide sign-in button, show user menu
    const authButton = document.getElementById("auth-button")
    const userMenu = document.getElementById("user-menu")

    if (authButton) authButton.style.display = "none"
    if (userMenu) userMenu.style.display = "block"
  }

  private async fallbackToManualLogin(): Promise<void> {
    // Implement manual login flow if SSO fails
    console.log("Falling back to manual login...")
    // This could redirect to a popup or alternative auth flow
  }

  getCurrentUser(): UserInfo | null {
    return this.currentUser
  }

  async signOut(): Promise<void> {
    this.currentUser = null

    // Clear UI
    const userDisplay = document.getElementById("user-display")
    const authButton = document.getElementById("auth-button")
    const userMenu = document.getElementById("user-menu")

    if (userDisplay) userDisplay.innerHTML = ""
    if (authButton) authButton.style.display = "block"
    if (userMenu) userMenu.style.display = "none"

    console.log("User signed out")
  }
}

// Initialize authentication
const authConfig: AuthConfig = {
  clientId: "YOUR-AZURE-CLIENT-ID",
  backendUrl: "https://athena-app.imperialai.ai",
  scope: "api://YOUR-AZURE-CLIENT-ID/access_as_user"
}

export const authService = new AuthService(authConfig)
