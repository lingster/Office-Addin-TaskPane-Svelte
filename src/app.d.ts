// Global type declarations for Svelte app
declare global {
  namespace App {
    interface Error {
      code?: string
      details?: unknown
    }

    interface Locals {
      user?: {
        id: string
        email: string
        name: string
      }
    }

    interface PageData {
      user?: App.Locals['user']
    }

    // Add platform-specific interfaces if needed
    interface Platform {}
  }

  // Office.js global declarations
  interface Window {
    Office: typeof Office
    OfficeRuntime: typeof OfficeRuntime
  }

  // Environment variables with full typing
  interface ImportMetaEnv {
    readonly VITE_CLIENT_ID: string
    readonly VITE_TENANT_ID: string
    readonly VITE_API_SCOPE: string
    readonly VITE_BACKEND_URL: string
    readonly VITE_ENVIRONMENT: 'development' | 'staging' | 'production'
    readonly VITE_DEBUG: string
  }

  interface ImportMeta {
    readonly env: ImportMetaEnv
  }
}

// Office.js module declarations
declare module 'office-js' {
  export = Office
}

// Ensure this file is treated as a module
export {}
