<script lang="ts">
  import { onMount } from 'svelte'
  import {
    officeService,
    useOfficeReady,
    useCurrentUser,
    useAuthError,
    useIsAuthenticated
  } from '../lib/office'
  import HeroList from "../components/HeroList.svelte";

  // Reactive stores
  const isReady = useOfficeReady()
  const currentUser = useCurrentUser()
  const authError = useAuthError()
  const isAuthenticated = useIsAuthenticated()

  // Component state
  let isAuthenticating = false

  // Handlers
  async function handleAuthenticate() {
    if (isAuthenticating) return

    isAuthenticating = true
    try {
      await officeService.authenticateWithSSO()
    } catch (error) {
      console.error('Authentication failed:', error)
    } finally {
      isAuthenticating = false
    }
  }

  async function handleSignOut() {
    await officeService.signOut()
  }

  const click = async () => {
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph(
        "Hello World",
        Word.InsertLocation.end,
      );

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      await context.sync();
    });
  };

  onMount(() => {
    console.log('Home component mounted')
  })
</script>

<main class="flex flex-col items-center justify-center h-screen">
  {#if !$isReady}
    <div class="status-indicator loading">
      <span>Initializing Office.js...</span>
    </div>
  {:else if $isAuthenticated && $currentUser}
    <!-- Authenticated state -->
    <div class="auth-success">
      <div class="user-info">
        <div class="user-details">
          <div class="user-name">{$currentUser.user || 'Unknown User'}</div>
          <div class="user-email">{$currentUser.email || 'No email available'}</div>
        </div>
      </div>

      <button
        class="btn btn-outline"
        on:click={handleSignOut}
        type="button"
      >
        Sign Out
      </button>
    </div>
    <HeroList />
    <div class="run-button">
        <fluent-button appearance="accent" onclick={click}>Run</fluent-button>
    </div>
  {:else}
    <!-- Unauthenticated state -->
    <div class="auth-prompt">
      {#if $authError}
        <div class="error-message">
          <div>
            <h4>Authentication Error</h4>
            <p>{$authError.error_description}</p>
            {#if $authError.error_code}
              <small>Error Code: {$authError.error_code}</small>
            {/if}
          </div>
        </div>
      {/if}

      <button
        class="btn btn-primary"
        on:click={handleAuthenticate}
        disabled={isAuthenticating}
        type="button"
      >
        {#if isAuthenticating}
          Signing in...
        {:else}
          Test
        {/if}
      </button>
    </div>
    <HeroList />
    <div class="run-button">
        <fluent-button appearance="accent" onclick={click}>Run</fluent-button>
    </div>
  {/if}
</main>

<style>
  .auth-success {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 1rem;
    margin-bottom: 1rem;
  }

  .user-info {
    display: flex;
    align-items: center;
    gap: 0.75rem;
  }

  .user-details {
    display: flex;
    flex-direction: column;
  }

  .user-name {
    font-weight: 600;
  }

  .user-email {
    font-size: 0.875rem;
  }

  .auth-prompt {
    text-align: center;
  }

  .error-message {
    display: flex;
    align-items: flex-start;
    gap: 0.75rem;
    padding: 1rem;
    background: #fef2f2;
    border: 1px solid #fecaca;
    border-radius: 0.375rem;
    margin-bottom: 1rem;
    text-align: left;
  }

  .error-message h4 {
    margin: 0 0 0.25rem 0;
    font-weight: 600;
    color: #dc2626;
  }

  .error-message p {
    margin: 0;
    color: #7f1d1d;
  }

  .error-message small {
    color: #991b1b;
    font-size: 0.75rem;
  }

  .btn {
    padding: 0.625rem 1.25rem;
    border-radius: 0.375rem;
    font-weight: 500;
    border: 1px solid transparent;
    cursor: pointer;
    transition: all 0.2s;
    display: inline-flex;
    align-items: center;
    gap: 0.5rem;
  }

  .btn:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }

  .btn-primary {
    background: #3b82f6;
    color: white;
  }

  .btn-primary:hover:not(:disabled) {
    background: #2563eb;
  }

  .btn-outline {
    background: transparent;
    color: #6b7280;
    border-color: #d1d5db;
  }

  .btn-outline:hover:not(:disabled) {
    background: #f9fafb;
    border-color: #9ca3af;
  }
</style>