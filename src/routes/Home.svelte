<script lang="ts">
  import { onMount } from "svelte";
  import { authService, type UserInfo } from "../auth";
  import {
    provideFluentDesignSystem,
    allComponents,
  } from "@fluentui/web-components";

  provideFluentDesignSystem().register(allComponents);

  let isOfficeInitialized = false;
  let currentUser: UserInfo | null = null;
  let isLoading = true;
  let errorMessage = "";

  // This function is the bridge between Svelte and the static elements in taskpane.html
  function updateDOM(): void {
    const authContainer = document.getElementById("auth-container");
    const userContainer = document.getElementById("user-container");
    const userDisplay = document.getElementById("user-display");
    const errorContainer = document.getElementById("error-message");

    if (isLoading) {
      // You might want a loading indicator
    }

    if (currentUser) {
      if (authContainer) authContainer.style.display = "none";
      if (userContainer) userContainer.style.display = "block";
      if (userDisplay) {
        userDisplay.innerHTML = `
          <div class="user-info">
            <div class="ms-fontSize-xl ms-fontWeight-semibold">${currentUser.user || 'Unknown User'}</div>
            <div class="ms-fontSize-m">${currentUser.email || 'No email available'}</div>
          </div>
        `;
      }
    } else {
      if (authContainer) authContainer.style.display = "block";
      if (userContainer) userContainer.style.display = "none";
    }

    if (errorContainer) {
      errorContainer.textContent = errorMessage;
      errorContainer.style.display = errorMessage ? "block" : "none";
    }
  }

  async function handleAuth(): Promise<void> {
    isLoading = true;
    errorMessage = "";
    updateDOM();
    try {
      currentUser = await authService.authenticateWithSSO();
    } catch (error) {
      console.error("Authentication failed:", error);
      errorMessage = error instanceof Error ? error.message : "An unknown error occurred.";
      setTimeout(() => { errorMessage = "" }, 5000);
    } finally {
      isLoading = false;
      updateDOM();
    }
  }

  async function handleSignOut(): Promise<void> {
    await authService.signOut();
    currentUser = null;
    updateDOM();
  }

  async function attemptAutoAuth(): Promise<void> {
    try {
      currentUser = await authService.authenticateWithSSO();
    } catch (error) {
      console.log("Auto-authentication not available:", error);
    }
    isLoading = false;
    updateDOM();
  }

  onMount(async () => {
    await Office.onReady();
    isOfficeInitialized = true;

    // Wire up event listeners
    document.getElementById("auth-button")?.addEventListener("click", handleAuth);
    document.getElementById("sign-out-button")?.addEventListener("click", handleSignOut);

    await attemptAutoAuth();
  });
</script>

<!-- This Svelte component is now mostly for logic. The UI is in taskpane.html -->
<!-- We can still have Svelte-controlled elements here if we want. -->

<style>
  :global(.user-info) {
    padding: 10px;
    border: 1px solid #ccc;
    border-radius: 4px;
    background-color: #f9f9f9;
    text-align: left;
  }
</style>