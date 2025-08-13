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
  function updateDOM(authResult: 'success' | 'fail' | 'pending' | 'initial' = 'initial'): void {
    const nameBox = document.getElementById("user-name-box");
    const emailBox = document.getElementById("user-email-box");
    const statusIcon = document.getElementById("status-icon");
    const errorContainer = document.getElementById("error-message");

    // Update textboxes
    if (nameBox) nameBox.textContent = currentUser?.user || "";
    if (emailBox) emailBox.textContent = currentUser?.email || "";

    // Update status icon
    if (statusIcon) {
      statusIcon.classList.remove('bg-gray-500', 'bg-green-500', 'bg-red-500', 'bg-yellow-500');
      if (authResult === 'success') {
        statusIcon.classList.add('bg-green-500');
      } else if (authResult === 'fail') {
        statusIcon.classList.add('bg-red-500');
      } else if (authResult === 'pending') {
         statusIcon.classList.add('bg-yellow-500');
      } else {
        statusIcon.classList.add('bg-gray-500');
      }
    }

    // Update error message
    if (errorContainer) {
      errorContainer.textContent = errorMessage;
      errorContainer.style.display = errorMessage ? "block" : "none";
    }
  }

  async function handleTestAuth(): Promise<void> {
    isLoading = true;
    errorMessage = "";
    currentUser = null;
    updateDOM('pending');

    try {
      currentUser = await authService.authenticateWithSSO();
      updateDOM('success');
    } catch (error) {
      console.error("Authentication failed:", error);
      errorMessage = error instanceof Error ? error.message : "An unknown error occurred.";
      updateDOM('fail');
      setTimeout(() => {
        errorMessage = "";
        if (errorContainer) errorContainer.style.display = "none";
      }, 5000);
    } finally {
      isLoading = false;
    }
  }

  onMount(async () => {
    await Office.onReady();
    isOfficeInitialized = true;

    // Wire up event listener for the new test button
    document.getElementById("test-sso-button")?.addEventListener("click", handleTestAuth);

    // Initial UI state
    updateDOM('initial');
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