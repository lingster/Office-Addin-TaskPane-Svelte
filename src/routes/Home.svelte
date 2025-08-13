/// <reference types="@microsoft/office-js" />

<script lang="ts">
import {
	allComponents,
	provideFluentDesignSystem,
} from "@fluentui/web-components";
import { onMount } from "svelte";
provideFluentDesignSystem().register(allComponents);
import HeroList from "../components/HeroList.svelte";
import Navbar from "../components/Navbar.svelte";
import Progress from "../components/Progress.svelte";
import { authService } from "../auth";

let isOfficeInitialized = false;
onMount(async () => {
	const Office = window.Office;
	await Office.onReady();
	isOfficeInitialized = true;
});

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

const handleAuth = async () => {
  try {
    await authService.authenticateWithSSO()
    console.log("Authentication successful")
  } catch (error) {
    console.error("Authentication failed:", error)
    showErrorMessage(error instanceof Error ? error.message : "Authentication failed")
  }
}

function showErrorMessage(message: string): void {
  const errorDiv = document.getElementById("error-message")
  if (errorDiv) {
    errorDiv.textContent = message
    errorDiv.style.display = "block"

    // Hide error after 5 seconds
    setTimeout(() => {
      errorDiv.style.display = "none"
    }, 5000)
  }
}
</script>
    
{#if !isOfficeInitialized}
  <Progress
    title="Contoso Task Pane Add-in"
    message="Please sideload your addin to see app body."
  />
{:else}
  <main class="flex flex-col items-center justify-center h-screen">
    <div id="user-display"></div>
    <div id="error-message" style="display: none; color: red; margin-bottom: 10px;"></div>
    <HeroList />
    <div>
      <div class="text-blue-500 mt-4">
        Modify the source files, then click <b>Run</b>.
      </div>
    </div>
    <div class="run-button">
        <fluent-button appearance="accent" onclick={click}>Run</fluent-button>
        <fluent-button appearance="accent" id="auth-button" onclick={handleAuth}>Test</fluent-button>
    </div>
  </main>
{/if}

<style>
  :global(.run-button) {
    margin: 20px !important;
    text-align: center;
  }

  :global(body) {
    background-color: var(--fds-solid-background-base);
    color: var(--fds-text-primary);
  }
</style>