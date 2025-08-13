import { resolve } from "node:path";
import { svelte } from "@sveltejs/vite-plugin-svelte";
import { defineConfig } from "vite";
import { createHtmlPlugin } from "vite-plugin-html";

// https://vitejs.dev/config/
export default defineConfig({
	plugins: [
		svelte(),
		createHtmlPlugin({
			minify: true,
			pages: [
				{
					entry: "src/main.ts",
					filename: "index.html", // updated this to index.html now we serve the taskpane.html from https:localhost:3000/
					template: "taskpane.html",
					injectOptions: {
						data: {
							injectScript: `<script src="./main.js"></script>`,
						},
					},
				},
				{
					entry: "src/commands.ts",
					filename: "commands.html",
					template: "commands.html",
					injectOptions: {
						data: {
							injectScript: `<script src="./commands.js"></script>`,
						},
					},
				},
			],
		}),
	],
});
