import App from "./App.svelte";
import "./app.css";

const target = document.getElementById("app");

if (!target) {
  throw new Error("Could not find app target");
}

const app = new App({
	target,
});

export default app;
