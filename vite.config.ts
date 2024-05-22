import { svelte } from "@sveltejs/vite-plugin-svelte";
import * as devCerts from "office-addin-dev-certs";
import { defineConfig } from "vite";
import { createHtmlPlugin } from "vite-plugin-html";

async function getHttpsOptions() {
	const httpsOptions = await devCerts.getHttpsServerOptions();
	return {
		ca: httpsOptions.ca,
		key: httpsOptions.key,
		cert: httpsOptions.cert,
	};
}

// https://vitejs.dev/config/
export default defineConfig({
	plugins: [
		svelte(),
		createHtmlPlugin({
			minify: true,
			pages: [
				{
					entry: "src/main.ts",
					filename: "taskpane.html",
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
	server: {
		https: await getHttpsOptions(),
		port: Number.parseInt(process.env.server_port) || 3000,
		strictPort: false,
		open: "/taskpane.html", // opens the correct /taskpane.html when opening browser to view in web
	},
	preview: {
		https: await getHttpsOptions(),
		port: Number.parseInt(process.env.preview_server_port) || 3000,
		strictPort: false,
		open: "/taskpane.html", // opens the correct /taskpane.html when opening browser to view in web
	},
});
