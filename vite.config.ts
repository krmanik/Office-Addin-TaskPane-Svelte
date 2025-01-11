import { defineConfig } from 'vite'
import { svelte } from '@sveltejs/vite-plugin-svelte'
import * as devCerts from "office-addin-dev-certs";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return {
    ca: httpsOptions.ca,
    key: httpsOptions.key,
    cert: httpsOptions.cert,
  };
}

// https://vite.dev/config/
export default defineConfig({
  plugins: [
    svelte()
  ],
  server: {
    https: await getHttpsOptions(),
    port: Number.parseInt(process.env.server_port) || 3000,
    strictPort: false,
    open: "/", // opens the correct /taskpane.html when opening browser to view in web
  },
  preview: {
    https: await getHttpsOptions(),
    port: Number.parseInt(process.env.preview_server_port) || 3000,
    strictPort: false,
    open: "/taskpane.html", // opens the correct /taskpane.html when opening browser to view in web
  },
})
