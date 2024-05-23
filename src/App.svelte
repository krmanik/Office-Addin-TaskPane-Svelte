<script lang="ts">
import Router from "svelte-spa-router";
import AppLayout from "./layouts/AppLayout.svelte";
import Home from "./routes/Home.svelte";
import Login from "./routes/Login.svelte";

const routes = {
	// Exact path
	"/": Home,

	// Using named parameters, with last being optional
	"/login": Login,

	// Wildcard parameter
	// '/book/*': Book,

	// Catch-all
	// This is optional, but if present it must be the last
	// '*': NotFound,
};
</script>

<AppLayout>
	<Router {routes}/>
</AppLayout>

