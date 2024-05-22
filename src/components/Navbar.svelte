<script lang="ts">
import { link } from "svelte-spa-router";
import { location } from "svelte-spa-router";
</script>

<div class="flex justify-around items-center space-x-4 border-b border-gray-200 py-2">
    <div>
      <img width="35" height="35" src="assets/logo-filled.png" alt="logo" />
    </div>
    <div>
        <p>Path: {$location}</p>
    </div>
    <ul class="flex space-x-4">
        <li><a href="#/">Home</a></li>
        <li><a href="#/login">Login</a></li>
    </ul>
</div>

