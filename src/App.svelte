<script lang="ts">
  import {type RouterConf} from 'svelte-mini-router';
  import {Router} from 'svelte-mini-router';
  import Navbar from './components/Navbar.svelte';

  export const routerConf: RouterConf = {
      routes: [
          // this is your home page
          {path: '/', render: () => import('./routes/Home.svelte')},

          // login page
          {path: '/login', render: () => import('./routes/Login.svelte')},

          // taskpane
          {path: '/taskpane.html', render: () => import('./routes/Home.svelte')},

          // commands
          {path: '/commands.html', render: () => import('./routes/Command.svelte')},
      ],
  };
</script>

<div class="flex flex-col max-w-[480px] m-auto">
  <Navbar />
  <main class="border-x border-gray-200">
    <Router {routerConf} />
  </main>
</div>
