<script>
import Navbar from "../components/Navbar.svelte";
  /** @type {{children?: import('svelte').Snippet}} */
  let { children } = $props();
</script>


<body class="bg-gray-50 w-screen">
  <div class="flex flex-col max-w-[480px] m-auto">
    <Navbar />
    <main class="border-x border-gray-200">
      {@render children()}
    </main>
  </div>
</body>

