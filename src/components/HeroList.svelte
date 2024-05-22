<script lang="ts">
const listItems = [
	{
		icon: "Ribbon",
		primaryText: "Achieve more with Office integration",
	},
	{
		icon: "Unlock",
		primaryText: "Unlock features and functionality",
	},
	{
		icon: "Design",
		primaryText: "Create and visualize like a pro",
	},
];
</script>

<div class="flex flex-col gap-y-6 items-center justify-center">
  <div>
    <img width="120" height="120" src="assets/logo-filled.png" alt="" />
  </div>
  <ul class="list-disc list-inside">
    {#each listItems as item}
      <li class="">
        {item.primaryText}
      </li>
    {/each}
    </ul>
</div>

