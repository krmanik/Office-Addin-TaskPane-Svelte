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

<img width="120" height="120" src="assets/logo-filled.png" alt="" />
<div class="list-container">
  {#each listItems as item}
    <div class="list-item">
      <i class={`icon-padding ms-Icon ms-Icon--${item.icon}`} />
      {item.primaryText}
    </div>
  {/each}
</div>

<style>
  .icon-padding {
    padding-left: 8px;
    padding-right: 8px;
  }

  .list-item {
    text-align: left;
    font-size: 16px;
  }
</style>
