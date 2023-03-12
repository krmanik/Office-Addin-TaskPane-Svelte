<script lang="ts">
  export let message: string;
  export let title: string;
</script>

<section class="ms-welcome__progress ms-u-fadeIn500">
  <div class="progress-item">{title}</div>
  <div class="progress-item"><fluent-progress-ring style="margin: 0 auto;"/></div>
  <div class="progress-item">{message}</div>
</section>

<style>
  :global(.progress-item) {
    padding: 14px;
    display: block;
    text-align: center;
  }
</style>
