<script lang="ts">
  interface Props {
    message: string;
    title: string;
  }

  let { message, title }: Props = $props();
</script>

<section class="ms-welcome__progress ms-u-fadeIn500">
  <div class="progress-item">{title}</div>
  <div class="progress-item"><fluent-progress-ring style="margin: 0 auto;"></fluent-progress-ring></div>
  <div class="progress-item">{message}</div>
</section>

<style>
  :global(.progress-item) {
    padding: 14px;
    display: block;
    text-align: center;
  }
</style>
