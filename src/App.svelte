<script lang="ts">
  import Progress from "./components/Progress.svelte";
  import HeroList from "./components/HeroList.svelte";

  import {
    allComponents,
    provideFluentDesignSystem,
  } from "@fluentui/web-components";
  import { onMount } from "svelte";
  provideFluentDesignSystem().register(allComponents);

  let isOfficeInitialized = false;
  onMount(async () => {
    const Office = window.Office;
    await Office.onReady();
    isOfficeInitialized = true;
  })

  const click = async () => {
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph(
        "Hello World",
        Word.InsertLocation.end
      );

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      await context.sync();
    });
  };
</script>

{#if !isOfficeInitialized}
  <Progress
    title="Contoso Task Pane Add-in"
    message="Please sideload your addin to see app body."
  />
{:else}
  <main>
    <HeroList />
    <div style="margin-top: 20px; font-size: 18px;">
      <div>
        Modify the source files, then click <b>Run</b>.
      </div>
    </div>

    <div class="run-button">
      <fluent-button appearance="accent" onclick={click}>Run</fluent-button>
    </div>
  </main>
{/if}

<style>
  :global(.run-button) {
    margin: 20px !important;
    text-align: center;
  }

  :global(body) {
    background-color: var(--fds-solid-background-base);
    color: var(--fds-text-primary);
  }
</style>
