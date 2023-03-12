<script lang="ts">
  import Progress from "./components/Progress.svelte";
  import HeroList from "./components/HeroList.svelte";

  import {
    allComponents,
    provideFluentDesignSystem,
  } from "@fluentui/web-components";
  provideFluentDesignSystem().register(allComponents);

  let isOfficeInitialized = false;
  window.onload = function () {
    const Office = window.Office;
    Office.onReady(() => {
      console.log("Office Ready");
      isOfficeInitialized = true;
    });
  };

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

<svelte:head>
  <!-- Office JavaScript API -->
  <script
    type="text/javascript"
    src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"
  ></script>
  <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->
  <link
    rel="stylesheet"
    href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/11.0.0/css/fabric.min.css"
  />
</svelte:head>

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
