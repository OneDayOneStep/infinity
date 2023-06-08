<script>
	import { getContext } from 'svelte';
	import { fade, fly } from 'svelte/transition';

	let visible = getContext("visible");
	export let nextStepDisplay;

</script>

{#if $visible}
  <div class="material-next-mask" transition:fade="{{ duration: 400 }}" on:dblclick={() => nextStepDisplay(false)}></div>
  <div class="material-next" transition:fly="{{ x: '100%', duration: 350 }}">


  </div>
{/if}

<style lang="scss">
  .material-next-mask {
    background-color: rgba(0, 0, 0, 0.5);
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    cursor: pointer;
    z-index: 499;
  }
  .material-next {
    position: fixed;
    top: 0;
    bottom: 0;
    right: 0;
    width: 80vw;
    z-index: 500;
    background-color: #FFF;
    box-shadow: 0 0 0.5rem #555;
  }
</style>