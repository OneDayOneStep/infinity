<script>
	import { getContext } from 'svelte';
	import { fade, fly } from 'svelte/transition';
	import { derived } from 'svelte/store';
  import exportExcel from "../exportExcel";

	let visible = getContext("visible");
	export let changeNextDisplay;
	let excelData = getContext("excelData");
	let selectedExcelData = derived(excelData, $excelData => {
		const filterData = [];
		let waitTitleRow = null;
		for (let i = 0;i < $excelData.length;i++) {
			const row = $excelData[i];
			if (row.isTitle) {
				waitTitleRow = row;
				continue;
			}
			if (row.isChecked) {
				if (waitTitleRow) {
					filterData.push(waitTitleRow);
					waitTitleRow = null;
        }
				filterData.push(row);
      }
    }
		return filterData;
	});
	const goodsCount = derived(selectedExcelData, $selectedExcelData => {
		const count = {
			price: 0,
      size: 0
    }
		$selectedExcelData.forEach(row => {
			if (!row.isTitle) {
				count.size++;
				count.price += row.__EMPTY_3 * row.goodsSize;
      }
		});
		return count;
	});
	export let updateExcelData;

	const unSelect = row => {
		updateExcelData(row.symbolId, "isChecked", false);
  }

	const startExport = () => {
		exportExcel($selectedExcelData);
		// window.requestAnimationFrame(() => changeNextDisplay(false));
  }

</script>

{#if $visible}
  <div class="material-next-mask" transition:fade="{{ duration: 400 }}" on:click={() => changeNextDisplay(false)}></div>
  <div class="material-next" transition:fly="{{ x: '100%', duration: 350 }}">

    <div class="material-excel-list" style="margin-top: 1rem;">
      {#each $selectedExcelData as row, i}
        <div class:material-excel-row={true}
             class:material-excel-title={row.isTitle}
             class:material-excel-hidden={!row.isFiltered}>
          <div class="material-excel-item material-excel-item-en">{ row.__EMPTY_1 }</div>
          <div class="material-excel-item material-excel-item-cn">{ row.__EMPTY_2 }</div>
          <div class="material-excel-item material-excel-item-price">{ row.__EMPTY_3 }</div>
          <div class="material-excel-item material-excel-item-num">
            {#if row.isTitle}
              { row.goodsSize }
            {:else}
              <input type="number" value={row.goodsSize} class="material-excel-item-num-input" on:change={ev => {
		            updateExcelData(row.symbolId, "goodsSize", ev.target.value);
              }} />
            {/if}
          </div>
          <div class="material-excel-item material-excel-item-price-total">
            {#if row.isTitle}
              { row.priceCount }
            {:else}
              { row.__EMPTY_3 * row.goodsSize }
            {/if}
          </div>
          <div class="material-excel-item material-excel-item-control">
            {#if row.isTitle}
              { row.control }
            {:else}
              <span class="material-excel-item-del" on:click={() => unSelect(row)}>Del</span>
            {/if}
          </div>
        </div>
      {/each}
    </div>
    <!-- bottom -->
    <div class="material-excel-bottom">
      <span style="width: 11em">
        <span class="material-excel-bottom-mini-font">GOODS NUM : </span>
        { $goodsCount.size }
      </span>
      <span>
      <span class="material-excel-bottom-mini-font">PRICE TOTAL : </span>
        { $goodsCount.price }
      </span>
      {#if $selectedExcelData.length > 0}
        <div class="material-excel-detail" on:click="{startExport}">EXPORT</div>
      {/if}
    </div>

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
    width: 85vw;
    z-index: 500;
    background-color: #FFF;
    box-shadow: 0 0 0.5rem #555;
    display: flex;
    flex-direction: column;
    .material-excel-item-num-input {
      text-align: center;
      width: 100%;
      border: none;
    }
    .material-excel-item-num-input::-webkit-inner-spin-button,
    .material-excel-item-num-input::-webkit-outer-spin-button {
      -webkit-appearance: none;
      margin: 0;
    }
  }
  // material common
  @import "material";
  //
  .material-excel-row:active {
    transform: none;
  }
</style>