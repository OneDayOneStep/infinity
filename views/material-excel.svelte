<script>
  import {utils} from 'xlsx';
  import {onDestroy, setContext} from 'svelte';
  import {writable, derived} from 'svelte/store';
  import Checkbox from '@smui/checkbox';
  import Textfield from '@smui/textfield';
  //
  import MaterialNext from "./material-next.svelte";

  const excelList = writable([]);

  // const readWorker = new Worker("./read-excel-worker.js");
  // readWorker.onmessage = ev => {
  //   const excelData = ev.data;
  // 	initData(excelData);
  // }
  // fetch("./material_origin.xlsx")
  //   .then(res => res.arrayBuffer())
  //   .then(buffer => {
  //     readWorker.postMessage(buffer);
  //   })

  const initData = excelData => {
    console.log(excelData);
    const data = utils.sheet_to_json(excelData.Sheets[excelData.SheetNames[0]], {
      // header: 1, blankrows: true
    });
    let currentType = "normal";
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row.__EMPTY_1 && !row.__EMPTY_2 && !row.__EMPTY_3) {
        data.splice(i--, 1);
        continue;
      }
      if (typeof row.__EMPTY_3 === "string" && row.__EMPTY_3.includes("单价")) {
        if (row.__EMPTY_2.includes("耗材")) {
          currentType = "consumable";
        } else if (row.__EMPTY_2.includes("助理")) {
          currentType = "staff";
        } else {
          currentType = "normal";
        }
        row.isTitle = true;
        //
        row.goodsSize = "Num";
        row.priceCount = "PriceCount";
        row.control = "Control";
        //
        row.__EMPTY_3 = "Price";
      } else {
        row.goodsSize = 1;
        row.priceCount = row.__EMPTY_3 || 0;
      }
      row.rowType = currentType;
      row.symbolId = Symbol();
      row.isChecked = false;
      row.isFiltered = true;
      //
      row.__EMPTY_1 = row.__EMPTY_1 || "";
      row.__EMPTY_2 = row.__EMPTY_2 || "";
      row.__EMPTY_3 = row.__EMPTY_3 || "";
    }
    console.log(data);
    excelList.set(data);
  }
  initData(materialData);

  const clickRow = index => {
    excelList.update(arr => {
      arr[index].isChecked = !arr[index].isChecked;
      return arr;
    })
  }

  const selectedRows = derived(excelList, $excelList => {
    let size = 0;
    let priceCount = 0;
		let goodsPriceCount = 0;
    $excelList.forEach(row => {
      if (row.isChecked) {
        size++;
        priceCount += typeof row.__EMPTY_3 === "number" ? row.__EMPTY_3 : 0;
	      goodsPriceCount += row.__EMPTY_3 * row.goodsSize;
      }
    });
    return {
      size,
      priceCount,
	    goodsPriceCount
    }
  });

  const clearSelected = () => {
    $excelList.forEach(row => {
      row.isChecked = false;
    })
    excelList.set($excelList);
  }

  let searchText = writable("");
  let changeSearchConditionTimer = null;

  const filterList = () => {
    const text = $searchText.toUpperCase();
    let currentClassify = null;
    $excelList.forEach(row => {
      if (row.isTitle) {
        row.isFiltered = false;
        currentClassify = row;
      } else {
        row.isFiltered = `${row.__EMPTY_1}***${row.__EMPTY_2}`.toUpperCase().includes(text);
        if (row.isFiltered && currentClassify) {
          currentClassify.isFiltered = true;
        }
      }
    })
    excelList.set($excelList);
  }

  const unsubscribe_Search = searchText.subscribe(() => {
    clearTimeout(changeSearchConditionTimer);
    changeSearchConditionTimer = setTimeout(filterList, 500);
  });

  onDestroy(unsubscribe_Search);

  const nextVisible = writable(false);
  setContext("visible", nextVisible);
  const changeNextDisplay = status => {
    nextVisible.set(status);
  }

  // send to next page's data
  setContext("excelData", excelList);
  const updateExcelData = (rowId, propertyName, value) => {
    excelList.update(currentList => {
      const findIt = currentList.find(obj => obj.symbolId === rowId);
      findIt && (findIt[propertyName] = value);
      return currentList;
    });
  }

</script>

<MaterialNext changeNextDisplay={changeNextDisplay} updateExcelData={updateExcelData} />
<div class="material-excel">
  <!-- top -->
  <div class="material-excel-top">
    <Textfield
        style="width: 100%;"
        class="shaped-outlined"
        variant="outlined"
        label="SEARCH"
        bind:value={$searchText}
    >
<!--      <Icon class="material-icons" slot="leadingIcon">event</Icon>-->
    </Textfield>
  </div>
  <!-- list -->
  <div class="material-excel-list">
    {#each $excelList as { isFiltered, isChecked, isTitle, __EMPTY_1, __EMPTY_2, __EMPTY_3, goodsSize, priceCount, symbolId }, i}
      <div class:material-excel-row={true}
           class:material-excel-title={isTitle}
           class:material-excel-hidden={!isFiltered}
           class:material-excel-selected={isChecked}
           on:click={() => clickRow(i)} on:keydown={() => {}}>
        <div class="material-excel-item material-excel-item-check">
          {#if !isTitle}
            <Checkbox bind:checked={isChecked} />
          {/if}
        </div>
        <div class="material-excel-item material-excel-item-en">{ __EMPTY_1 }</div>
        <div class="material-excel-item material-excel-item-cn">{ __EMPTY_2 }</div>
        <div class="material-excel-item material-excel-item-price">{ __EMPTY_3 }</div>
        <div class="material-excel-item material-excel-item-num">
          {#if isTitle}
            { goodsSize }
          {:else}
            <input type="number" value={goodsSize} class="material-excel-item-num-input" on:change={ev => {
		            updateExcelData(symbolId, "goodsSize", ev.target.value);
              }} />
          {/if}
        </div>
        <div class="material-excel-item material-excel-item-price-total">
          {#if isTitle}
            { priceCount }
          {:else}
            { __EMPTY_3 * goodsSize }
          {/if}
        </div>
      </div>
    {/each}
  </div>
  <!-- bottom -->
  <div class="material-excel-bottom">
    <span style="width: 11em">
      <span class="material-excel-bottom-mini-font">SELECTED : </span>
      { $selectedRows.size }
      <span class="material-excel-bottom-mini-font"
            style="opacity: { $selectedRows.size > 0 ? 1 : 0 }"
            on:click={clearSelected}>
         ( EMPTY )
      </span>
    </span>

    <span style="width: 11em">
      <span class="material-excel-bottom-mini-font">BASE PRICE : </span>
      <span>{ $selectedRows.priceCount }</span>
    </span>

    <span class="material-excel-bottom-mini-font">PRICE TOTAL : </span>
    <span>{ $selectedRows.goodsPriceCount }</span>
    <div class="material-excel-detail"
         style="opacity: { $selectedRows.size > 0 ? 1 : 0 }"
         on:click={() => changeNextDisplay(true)}>NEXT</div>
  </div>
</div>


<style lang="scss">
  // parent
  .material-excel {
    height: 100%;
    display: flex;
    flex-direction: column;
  }
  // top
  .material-excel-top {
    margin: 1rem;
    user-select: none;
  }
  // material common
  @import "material";
</style>
