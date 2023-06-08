<script>
  import {utils} from 'xlsx';
  import {onDestroy} from 'svelte';
  import {writable, derived} from 'svelte/store';
  import Checkbox from '@smui/checkbox';
  import Textfield from '@smui/textfield';

  const readWorker = new Worker("./read-excel-worker.js");

  const excelList = writable([]);
  fetch("./material_origin.xlsx")
    .then(res => res.arrayBuffer())
    .then(buffer => {
      readWorker.postMessage(buffer);
    })

  readWorker.onmessage = ev => {
    const excelData = ev.data;
    const data = utils.sheet_to_json(excelData.Sheets[excelData.SheetNames[0]], {
      // header: 1,
      // blankrows: true
    });
    console.log(data);
    data.forEach(row => {
      row.isChecked = false;
      row.isFiltered = true;
      if (typeof row.__EMPTY_3 === "string" && row.__EMPTY_3.includes("单价")) {
        row.isTitle = true;
        row.__EMPTY_3 = "Price";
      }
      row.__EMPTY_1 = row.__EMPTY_1 || "";
      row.__EMPTY_2 = row.__EMPTY_2 || "";
      row.__EMPTY_3 = row.__EMPTY_3 || "";
    })
    excelList.set(data);
  }

  const clickRow = index => {
    excelList.update(arr => {
      arr[index].isChecked = !arr[index].isChecked;
      return arr;
    })
  }

  const selectedRows = derived(excelList, $excelList => {
    let size = 0;
    let priceCount = 0;
    $excelList.forEach(row => {
      if (row.isChecked) {
        size++;
        priceCount += typeof row.__EMPTY_3 === "number" ? row.__EMPTY_3 : 0;
      }
    });
    return {
      size,
      priceCount
    }
  });

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
</script>

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
    {#each $excelList as { isFiltered, isChecked, isTitle, __EMPTY_1, __EMPTY_2, __EMPTY_3 }, i}
      <div class:material-excel-row={true}
           class:material-excel-title={isTitle}
           class:material-excel-hidden={!isFiltered}
           on:click={() => clickRow(i)} on:keydown={() => {}}>
        <div class="material-excel-item material-excel-item-check">
          {#if !isTitle}
            <Checkbox bind:checked={isChecked} />
          {/if}
        </div>
        <div class="material-excel-item material-excel-item-en">{ __EMPTY_1 }</div>
        <div class="material-excel-item material-excel-item-cn">{ __EMPTY_2 }</div>
        <div class="material-excel-item material-excel-item-price">{ __EMPTY_3 }</div>
      </div>
    {/each}
  </div>
  <!-- bottom -->
  <div class="material-excel-bottom">
    <span style="margin-right: 3em;">
      <span style="font-size: 0.7em;font-weight: 600;">{ $searchText }SELECTED : </span>
      { $selectedRows.size }
    </span>
    <span>
      <span style="font-size: 0.7em;font-weight: 600;">BASE PRICE : </span>
      { $selectedRows.priceCount }
    </span>
    <div class="material-excel-detail">NEXT</div>
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
  }
  // bottom
  .material-excel-bottom {
    box-shadow: 0 0 0.35rem #AAA;
    position: relative;
    z-index: 50;
    height: 2em;
    display: flex;
    align-items: center;
    padding-left: 1.5em;
    font-size: 1.5em;
    margin-top: 0.35rem;
    .material-excel-detail {
      height: 100%;
      background-color: #58BE6B;
      color: #FFF;
      cursor: pointer;
      margin-left: auto;
      display: flex;
      align-items: center;
      padding: 0 2rem;
    }
  }
  // list
  .material-excel-list {
    flex: 1;
    height: 50vh;
    margin-left: 1rem;
    padding-right: 1rem;
    padding-bottom: 5rem;
    overflow: auto;
  }
  .material-excel-row {
    display: flex;
    border-style: solid;
    border-color: #EEE;
    border-width: 1px 0 0 1px;
    transition: background-color 200ms, color 100ms;
    color: #444;
    &.material-excel-hidden {
      display: none;
    }
    &.material-excel-title {
      background-color: #58BE6B !important;
      border-color: #58BE6B !important;
      border-bottom: #239538 1px solid !important;
      pointer-events: none;
      position: sticky;
      top: 0;
      z-index: 5;
      .material-excel-item {
        color: #FFF !important;
        border-color: #58BE6B !important;
        height: 2.6em;
      }
    }
    &:last-child {
      border-width: 1px 0 1px 1px;
    }
    &:hover {
      background-color: #58BE6B44;
      color: #000;
      cursor: pointer;
      .material-excel-item {
        border-color: transparent;
      }
    }
    &:active {
      transform: translateY(1px);
    }
    .material-excel-item {
      border-style: solid;
      border-color: #EEE;
      border-width: 0 1px 0 0;
      padding: 0 0.5rem;
      display: flex;
      align-items: center;
      transition: border-color 200ms;
      &.material-excel-item-check {
        width: 40px;
        padding: 0;
        text-align: center;
      }
      &.material-excel-item-en,
      &.material-excel-item-cn {
        width: 10rem;
        flex: 1;
      }
      &.material-excel-item-price {
        width: 10rem;
        justify-content: end;
      }
    }
  }
</style>
