<script>
	import { read, utils, writeFile } from 'xlsx';
	import { writable } from 'svelte/store';
	import Checkbox from '@smui/checkbox';

  let list = writable([]);
	fetch("./material_origin.xlsx")
		.then(res => res.arrayBuffer())
    .then(buffer => {
	    const excel = read(buffer);
	    const data = utils.sheet_to_json(excel.Sheets[excel.SheetNames[0]]);
			list.set(data);
    })

  const clickRow = index => {
		console.log(list);
  }
</script>

<div class="material-excel-list">
  {#each $list as { checked, __EMPTY_1 = "", __EMPTY_2 = "", __EMPTY_3 = "" }, i}
    <div class="material-excel-row" on:click={() => clickRow(i)}>
      <div class="material-excel-item material-excel-item-check">
        <Checkbox bind:checked />
      </div>
      <div class="material-excel-item material-excel-item-en">{ __EMPTY_1 }</div>
      <div class="material-excel-item material-excel-item-cn">{ __EMPTY_2 }</div>
      <div class="material-excel-item material-excel-item-price">{ __EMPTY_3 }</div>
    </div>
  {/each}
</div>

<style lang="scss">
  .material-excel-list {
    margin: 0 1rem;
  }
  .material-excel-row {
    display: flex;
    border-style: solid;
    border-color: #EEE;
    border-width: 1px 0 0 1px;
    transition: background-color 200ms, color 100ms;
    color: #444;
    &:last-child {
      border-width: 1px 0 1px 1px;
    }
    &:hover {
      background-color: #58BE6B44;
      color: #000;
      cursor: pointer;
    }
    .material-excel-item {
      border-style: solid;
      border-color: #EEE;
      border-width: 0 1px 0 0;
      padding: 0.2rem 0.3rem;
      display: flex;
      align-items: center;
      &.material-excel-item-check {
        width: 40px;
        text-align: center;
      }
      &.material-excel-item-en,
      &.material-excel-item-cn {
        width: 10rem;
        flex: 1;
      }
      &.material-excel-item-price {
        width: 10rem;
        text-align: right;
      }
    }
  }
</style>