<script>
    import MaterialExcel from "./material-excel.svelte";
</script>
<section class="layout">
    <header class="layout-header">InfinitySTUDIO</header>
    <section class="layout-content">
        <MaterialExcel />
    </section>
    <footer class="layout-footer"></footer>
</section>
<style lang="scss">
  .layout {
    height: 100vh;
    width: 100vw;
    overflow: hidden;
    display: flex;
    flex-direction: column;
    .layout-header {
      font-size: 1.5rem;
      height: 3rem;
      line-height: 3rem;
      text-indent: 2rem;
      background-color: #58BE6B;
      color: #FFF;
      text-shadow: 1px 2px 0 #444, 0 0 0.5rem #FFF;
      box-shadow: 0 0 0.35rem #AAA;
    }
    .layout-content {
      flex: 1;
      height: 50vh;
    }
  }
</style>