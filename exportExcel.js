import { utils, write } from "xlsx-js-style";
import baseExportData, { setSpaceRow, wTn } from "./baseExportData";

const RMB = '"¥"#,##0.00_);\\("¥"#,##0.00\\)';
export default data => {
  const workbook = utils.book_new();
  const worksheet = baseExportData();
  // user select data
  const normals = [];
  const consumables = [];
  const staffs = [];
  data.forEach(row => {
    if (!row.isTitle) {
      switch (row.rowType) {
        case "normal":
          normals.push(row);
          break;
        case "consumable":
          consumables.push(row);
          break;
        case "staff":
          staffs.push(row);
          break;
      }
    }
  });
  let rowIndex = 6;
  for (let i = 0;i < normals.length;i++) {
    setRowData(worksheet, rowIndex, normals[i]);
    rowIndex++;
  }
  // real merge r
  worksheet['!merges'].forEach(mer => {
    mer.s.r--;
    mer.e.r--;
  });
  // real lineHeight
  worksheet['!rows'].shift();
  // effect
  worksheet['!ref'] = utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 1000, c: wTn.J } });
  utils.book_append_sheet(workbook, worksheet, 'Sheet');
  const excelBuffer = write(workbook, { bookType: 'xlsx', type: 'array', cellStyles: true });
  const excelData = new Blob([excelBuffer], { type: 'application/octet-stream' });
  // 下载Excel文件
  const downloadLink = document.createElement('a');
  downloadLink.href = URL.createObjectURL(excelData);
  downloadLink.download = 'output.xlsx';
  downloadLink.click();
  URL.revokeObjectURL(downloadLink.href);
}

const setRowData = (ws, rowIndex, row) => {
  const style = {
    font: { name: "微软雅黑", sz: 8 },
    alignment: { horizontal: 'center', vertical: 'center' },
  };
  const alignRightStyle = {
    font: { name: "微软雅黑", sz: 8 },
    alignment: { horizontal: 'right', vertical: 'center' },
  };
  console.log(rowIndex);
  ws[`B${rowIndex}`] = {
    v: row.__EMPTY_2.trim(),
    t: "s",
    s: style
  }
  ws[`D${rowIndex}`] = {
    v: row.goodsSize,
    t: "n",
    s: style
  }
  ws[`E${rowIndex}`] = {
    v: row.__EMPTY_3,
    t: "n",
    s: alignRightStyle,
    z: RMB
  }
  ws[`F${rowIndex}`] = {
    v: 1,
    t: "n",
    s: style
  }
  ws[`G${rowIndex}`] = {
    t: "n",
    s: alignRightStyle,
    f: `D${rowIndex}*E${rowIndex}*F${rowIndex}`,
    z: RMB
  }
  // utils.format_cell(ws[`E${rowIndex}`]);
  ws['!rows'][rowIndex] = { hpt: 24 };
  ws['!merges'].push({ s: { c: wTn.B, r: rowIndex }, e: { c: wTn.C, r: rowIndex } });
}
