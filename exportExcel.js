import { utils, write } from "xlsx-js-style";
import baseExportData from "./baseExportData";

const wTn =
	{ "A": 0, "B": 1, "C": 2, "D": 3, "E": 4, "F": 5, "G": 6, "H": 7, "I": 8, "J": 9, "K": 10, "L": 11, "M": 12,
		"N": 13, "O": 14, "P": 15, "Q": 16, "R": 17, "S": 18, "T": 19, "U": 20, "V": 21, "W": 22, "X": 23, "Y": 24, "Z": 25 };
export default data => {
	const workbook = utils.book_new();
	const worksheet = baseExportData();
	// real merge r
	worksheet['!merges'].forEach(mer => {
		mer.s.r--;
		mer.e.r--;
	});
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