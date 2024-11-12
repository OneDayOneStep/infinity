import { utils } from "xlsx";
import { COMPANY_INFOS } from "./companys";

console.log(import.meta, 1)
const companyInfo = COMPANY_INFOS[import.meta.env.VITE_COMPANY_NAME];

export const wTn =
	{ "A": 0, "B": 1, "C": 2, "D": 3, "E": 4, "F": 5, "G": 6, "H": 7, "I": 8, "J": 9, "K": 10, "L": 11, "M": 12,
		"N": 13, "O": 14, "P": 15, "Q": 16, "R": 17, "S": 18, "T": 19, "U": 20, "V": 21, "W": 22, "X": 23, "Y": 24, "Z": 25 };
const _nTw = [];
Object.keys(wTn).forEach((k, i) => {
	_nTw[i] = k;
})
export const nTw = _nTw;

const spaceHeight = 5;
export default () => {
	const ws = utils.aoa_to_sheet([]);
	// colWidth
	ws['!cols'] = ws['!cols'] || [];
	ws["!cols"][wTn.A] = { wch: 17 };
	ws["!cols"][wTn.B] = { wch: 23 };
	ws["!cols"][wTn.C] = { wch: 46 };
	ws["!cols"][wTn.D] = { wch: 12 };
	ws["!cols"][wTn.E] = { wch: 10 };
	ws["!cols"][wTn.F] = { wch: 5 };
	ws["!cols"][wTn.G] = { wch: 11 };
	ws["!cols"][wTn.H] = { wch: 12 };
	ws["!cols"][wTn.I] = { wch: 18 };
	ws["!cols"][wTn.J] = { wch: 30 };
	// rowHeight
	ws['!rows'] = ws['!rows'] || [];
	ws['!rows'][1] = { hpt: 40 };
	ws['!rows'][3] = { hpt: 35 };
	ws['!rows'][4] = { hpt: 35 };
	// merge
	ws['!merges'] = ws['!merges'] || [];
	ws['!merges'].push({ s: { c: wTn.A, r: 1 }, e: { c: wTn.J, r: 1,  } });
	ws['!merges'].push({ s: { c: wTn.A, r: 2 }, e: { c: wTn.J, r: 2   } });
	ws['!merges'].push({ s: { c: wTn.B, r: 3 }, e: { c: wTn.D, r: 3   } });
	ws['!merges'].push({ s: { c: wTn.E, r: 3 }, e: { c: wTn.F, r: 3   } });
	ws['!merges'].push({ s: { c: wTn.G, r: 3 }, e: { c: wTn.J, r: 3   } });
	ws['!merges'].push({ s: { c: wTn.B, r: 4 }, e: { c: wTn.C, r: 4   } });
	ws['A1'] = {
		v: "InfinitySTUDIO器材租赁 报价单",
		t: "s",
		s: {
			font: { name: "微软雅黑", sz: 28, bold: true },
			alignment: { horizontal: 'center', vertical: 'center' },
		}
	}
	ws['A2'] = {
		v: "地址Add：广州荔湾区海龙路304号       联系Contact：+86 153 6041 6740",
		t: "s",
		s: {
			font: { name: "微软雅黑", sz: 9 },
			alignment: { horizontal: 'left', vertical: 'center' },
		}
	}
	ws['A3'] = {
		v: "项目：",
		t: "s",
		s: {
			font: { name: "微软雅黑", sz: 12, bold: true },
			alignment: { horizontal: 'center', vertical: 'center' },
		}
	}
	ws['B3'] = {
		v: "Project S",
		t: "s",
		s: {
			font: { name: "微软雅黑", sz: 12, bold: true },
			alignment: { horizontal: 'center', vertical: 'center' },
		}
	}
	ws['E3'] = {
		v: "日期：",
		t: "s",
		s: {
			font: { name: "微软雅黑", sz: 12, bold: true },
			alignment: { horizontal: 'center', vertical: 'center' },
		}
	}
	ws['G3'] = {
		v: new Date().getFullYear().toString().slice(2, 4) + "/" +
			(new Date().getMonth() + 1).toString().padStart(2, "0") + "/" +
			new Date().getDate().toString(),
		t: "s",
		s: {
			font: { name: "微软雅黑", sz: 12, bold: true },
			alignment: { horizontal: 'center', vertical: 'center' },
		}
	}
	const headerStyle = {
		fill: { fgColor: { rgb: companyInfo.headerBgColor } },
		font: { name: "微软雅黑", sz: 9, bold: true },
		alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
	};
	ws['A4'] = {
		v: companyInfo.H_XH,
		t: "s",
		s: headerStyle
	}
	ws['B4'] = {
		v: companyInfo.H_XX,
		t: "s",
		s: headerStyle
	}
	ws['D4'] = {
		v: companyInfo.H_SL,
		t: "s",
		s: headerStyle
	}
	ws['E4'] = {
		v: companyInfo.H_CZJ,
		t: "s",
		s: headerStyle
	}
	ws['F4'] = {
		v: companyInfo.H_TS,
		t: "s",
		s: headerStyle
	}
	ws['G4'] = {
		v: companyInfo.H_JE,
		t: "s",
		s: headerStyle
	}
	ws['H4'] = {
		v: companyInfo.H_ZK,
		t: "s",
		s: headerStyle
	}
	ws['I4'] = {
		v: companyInfo.H_ZKJE,
		t: "s",
		s: headerStyle
	}
	ws['J4'] = {
		v: companyInfo.H_BZ,
		t: "s",
		s: headerStyle
	}
	setSpaceRow(ws, 5);
	return ws;
}

export const setSpaceRow = (ws, rowNum) => {
	ws['!rows'][rowNum] = { hpt: spaceHeight };
	ws[`A${rowNum}`] = {
		v: "",
		t: "s",
		s: {
			fill: { fgColor: { rgb: companyInfo.spaceBgColor } }
		}
	}
	ws['!merges'].push({ s: { c: wTn.A, r: rowNum }, e: { c: wTn.J, r: rowNum   } });
}
