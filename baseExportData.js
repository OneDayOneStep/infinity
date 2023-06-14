import { utils } from "xlsx";

export const wTn =
	{ "A": 0, "B": 1, "C": 2, "D": 3, "E": 4, "F": 5, "G": 6, "H": 7, "I": 8, "J": 9, "K": 10, "L": 11, "M": 12,
		"N": 13, "O": 14, "P": 15, "Q": 16, "R": 17, "S": 18, "T": 19, "U": 20, "V": 21, "W": 22, "X": 23, "Y": 24, "Z": 25 };

const spaceHeight = 5;
const spaceBgColor = "4F6228";
const headerBgColor = "E2EFDA";
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
		v: "小黑（澳门星光综艺馆）",
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
		v: "23/07/6-8",
		t: "s",
		s: {
			font: { name: "微软雅黑", sz: 12, bold: true },
			alignment: { horizontal: 'center', vertical: 'center' },
		}
	}
	const headerStyle = {
		fill: { fgColor: { rgb: headerBgColor } },
		font: { name: "微软雅黑", sz: 9, bold: true },
		alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
	};
	ws['A4'] = {
		v: "序号\nNumber",
		t: "s",
		s: headerStyle
	}
	ws['B4'] = {
		v: "详细\nDetailed",
		t: "s",
		s: headerStyle
	}
	ws['D4'] = {
		v: "数量\nQuantity",
		t: "s",
		s: headerStyle
	}
	ws['E4'] = {
		v: "出租价/天\nUnit Price",
		t: "s",
		s: headerStyle
	}
	ws['F4'] = {
		v: "天数\nDays",
		t: "s",
		s: headerStyle
	}
	ws['G4'] = {
		v: "金额\nAmount",
		t: "s",
		s: headerStyle
	}
	ws['H4'] = {
		v: "折扣\nDiscount",
		t: "s",
		s: headerStyle
	}
	ws['I4'] = {
		v: "折扣金额\nDiscount amount",
		t: "s",
		s: headerStyle
	}
	ws['J4'] = {
		v: "备注\nRemarks",
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
			fill: { fgColor: { rgb: spaceBgColor } }
		}
	}
	ws['!merges'].push({ s: { c: wTn.A, r: rowNum }, e: { c: wTn.J, r: rowNum   } });
}
