import { utils, write } from "xlsx-js-style";
import baseExportData, { setSpaceRow, wTn, nTw } from "./baseExportData";

const RMB = '"¥"#,##0.00_);\\("¥"#,##0.00\\)';
const style = {
	font: { name: "微软雅黑", sz: 8 },
	alignment: { horizontal: 'center', vertical: 'center' },
};
const alignRightStyle = {
	font: { name: "微软雅黑", sz: 8 },
	alignment: { horizontal: 'right', vertical: 'center' },
};
const contentStartRow = 6;
export default (data, params) => {
  const workbook = utils.book_new();
  const worksheet = baseExportData();
  // user select data
	const selectDatas = {
		normal: {
			title: "器材租金\nEquipment Rental",
			data: []
		},
		consumable: {
			title: "耗材费用\nConsumables",
			data: []
		},
		staff: {
			title: "人员费用\nAssistant",
			data: []
		}
	}
  data.forEach(row => {
    if (!row.isTitle) {
      switch (row.rowType) {
	      case "normal":
	        selectDatas.normal.data.push(row);
          break;
        case "consumable":
	        selectDatas.consumable.data.push(row);
          break;
        case "staff":
	        selectDatas.staff.data.push(row);
          break;
      }
    }
  });
	let rowIndex = contentStartRow;
	let startRowIndex;
	// normal consumable staff
	Object.keys(selectDatas).forEach(k => {
		const kData = selectDatas[k].data;
		if (kData.length > 0) {
			startRowIndex = rowIndex;
			for (let i = 0;i < kData.length;i++) {
				setRowData(worksheet, rowIndex++, kData[i], params);
			}
			worksheet[`A${startRowIndex}`] = {
				v: selectDatas[k].title,
				t: "s",
				s: {
					font: { name: "微软雅黑", sz: 10, bold: true },
					alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
				}
			}
			worksheet['!merges'].push({ s: { c: wTn.A, r: startRowIndex }, e: { c: wTn.A, r: kData.length + startRowIndex - 1,  } });
			setSpaceRow(worksheet, rowIndex++);
		}
	})
	// travel
	setRowData(worksheet, rowIndex, {
		rowType: "travel",
		goodsSize: 1,
		__EMPTY_2: "器材报关/人员差旅费",
		__EMPTY_3: 0,
	}, params);
	worksheet['!rows'][rowIndex] = { hpt: 35 };
	worksheet[`A${rowIndex}`] = {
		v: "差旅\nTravel expenses",
		t: "s",
		s: {
			font: { name: "微软雅黑", sz: 10, bold: true },
			alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
		}
	}
	worksheet[`J${rowIndex++}`] = {
		v: "实报实销",
		t: "s",
		s: {
			font: { name: "微软雅黑", sz: 8, bold: true },
			alignment: { horizontal: 'center', vertical: 'center' },
		}
	}
	// total row
	worksheet[`A${rowIndex}`] = {
		v: "原价总价（CNY）：",
		t: "s",
		s: {
			font: { name: "微软雅黑", sz: 8, bold: true },
			alignment: { horizontal: 'right', vertical: 'center' },
		}
	}
	const contentEndRow = rowIndex - 1;
	worksheet[`G${rowIndex}`] = {
		t: "n",
		s: {
			...alignRightStyle,
			font: {
				...alignRightStyle.font,
				bold: true
			}
		},
		f: `SUM(G${contentStartRow}:G${contentEndRow})`,
		z: RMB
	}
	worksheet['!merges'].push({ s: { c: wTn.A, r: rowIndex }, e: { c: wTn.F, r: rowIndex,  } });
	worksheet['!rows'][rowIndex++] = { hpt: 24 };
	// remark row
	worksheet[`A${rowIndex}`] = {
		v: "备注：\n" +
			"1）以上器材价格，不含税\n" +
			"（人员/器材运输/耗材，核酸检测费用，实报实销）。\n" +
			" 餐费另计，如有包餐可不计。 \n" +
			"餐费收费标准准 早餐15元/人，午餐30元/人，晚餐30元/人\n" +
			"\n" +
			"2）器材出库前3天，器材还可作调整。器材一旦出库后，\n" +
			"如不使用，不能作取消处理。\n" +
			"\n" +
			"3）凡客户租用InfinitySTUDIO器材，必须配备器材助理。\n" +
			"未配备助理，器材如有损坏全部责任由租方负责。",
		t: "s",
		s: {
			font: { name: "微软雅黑", sz: 10 },
			alignment: { horizontal: 'left', vertical: 'top', wrapText: true },
		}
	}
	worksheet['!merges'].push({ s: { c: wTn.A, r: rowIndex }, e: { c: wTn.D, r: rowIndex,  } });
	worksheet[`E${rowIndex}`] = {
		v: "助理工作收费标准：\n" +
			"1）Calltime 06:00AM-00:00AM 时间段 \n" +
			"专业助理 InfinitySTUDIO影棚：800元/人\n" +
			"（到场起计工作10小时内，超时 100元/小时/人，不设停钟）\n" +
			"\n" +
			"2）Calltime 00:00AM-06:00AM 时间段 ：1500元/人\n" +
			"（到场起计工作6小时内，超时300元/小时/人，不设停钟）\n" +
			"\n" +
			"3）非广东省内助理费用加收200元/人\n" +
			"\n" +
			"4）路程时间4-8h，人员按照半天收费\n" +
			"路程时间8h-12h，人员按照一天收费",
		t: "s",
		s: {
			font: { name: "微软雅黑", sz: 10 },
			alignment: { horizontal: 'left', vertical: 'top', wrapText: true },
		}
	}
	worksheet['!merges'].push({ s: { c: wTn.E, r: rowIndex }, e: { c: wTn.J, r: rowIndex,  } });
	worksheet['!rows'][rowIndex++] = { hpt: 190 };
	// finally price
	worksheet[`A${rowIndex}`] = {
		v: "优惠价（CNY）：",
		t: "s",
		s: {
			font: { name: "微软雅黑", sz: 8, bold: true },
			alignment: { horizontal: 'right', vertical: 'center' },
		}
	}
	worksheet['!merges'].push({ s: { c: wTn.A, r: rowIndex }, e: { c: wTn.H, r: rowIndex,  } });
	worksheet[`I${rowIndex}`] = {
		v: "优惠价（CNY）：",
		t: "s",
		s: {
			font: { name: "微软雅黑", sz: 8, bold: true },
			alignment: { horizontal: 'right', vertical: 'center' },
		}
	}
	worksheet[`I${rowIndex}`] = {
		t: "n",
		s: {
			...style,
			font: {
				...style.font,
				bold: true,
				underline: true,
				color: { rgb: "FF0000" }
			}
		},
		f: `SUM(I${contentStartRow}:I${contentEndRow})`,
		z: RMB
	}
	worksheet[`J${rowIndex}`] = {
		v: "不含税，不含人员超时费用\n器材报关/运费费用实报实销",
		t: "s",
		s: {
			font: { name: "微软雅黑", sz: 8, bold: true },
			alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
		}
	}
	worksheet['!rows'][rowIndex++] = { hpt: 24 };
	// last
	worksheet[`A${rowIndex}`] = {
		v: "租赁条款：\n" +
			"一、 本条款中所涉及的费用均已天（二十四小时）计算，并从器材交付时起算。周租金（五天）按五天日租金的优惠价计算。\n" +
			"\n" +
			"二、 所有涉及器材租赁以及人员费用应在返还器材时结算。\n" +
			"\n" +
			"三、 所有器材租赁前，需支付所租赁器材等值的押金。我们接受租赁人身份证明原件（中国公民身份证，外国公民护照，或与租赁器材等值的现金作为抵押\n" +
			"（ 不接收支票）。\n" +
			"\n" +
			"四、   承租人在器材租赁及使用期间，需对租赁器材负全责。如因承租人疏忽引起器材损坏或遗失，承租人必须按器材的等价进行赔偿，并付清到归还日期为止的器材租赁费。\n" +
			"逾期赔偿，将按逾期天数加收逾期费用（按日租金计），直到承租人付清全额赔偿款为止。保险索赔均由承租人承担，InfinitySTUDIO不负责保险索赔事项。\n" +
			"\n" +
			"五、   承租人需对器材的任何不适当使用而引起的器材损坏负责。返还器材时必须保障器材的状态良好。一旦发生器材损坏情况，须第一时间通知Infinity STUDIO。因损坏而发生的修理费用，在返还当天应由承租人全额支付，包括因器材维修所造成的租金的损失。\n" +
			"\n" +
			"六、InfinitySTUDIO的器材在出租时均保持清洁状态。因此，承租人在返还时也必须保持器材清洁，避免发生清理器材的费用。\n" +
			"\n" +
			"七、InfinitySTUDIO 所有器材在离开公司前均进行过精心检查和测试，建议承租人在交收器材时仔细检查。\n" +
			"\n" +
			"八、 承租人需在租赁前清楚器材的功能及使用方法，并由专业人员使用所租赁的器材。如器材因操作不当而引起损坏，承租人不能以此为拒绝赔偿的理由。\n" +
			"\n" +
			"九、 如器材发生故障，建议承租人及时与我们联系。我们将及时采取无偿更换器材等措施。如因承租人原因造成的故障（忽略、滥用或使用不当），承租人必 须全额支付所更换器材的租金。\n" +
			"\n" +
			"十、   所有租赁器材在通常情况下需由承租人亲自过来本公司领取。如承租人要求我们接送器材，运输费与人员费由承租人承担。我们建议承租人使用自己指定的运输公司，因为一旦器材离开工作室，承租人必须负全责。如承租人要求我们安排运输公司，承租人也要如上述情况负全责。本公司不对运输途中发生的任何问题负责。\n" +
			"\n" +
			"十一、所有器材的预定和取消要提前一天（二十四小时）与我们联系。如当天取消订单的按一天收取费用。",
		t: "s",
		s: {
			font: { name: "微软雅黑", sz: 7 },
			alignment: { horizontal: 'left', vertical: 'top', wrapText: true },
		}
	}
	worksheet['!merges'].push({ s: { c: wTn.A, r: rowIndex }, e: { c: wTn.J, r: rowIndex,  } });
	worksheet['!rows'][rowIndex++] = { hpt: 300 };
	// border
	// const border = {
	// 	top: { style: 'hair' },
	// 	bottom: { style: 'hair' },
	// 	left: { style: 'hair' },
	// 	right: { style: 'hair' }
	// }
	// for (let i = 0;i < wTn.K;i++) {
	// 	for (let j = 3;j < rowIndex;j++) {
	// 		const key = `${nTw[i]}${j}`;
	// 		worksheet[key] = worksheet[key] || {};
	// 		worksheet[key].s = worksheet[key].s || {};
	// 		worksheet[key].s.border = border;
	// 	}
	// }
  // real merge r
  worksheet['!merges'].forEach(mer => {
    mer.s.r--;
    mer.e.r--;
  });
  // real lineHeight
  worksheet['!rows'].shift();
  // effect
  worksheet['!ref'] = utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rowIndex + 50, c: wTn.K } });
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

const setRowData = (ws, rowIndex, row, { days = 1, discount = "100%" }) => {
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
    v: days,
    t: "n",
    s: style
  }
  ws[`G${rowIndex}`] = {
    t: "n",
    s: alignRightStyle,
    f: `D${rowIndex}*E${rowIndex}*F${rowIndex}`,
    z: RMB
  }
	ws[`H${rowIndex}`] = {
		v: ["consumable", "staff", "travel"].includes(row.rowType) ? "100%" : discount,
		t: "s",
		s: style
	}
	ws[`I${rowIndex}`] = {
		t: "n",
		s: {
			...alignRightStyle,
			font: {
				...alignRightStyle.font,
				bold: true
			}
		},
		f: `G${rowIndex}*H${rowIndex}`,
		z: RMB
	}
  ws['!rows'][rowIndex] = { hpt: 24 };
  ws['!merges'].push({ s: { c: wTn.B, r: rowIndex }, e: { c: wTn.C, r: rowIndex } });
}
