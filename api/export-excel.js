const ExcelJS = require('exceljs');

// 跨域处理
const allowCors = fn => async (req, res) => {
  res.setHeader('Access-Control-Allow-Credentials', true)
  res.setHeader('Access-Control-Allow-Origin', '*')
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS')
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type')
  if (req.method === 'OPTIONS') return res.status(200).end()
  return await fn(req, res)
}

// 核心生成Excel逻辑
const handler = async (req, res) => {
  try {
    const list = req.body || [];
    if (!Array.isArray(list) || list.length === 0) {
      return res.status(400).json({ msg: '暂无导出数据' });
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('部门机构报备');

    // 全局列宽
    worksheet.columns = [
      { width: 23 },
      { width: 32.25 },
      { width: 28.875 },
      { width: 16.75 },
      { width: 33 },
      { width: 28.625 },
      { width: 9.375 },
      { width: 10 },
      { width: 9.25 },
      { width: 37 }
    ];

    // 第1行 A1-J1 合并 + 富文本标题
    const row1 = worksheet.getRow(1);
    row.height = 63.75;
    worksheet.mergeCells('A1:J1');
    const cellA1 = worksheet.getCell('A1');
    cellA1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cellA1.richText = [
      {
        font: { name: '方正小标宋简体', size: 20, color: { argb: '000000' } },
        text: '全区一体化政务服务平台系统区划变更清单\n'
      },
      {
        font: { name: '方正小标宋简体', size: 20, color: { argb: 'FF0000' }, bold: true },
        text: '变更/新增统一社会信用代码填写此表（如有多条数据请分条填写）'
      }
    ];

    // 第2行 A2-J2 联系人
    const row2 = worksheet.getRow(2);
    row2.height = 44.25;
    worksheet.mergeCells('A2:J2');
    const cellA2 = worksheet.getCell('A2');
    cellA2.value = '具体填报联系人（必填）：       乔鹏飞         联系电话（必填）：       18148340753';
    cellA2.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cellA2.font = { name: '方正仿宋_GBK', size: 11, color: { argb: '000000' } };

    // 第3行 表头
    const row3 = worksheet.getRow(3);
    row3.height = 129;
    const headers = [
      "统一社会信用代码（必填）",
      "组织机构全称（必填）",
      "组织机构简称（必填）",
      "所属区划编码必填（必填）",
      "上级组织机构编码必填（必填：统一填写上级政务服务局）",
      "所属行政区划（必填格式为：盟市——市/区/直属部门——镇/直属部门）",
      "组织机构类型",
      "虚拟部门",
      "新增/修改",
      "备注（修改请备注修改前和修改后）"
    ];
    headers.forEach((text, idx) => {
      const cell = row3.getCell(idx + 1);
      cell.value = text;
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.font = { name: '宋体', size: 11, color: { argb: '000000' } };
    });

    // 数据行
    let currentRow = 4;
    list.forEach(item => {
      const dataRow = worksheet.getRow(currentRow);
      dataRow.height = 129;
      dataRow.getCell(1).value = item.creditCode;
      dataRow.getCell(2).value = item.fullName;
      dataRow.getCell(3).value = item.fullName;
      dataRow.getCell(4).value = item.areaCode;
      dataRow.getCell(5).value = item.superCode;
      dataRow.getCell(6).value = item.areaName;
      dataRow.getCell(7).value = item.orgType;
      dataRow.getCell(8).value = '否';
      dataRow.getCell(9).value = item.operate;

      // 备注富文本（红名）
      const noteCell = dataRow.getCell(10);
      if (item.note) {
        noteCell.richText = [
          { font: { name: '宋体', size: 11, color: { argb: '000000' } }, text: '将原' },
          { font: { name: '宋体', size: 11, color: { argb: 'FF0000' }, bold: true }, text: item.oldName },
          { font: { name: '宋体', size: 11, color: { argb: '000000' } }, text: '修改为' },
          { font: { name: '宋体', size: 11, color: { argb: 'FF0000' }, bold: true }, text: item.fullName }
        ];
      }

      // 整行样式
      for (let i = 1; i <= 10; i++) {
        const cell = dataRow.getCell(i);
        cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
        cell.font = { name: '宋体', size: 11, color: { argb: '000000' } };
      }
      currentRow++;
    });

    // 全局边框
    const maxRow = currentRow - 1;
    for (let r = 1; r <= maxRow; r++) {
      for (let c = 1; c <= 10; c++) {
        const cell = worksheet.getCell(r, c);
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
          bottom: { style: 'thin' }
        };
      }
    }

    // 输出Excel流
    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="区划变更清单_${list.length}条数据.xlsx"`);
    res.end(buffer);

  } catch (err) {
    console.error(err);
    res.status(500.json({ msg: '服务器生成Excel失败' });
  }
};

module.exports = allowCors(handler);
