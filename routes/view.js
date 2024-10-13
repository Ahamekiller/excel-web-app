const express = require('express');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');

const router = express.Router();

// 处理查看文件的路由
router.get('/:filename', (req, res) => {
  const filePath = path.join(__dirname, '../public/uploads/', req.params.filename);
  
  // 检查文件是否存在
  if (!fs.existsSync(filePath)) {
    return res.status(404).send('File not found');
  }

  // 读取Excel文件
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0]; // 默认读取第一个sheet
  const worksheet = workbook.Sheets[sheetName];
  
  // 将Excel内容转换为JSON数组
  const data = xlsx.utils.sheet_to_json(worksheet);

  // 获取分页和搜索参数
  const page = parseInt(req.query.page) || 1;
  const search = req.query.search || '';
  const pageSize = 100;

  // 筛选搜索内容
  const filteredData = data.filter(row => {
    return Object.values(row).some(value => String(value).toLowerCase().includes(search.toLowerCase()));
  });

  // 计算分页内容
  const totalItems = filteredData.length;
  const totalPages = Math.ceil(totalItems / pageSize);
  const paginatedData = filteredData.slice((page - 1) * pageSize, page * pageSize);

  // 渲染页面
  res.render('table', {
    filename: req.params.filename,
    data: paginatedData,
    page,
    totalPages,
    search,
    totalItems
  });
});

module.exports = router;
