const express = require('express');
const axios = require('axios');

const app = express();

// CORS
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Access-Token, Client-Id, Open-Id');
  if (req.method === 'OPTIONS') {
    return res.sendStatus(200);
  }
  next();
});

// 健康检查接口（用于测试服务是否正常）
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// 腾讯文档代理接口
app.get('/api/tencent-doc', async (req, res) => {
  const clientId = process.env.CLIENT_ID;
  const accessToken = process.env.ACCESS_TOKEN;
  const openId = process.env.OPEN_ID || '';

  const { fileId, sheetId, range } = req.query;
  if (!clientId || !accessToken) {
    return res.status(500).json({ error: 'Missing CLIENT_ID or ACCESS_TOKEN env vars' });
  }
  if (!fileId || !sheetId) {
    return res.status(400).json({ error: 'Missing fileId or sheetId' });
  }

  const rangeParam = range || 'A1:Z1000';
  const url = `https://docs.qq.com/openapi/spreadsheet/v3/files/${fileId}/${sheetId}/${rangeParam}`;

  try {
    const response = await axios.get(url, {
      headers: {
        'Access-Token': accessToken,
        'Client-Id': clientId,
        ...(openId && { 'Open-Id': openId })
      }
    });

    const gridData = response.data?.data?.gridData;
    if (!gridData) {
      return res.status(500).json({ error: 'Invalid response structure' });
    }

    const rows = gridData.rows || [];
    const result = rows.map(row => {
      return (row.values || []).map(cell => cell?.cellValue?.text || '');
    });

    res.json({ code: 0, data: result });
  } catch (error) {
    console.error('Proxy error:', error.response?.data || error.message);
    res.status(error.response?.status || 500).json({
      error: 'Failed to fetch from Tencent Doc',
      details: error.response?.data || error.message
    });
  }
});

// 根路径响应，防止 404
app.get('/', (req, res) => {
  res.send('Proxy server is running.');
});

module.exports = app;
