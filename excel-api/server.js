require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const morgan = require('morgan');
const path = require('path');
const fs = require('fs');

// Import routes and middleware
const excelRoutes = require('./routes/excelRoutes');
const { errorHandler } = require('./middleware/errorHandler');

// Create Express application
const app = express();

// Create logs directory if it doesn't exist
const logsDir = path.join(__dirname, 'logs');
if (!fs.existsSync(logsDir)) {
  fs.mkdirSync(logsDir, { recursive: true });
}

// CORS configuration
const corsOptions = {
  origin: process.env.CORS_ORIGIN || 'http://localhost:8000',
  methods: ['GET', 'POST', 'PUT', 'DELETE'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  credentials: true
};

// Middleware
app.use(cors(corsOptions));
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true, limit: '50mb' }));
app.use(morgan('combined'));

// Health check endpoint
app.get('/health', (req, res) => {
  res.status(200).json({
    status: 'OK',
    message: 'Excel API Server is running',
    timestamp: new Date().toISOString(),
    version: '1.0.0'
  });
});

// API Documentation endpoint
app.get('/api/docs', (req, res) => {
  res.json({
    title: 'Excel Generation API',
    version: '1.0.0',
    description: 'Advanced Excel generation API with database integration, templates, and styling',
    endpoints: {
      'POST /api/excel/generate': 'Generate Excel file from JSON payload',
      'POST /api/excel/template': 'Upload Excel template',
      'GET /api/excel/templates': 'List available templates',
      'POST /api/excel/query': 'Generate Excel from database query',
      'GET /health': 'Health check endpoint',
      'GET /api/docs': 'API documentation'
    },
    documentation: 'See /docs/API.md for detailed documentation'
  });
});

// Mount API routes
app.use('/api/excel', excelRoutes);

// Serve static files from templates directory
app.use('/templates', express.static(path.join(__dirname, 'templates')));

// 404 handler
app.use((req, res) => {
  res.status(404).json({
    error: 'Not Found',
    message: `Route ${req.method} ${req.path} not found`,
    availableEndpoints: [
      'GET /health',
      'GET /api/docs',
      'POST /api/excel/generate',
      'POST /api/excel/template',
      'GET /api/excel/templates',
      'POST /api/excel/query'
    ]
  });
});

// Global error handler (must be last)
app.use(errorHandler);

// Start server
const PORT = process.env.PORT || 9000;
const server = app.listen(PORT, () => {
  console.log(`
╔══════════════════════════════════════════════════════════════╗
║                    Excel API Server                         ║
║                                                              ║
║  🚀 Server running on port ${PORT}                              ║
║  🌐 Environment: ${process.env.NODE_ENV || 'development'}                        ║
║  📊 Excel Generation API Ready                               ║
║                                                              ║
║  Endpoints:                                                  ║
║  • GET  /health                                              ║
║  • GET  /api/docs                                            ║
║  • POST /api/excel/generate                                  ║
║  • POST /api/excel/template                                  ║
║  • GET  /api/excel/templates                                 ║
║  • POST /api/excel/query                                     ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝
  `);
});

// Graceful shutdown
process.on('SIGTERM', () => {
  console.log('SIGTERM received. Shutting down gracefully...');
  server.close(() => {
    console.log('Process terminated');
  });
});

process.on('SIGINT', () => {
  console.log('SIGINT received. Shutting down gracefully...');
  server.close(() => {
    console.log('Process terminated');
  });
});

module.exports = app;
