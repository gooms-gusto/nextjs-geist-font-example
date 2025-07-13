const express = require('express');
const multer = require('multer');
const path = require('path');
const { asyncHandler } = require('../middleware/errorHandler');
const excelController = require('../controllers/excelController');

const router = express.Router();

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, path.join(__dirname, '../templates'));
  },
  filename: (req, file, cb) => {
    // Generate unique filename with timestamp
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
  }
});

const fileFilter = (req, file, cb) => {
  // Accept only Excel files
  const allowedTypes = ['.xlsx', '.xls'];
  const fileExtension = path.extname(file.originalname).toLowerCase();
  
  if (allowedTypes.includes(fileExtension)) {
    cb(null, true);
  } else {
    cb(new Error('Only Excel files (.xlsx, .xls) are allowed'), false);
  }
};

const upload = multer({
  storage: storage,
  limits: {
    fileSize: parseInt(process.env.MAX_FILE_SIZE) || 10485760 // 10MB default
  },
  fileFilter: fileFilter
});

/**
 * @route   POST /api/excel/generate
 * @desc    Generate Excel file from JSON payload
 * @access  Public
 * @body    {
 *   template?: string,
 *   filename?: string,
 *   sheets: Array<{
 *     name: string,
 *     cells?: Array<{
 *       cell: string,
 *       value: any,
 *       formula?: string,
 *       style?: object,
 *       dataType?: string
 *     }>,
 *     ranges?: Array<{
 *       range: string,
 *       data: Array<Array<any>>,
 *       style?: object
 *     }>,
 *     tables?: Array<{
 *       name: string,
 *       range: string,
 *       data: Array<object>,
 *       style?: object
 *     }>
 *   }>
 * }
 */
router.post('/generate', asyncHandler(excelController.generateExcel));

/**
 * @route   POST /api/excel/query
 * @desc    Generate Excel file from database query
 * @access  Public
 * @body    {
 *   query: string,
 *   params?: Array<any>,
 *   sheetName?: string,
 *   filename?: string,
 *   startCell?: string,
 *   includeHeaders?: boolean,
 *   style?: object
 * }
 */
router.post('/query', asyncHandler(excelController.generateFromQuery));

/**
 * @route   POST /api/excel/template
 * @desc    Upload Excel template file
 * @access  Public
 * @form    template: file
 */
router.post('/template', upload.single('template'), asyncHandler(excelController.uploadTemplate));

/**
 * @route   GET /api/excel/templates
 * @desc    List available Excel templates
 * @access  Public
 */
router.get('/templates', asyncHandler(excelController.listTemplates));

/**
 * @route   DELETE /api/excel/templates/:filename
 * @desc    Delete a specific template
 * @access  Public
 */
router.delete('/templates/:filename', asyncHandler(excelController.deleteTemplate));

/**
 * @route   POST /api/excel/template/fill
 * @desc    Fill template with data
 * @access  Public
 * @body    {
 *   template: string,
 *   data: object,
 *   filename?: string
 * }
 */
router.post('/template/fill', asyncHandler(excelController.fillTemplate));

/**
 * @route   POST /api/excel/multi-sheet
 * @desc    Generate Excel with multiple sheets from different queries
 * @access  Public
 * @body    {
 *   sheets: Array<{
 *     name: string,
 *     query?: string,
 *     params?: Array<any>,
 *     data?: Array<object>,
 *     style?: object
 *   }>,
 *   filename?: string
 * }
 */
router.post('/multi-sheet', asyncHandler(excelController.generateMultiSheet));

/**
 * @route   POST /api/excel/styled
 * @desc    Generate Excel with advanced styling
 * @access  Public
 * @body    {
 *   sheets: Array<{
 *     name: string,
 *     data: Array<object>,
 *     styles: {
 *       header?: object,
 *       rows?: object,
 *       alternateRows?: object,
 *       columns?: object
 *     },
 *     formatting?: {
 *       autoWidth?: boolean,
 *       freezeRows?: number,
 *       freezeCols?: number
 *     }
 *   }>,
 *   filename?: string
 * }
 */
router.post('/styled', asyncHandler(excelController.generateStyledExcel));

/**
 * @route   GET /api/excel/schema/:table?
 * @desc    Get database schema information
 * @access  Public
 */
router.get('/schema/:table?', asyncHandler(excelController.getSchema));

/**
 * @route   POST /api/excel/validate-query
 * @desc    Validate SQL query without executing
 * @access  Public
 * @body    { query: string }
 */
router.post('/validate-query', asyncHandler(excelController.validateQuery));

/**
 * @route   GET /api/excel/sample-data
 * @desc    Get sample data for testing
 * @access  Public
 */
router.get('/sample-data', asyncHandler(excelController.getSampleData));

module.exports = router;
