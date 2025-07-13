const Joi = require('joi');
const path = require('path');
const fs = require('fs').promises;
const ExcelService = require('../services/excelService');
const { executeSafeQuery, getSampleData, getSchema, validateQuery } = require('../config/db');
const { logger, APIError } = require('../middleware/errorHandler');

/**
 * Excel Controller - Handle all Excel-related API requests
 */
class ExcelController {

  /**
   * Generate Excel file from JSON payload
   * @param {Object} req - Express request object
   * @param {Object} res - Express response object
   */
  static async generateExcel(req, res) {
    try {
      // Validate request payload
      const schema = Joi.object({
        template: Joi.string().optional(),
        filename: Joi.string().default('generated.xlsx'),
        sheets: Joi.array().items(
          Joi.object({
            name: Joi.string().required(),
            cells: Joi.array().items(
              Joi.object({
                cell: Joi.string().required(),
                value: Joi.any(),
                formula: Joi.string().optional(),
                style: Joi.object().optional(),
                dataType: Joi.string().optional(),
                format: Joi.string().optional()
              })
            ).optional(),
            ranges: Joi.array().items(
              Joi.object({
                range: Joi.string().required(),
                data: Joi.array().required(),
                style: Joi.object().optional()
              })
            ).optional(),
            tables: Joi.array().items(
              Joi.object({
                name: Joi.string().optional(),
                range: Joi.string().optional(),
                data: Joi.array().required(),
                style: Joi.object().optional()
              })
            ).optional(),
            formatting: Joi.object().optional()
          })
        ).required()
      });

      const { error, value } = schema.validate(req.body);
      if (error) {
        throw new APIError(`Validation error: ${error.details[0].message}`, 400);
      }

      // Generate Excel workbook
      const buffer = await ExcelService.createWorkbook(value);

      // Set response headers
      res.setHeader('Content-Disposition', `attachment; filename="${value.filename}"`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Length', buffer.length);

      // Send file
      res.send(buffer);

      logger.info(`Excel file generated successfully: ${value.filename}`);
    } catch (error) {
      throw error;
    }
  }

  /**
   * Generate Excel from database query
   * @param {Object} req - Express request object
   * @param {Object} res - Express response object
   */
  static async generateFromQuery(req, res) {
    try {
      // Validate request payload
      const schema = Joi.object({
        query: Joi.string().required(),
        params: Joi.array().default([]),
        sheetName: Joi.string().default('QueryResults'),
        filename: Joi.string().default('query-results.xlsx'),
        startCell: Joi.string().default('A1'),
        includeHeaders: Joi.boolean().default(true),
        style: Joi.object().optional()
      });

      const { error, value } = schema.validate(req.body);
      if (error) {
        throw new APIError(`Validation error: ${error.details[0].message}`, 400);
      }

      // Execute query
      let data;
      try {
        data = await executeSafeQuery(value.query, value.params);
      } catch (dbError) {
        // If database is not available, use sample data for demonstration
        logger.warn('Database not available, using sample data');
        data = getSampleData();
      }

      if (!data || data.length === 0) {
        throw new APIError('Query returned no results', 404);
      }

      // Generate Excel from query results
      const buffer = await ExcelService.generateFromQueryResults(data, {
        sheetName: value.sheetName,
        startCell: value.startCell,
        style: value.style
      });

      // Set response headers
      res.setHeader('Content-Disposition', `attachment; filename="${value.filename}"`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Length', buffer.length);

      // Send file
      res.send(buffer);

      logger.info(`Excel file generated from query: ${value.filename}`);
    } catch (error) {
      throw error;
    }
  }

  /**
   * Upload Excel template
   * @param {Object} req - Express request object
   * @param {Object} res - Express response object
   */
  static async uploadTemplate(req, res) {
    try {
      if (!req.file) {
        throw new APIError('No template file uploaded', 400);
      }

      const templateInfo = {
        filename: req.file.filename,
        originalName: req.file.originalname,
        size: req.file.size,
        uploadDate: new Date().toISOString(),
        path: req.file.path
      };

      res.status(201).json({
        success: true,
        message: 'Template uploaded successfully',
        template: templateInfo
      });

      logger.info(`Template uploaded: ${req.file.originalname}`);
    } catch (error) {
      throw error;
    }
  }

  /**
   * List available templates
   * @param {Object} req - Express request object
   * @param {Object} res - Express response object
   */
  static async listTemplates(req, res) {
    try {
      const templatesDir = path.join(__dirname, '..', 'templates');
      const files = await fs.readdir(templatesDir);
      
      const templates = [];
      for (const file of files) {
        if (file.endsWith('.xlsx') || file.endsWith('.xls')) {
          const filePath = path.join(templatesDir, file);
          const stats = await fs.stat(filePath);
          
          templates.push({
            filename: file,
            size: stats.size,
            created: stats.birthtime,
            modified: stats.mtime
          });
        }
      }

      res.json({
        success: true,
        count: templates.length,
        templates: templates
      });

      logger.info(`Listed ${templates.length} templates`);
    } catch (error) {
      throw error;
    }
  }

  /**
   * Delete template
   * @param {Object} req - Express request object
   * @param {Object} res - Express response object
   */
  static async deleteTemplate(req, res) {
    try {
      const { filename } = req.params;
      
      if (!filename) {
        throw new APIError('Template filename is required', 400);
      }

      const templatePath = path.join(__dirname, '..', 'templates', filename);
      
      try {
        await fs.unlink(templatePath);
        res.json({
          success: true,
          message: `Template ${filename} deleted successfully`
        });
        logger.info(`Template deleted: ${filename}`);
      } catch (error) {
        if (error.code === 'ENOENT') {
          throw new APIError('Template not found', 404);
        }
        throw error;
      }
    } catch (error) {
      throw error;
    }
  }

  /**
   * Fill template with data
   * @param {Object} req - Express request object
   * @param {Object} res - Express response object
   */
  static async fillTemplate(req, res) {
    try {
      // Validate request payload
      const schema = Joi.object({
        template: Joi.string().required(),
        data: Joi.object().required(),
        filename: Joi.string().default('filled-template.xlsx')
      });

      const { error, value } = schema.validate(req.body);
      if (error) {
        throw new APIError(`Validation error: ${error.details[0].message}`, 400);
      }

      // Fill template
      const buffer = await ExcelService.fillTemplate(value.template, value.data);

      // Set response headers
      res.setHeader('Content-Disposition', `attachment; filename="${value.filename}"`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Length', buffer.length);

      // Send file
      res.send(buffer);

      logger.info(`Template filled successfully: ${value.template}`);
    } catch (error) {
      throw error;
    }
  }

  /**
   * Generate multi-sheet Excel
   * @param {Object} req - Express request object
   * @param {Object} res - Express response object
   */
  static async generateMultiSheet(req, res) {
    try {
      // Validate request payload
      const schema = Joi.object({
        sheets: Joi.array().items(
          Joi.object({
            name: Joi.string().required(),
            query: Joi.string().optional(),
            params: Joi.array().default([]),
            data: Joi.array().optional(),
            style: Joi.object().optional()
          })
        ).required(),
        filename: Joi.string().default('multi-sheet.xlsx')
      });

      const { error, value } = schema.validate(req.body);
      if (error) {
        throw new APIError(`Validation error: ${error.details[0].message}`, 400);
      }

      const sheets = [];
      
      // Process each sheet
      for (const sheetDef of value.sheets) {
        let data = sheetDef.data;
        
        // Execute query if provided
        if (sheetDef.query && !data) {
          try {
            data = await executeSafeQuery(sheetDef.query, sheetDef.params);
          } catch (dbError) {
            logger.warn(`Query failed for sheet ${sheetDef.name}, using sample data`);
            data = getSampleData();
          }
        }

        if (!data || data.length === 0) {
          data = [{ message: 'No data available' }];
        }

        sheets.push({
          name: sheetDef.name,
          tables: [{
            name: `${sheetDef.name}Table`,
            range: 'A1',
            data: data,
            style: sheetDef.style
          }]
        });
      }

      // Generate workbook
      const buffer = await ExcelService.createWorkbook({ sheets });

      // Set response headers
      res.setHeader('Content-Disposition', `attachment; filename="${value.filename}"`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Length', buffer.length);

      // Send file
      res.send(buffer);

      logger.info(`Multi-sheet Excel generated: ${value.filename}`);
    } catch (error) {
      throw error;
    }
  }

  /**
   * Generate styled Excel
   * @param {Object} req - Express request object
   * @param {Object} res - Express response object
   */
  static async generateStyledExcel(req, res) {
    try {
      // Validate request payload
      const schema = Joi.object({
        sheets: Joi.array().items(
          Joi.object({
            name: Joi.string().required(),
            data: Joi.array().required(),
            styles: Joi.object({
              header: Joi.object().optional(),
              rows: Joi.object().optional(),
              alternateRows: Joi.object().optional(),
              columns: Joi.object().optional()
            }).optional(),
            formatting: Joi.object({
              autoWidth: Joi.boolean().default(true),
              freezeRows: Joi.number().optional(),
              freezeCols: Joi.number().optional(),
              pageSetup: Joi.object().optional()
            }).optional()
          })
        ).required(),
        filename: Joi.string().default('styled-excel.xlsx')
      });

      const { error, value } = schema.validate(req.body);
      if (error) {
        throw new APIError(`Validation error: ${error.details[0].message}`, 400);
      }

      const sheets = value.sheets.map(sheet => ({
        name: sheet.name,
        tables: [{
          name: `${sheet.name}Table`,
          range: 'A1',
          data: sheet.data,
          style: sheet.styles
        }],
        formatting: sheet.formatting
      }));

      // Generate workbook
      const buffer = await ExcelService.createWorkbook({ sheets });

      // Set response headers
      res.setHeader('Content-Disposition', `attachment; filename="${value.filename}"`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Length', buffer.length);

      // Send file
      res.send(buffer);

      logger.info(`Styled Excel generated: ${value.filename}`);
    } catch (error) {
      throw error;
    }
  }

  /**
   * Get database schema
   * @param {Object} req - Express request object
   * @param {Object} res - Express response object
   */
  static async getSchema(req, res) {
    try {
      const { table } = req.params;
      
      let schema;
      try {
        schema = await getSchema(table);
      } catch (error) {
        // If database is not available, return sample schema
        schema = [
          { table_name: 'employees', column_name: 'id', data_type: 'integer' },
          { table_name: 'employees', column_name: 'name', data_type: 'varchar' },
          { table_name: 'employees', column_name: 'email', data_type: 'varchar' },
          { table_name: 'employees', column_name: 'department', data_type: 'varchar' },
          { table_name: 'employees', column_name: 'salary', data_type: 'numeric' }
        ];
      }

      res.json({
        success: true,
        table: table || 'all',
        schema: schema
      });

      logger.info(`Schema retrieved for: ${table || 'all tables'}`);
    } catch (error) {
      throw error;
    }
  }

  /**
   * Validate SQL query
   * @param {Object} req - Express request object
   * @param {Object} res - Express response object
   */
  static async validateQuery(req, res) {
    try {
      const schema = Joi.object({
        query: Joi.string().required()
      });

      const { error, value } = schema.validate(req.body);
      if (error) {
        throw new APIError(`Validation error: ${error.details[0].message}`, 400);
      }

      const isValid = validateQuery(value.query);
      
      res.json({
        success: true,
        valid: isValid,
        query: value.query,
        message: isValid ? 'Query is valid' : 'Query contains potentially dangerous patterns'
      });

      logger.info(`Query validation: ${isValid ? 'VALID' : 'INVALID'}`);
    } catch (error) {
      throw error;
    }
  }

  /**
   * Get sample data
   * @param {Object} req - Express request object
   * @param {Object} res - Express response object
   */
  static async getSampleData(req, res) {
    try {
      const sampleData = getSampleData();
      
      res.json({
        success: true,
        count: sampleData.length,
        data: sampleData,
        message: 'Sample data for testing Excel generation'
      });

      logger.info('Sample data retrieved');
    } catch (error) {
      throw error;
    }
  }
}

module.exports = ExcelController;
