const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const { logger, APIError } = require('../middleware/errorHandler');

/**
 * Excel Service - Core Excel generation and manipulation functions
 */
class ExcelService {
  
  /**
   * Create a new workbook from JSON payload
   * @param {Object} payload - Excel generation payload
   * @returns {Promise<Buffer>} - Excel file buffer
   */
  static async createWorkbook(payload) {
    try {
      let workbook = new ExcelJS.Workbook();
      
      // Set workbook properties
      workbook.creator = 'Excel API Service';
      workbook.lastModifiedBy = 'Excel API Service';
      workbook.created = new Date();
      workbook.modified = new Date();

      // Load template if specified
      if (payload.template) {
        await this.loadTemplate(workbook, payload.template);
      }

      // Process each sheet
      for (const sheetDef of payload.sheets || []) {
        await this.processSheet(workbook, sheetDef);
      }

      // Generate buffer
      const buffer = await workbook.xlsx.writeBuffer();
      logger.info(`Excel workbook created successfully with ${payload.sheets?.length || 0} sheets`);
      
      return buffer;
    } catch (error) {
      logger.error('Excel workbook creation failed:', error);
      throw new APIError(`Excel generation failed: ${error.message}`, 422);
    }
  }

  /**
   * Load Excel template
   * @param {ExcelJS.Workbook} workbook - Workbook instance
   * @param {string} templateName - Template filename
   */
  static async loadTemplate(workbook, templateName) {
    try {
      const templatePath = path.join(__dirname, '..', 'templates', templateName);
      await workbook.xlsx.readFile(templatePath);
      logger.info(`Template loaded: ${templateName}`);
    } catch (error) {
      throw new APIError(`Template loading failed: ${error.message}`, 404);
    }
  }

  /**
   * Process individual sheet
   * @param {ExcelJS.Workbook} workbook - Workbook instance
   * @param {Object} sheetDef - Sheet definition
   */
  static async processSheet(workbook, sheetDef) {
    try {
      // Get or create worksheet
      let worksheet = workbook.getWorksheet(sheetDef.name) || workbook.addWorksheet(sheetDef.name);

      // Process individual cells
      if (sheetDef.cells) {
        for (const cellDef of sheetDef.cells) {
          await this.processCell(worksheet, cellDef);
        }
      }

      // Process ranges
      if (sheetDef.ranges) {
        for (const rangeDef of sheetDef.ranges) {
          await this.processRange(worksheet, rangeDef);
        }
      }

      // Process tables
      if (sheetDef.tables) {
        for (const tableDef of sheetDef.tables) {
          await this.processTable(worksheet, tableDef);
        }
      }

      // Apply sheet-level formatting
      if (sheetDef.formatting) {
        await this.applySheetFormatting(worksheet, sheetDef.formatting);
      }

      logger.info(`Sheet processed: ${sheetDef.name}`);
    } catch (error) {
      throw new APIError(`Sheet processing failed: ${error.message}`, 422);
    }
  }

  /**
   * Process individual cell
   * @param {ExcelJS.Worksheet} worksheet - Worksheet instance
   * @param {Object} cellDef - Cell definition
   */
  static async processCell(worksheet, cellDef) {
    try {
      const cell = worksheet.getCell(cellDef.cell);

      // Set value or formula
      if (cellDef.formula) {
        cell.value = { formula: cellDef.formula, result: cellDef.value };
      } else {
        cell.value = cellDef.value;
      }

      // Apply styling
      if (cellDef.style) {
        this.applyCellStyle(cell, cellDef.style);
      }

      // Set data type and formatting
      if (cellDef.dataType) {
        this.applyCellDataType(cell, cellDef.dataType, cellDef.format);
      }

    } catch (error) {
      throw new APIError(`Cell processing failed for ${cellDef.cell}: ${error.message}`, 422);
    }
  }

  /**
   * Process range of cells
   * @param {ExcelJS.Worksheet} worksheet - Worksheet instance
   * @param {Object} rangeDef - Range definition
   */
  static async processRange(worksheet, rangeDef) {
    try {
      const range = worksheet.getCell(rangeDef.range);
      
      // Set data for range
      if (rangeDef.data && Array.isArray(rangeDef.data)) {
        const startCell = rangeDef.range.split(':')[0];
        const startRow = parseInt(startCell.match(/\d+/)[0]);
        const startCol = startCell.match(/[A-Z]+/)[0];
        
        rangeDef.data.forEach((rowData, rowIndex) => {
          if (Array.isArray(rowData)) {
            rowData.forEach((cellValue, colIndex) => {
              const cellAddress = this.getCellAddress(startCol, startRow + rowIndex, colIndex);
              const cell = worksheet.getCell(cellAddress);
              cell.value = cellValue;
              
              if (rangeDef.style) {
                this.applyCellStyle(cell, rangeDef.style);
              }
            });
          }
        });
      }

    } catch (error) {
      throw new APIError(`Range processing failed: ${error.message}`, 422);
    }
  }

  /**
   * Process table
   * @param {ExcelJS.Worksheet} worksheet - Worksheet instance
   * @param {Object} tableDef - Table definition
   */
  static async processTable(worksheet, tableDef) {
    try {
      if (!tableDef.data || !Array.isArray(tableDef.data) || tableDef.data.length === 0) {
        throw new Error('Table data is required and must be a non-empty array');
      }

      const startCell = tableDef.range ? tableDef.range.split(':')[0] : 'A1';
      const headers = Object.keys(tableDef.data[0]);
      
      // Add headers
      const startRow = parseInt(startCell.match(/\d+/)[0]);
      const startColLetter = startCell.match(/[A-Z]+/)[0];
      
      headers.forEach((header, index) => {
        const cellAddress = this.getCellAddress(startColLetter, startRow, index);
        const cell = worksheet.getCell(cellAddress);
        cell.value = header;
        
        // Apply header style
        if (tableDef.style?.header) {
          this.applyCellStyle(cell, tableDef.style.header);
        } else {
          // Default header style
          this.applyCellStyle(cell, {
            font: { bold: true, color: { argb: 'FFFFFFFF' } },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF366092' } },
            alignment: { horizontal: 'center', vertical: 'middle' }
          });
        }
      });

      // Add data rows
      tableDef.data.forEach((rowData, rowIndex) => {
        headers.forEach((header, colIndex) => {
          const cellAddress = this.getCellAddress(startColLetter, startRow + rowIndex + 1, colIndex);
          const cell = worksheet.getCell(cellAddress);
          cell.value = rowData[header];
          
          // Apply row style
          if (tableDef.style?.rows) {
            this.applyCellStyle(cell, tableDef.style.rows);
          }
          
          // Apply alternate row style
          if (rowIndex % 2 === 1 && tableDef.style?.alternateRows) {
            this.applyCellStyle(cell, tableDef.style.alternateRows);
          }
        });
      });

      // Create Excel table if name is provided
      if (tableDef.name) {
        const endRow = startRow + tableDef.data.length;
        const endCol = this.getCellAddress(startColLetter, 1, headers.length - 1).match(/[A-Z]+/)[0];
        const tableRange = `${startCell}:${endCol}${endRow}`;
        
        worksheet.addTable({
          name: tableDef.name,
          ref: tableRange,
          headerRow: true,
          columns: headers.map(header => ({ name: header })),
          rows: tableDef.data.map(row => headers.map(header => row[header]))
        });
      }

    } catch (error) {
      throw new APIError(`Table processing failed: ${error.message}`, 422);
    }
  }

  /**
   * Apply cell styling
   * @param {ExcelJS.Cell} cell - Cell instance
   * @param {Object} style - Style definition
   */
  static applyCellStyle(cell, style) {
    try {
      if (style.font) {
        cell.font = {
          name: style.font.name || 'Calibri',
          size: style.font.size || 11,
          bold: style.font.bold || false,
          italic: style.font.italic || false,
          underline: style.font.underline || false,
          color: style.font.color ? { argb: this.normalizeColor(style.font.color) } : undefined
        };
      }

      if (style.fill || style.backgroundColor) {
        const bgColor = style.backgroundColor || style.fill?.fgColor;
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: this.normalizeColor(bgColor) }
        };
      }

      if (style.alignment) {
        cell.alignment = {
          horizontal: style.alignment.horizontal || 'left',
          vertical: style.alignment.vertical || 'top',
          wrapText: style.alignment.wrapText || false,
          indent: style.alignment.indent || 0
        };
      }

      if (style.border) {
        cell.border = {
          top: style.border.top || { style: 'thin' },
          left: style.border.left || { style: 'thin' },
          bottom: style.border.bottom || { style: 'thin' },
          right: style.border.right || { style: 'thin' }
        };
      }

      if (style.numFmt) {
        cell.numFmt = style.numFmt;
      }

    } catch (error) {
      logger.warn(`Style application failed: ${error.message}`);
    }
  }

  /**
   * Apply cell data type and formatting
   * @param {ExcelJS.Cell} cell - Cell instance
   * @param {string} dataType - Data type
   * @param {string} format - Number format
   */
  static applyCellDataType(cell, dataType, format) {
    try {
      switch (dataType.toLowerCase()) {
        case 'number':
          cell.numFmt = format || '0.00';
          break;
        case 'currency':
          cell.numFmt = format || '$#,##0.00';
          break;
        case 'percentage':
          cell.numFmt = format || '0.00%';
          break;
        case 'date':
          cell.numFmt = format || 'mm/dd/yyyy';
          break;
        case 'datetime':
          cell.numFmt = format || 'mm/dd/yyyy hh:mm:ss';
          break;
        case 'time':
          cell.numFmt = format || 'hh:mm:ss';
          break;
        case 'text':
          cell.numFmt = '@';
          break;
        default:
          // No specific formatting
          break;
      }
    } catch (error) {
      logger.warn(`Data type application failed: ${error.message}`);
    }
  }

  /**
   * Apply sheet-level formatting
   * @param {ExcelJS.Worksheet} worksheet - Worksheet instance
   * @param {Object} formatting - Formatting options
   */
  static async applySheetFormatting(worksheet, formatting) {
    try {
      // Auto-width columns
      if (formatting.autoWidth) {
        worksheet.columns.forEach(column => {
          let maxLength = 0;
          column.eachCell({ includeEmpty: true }, (cell) => {
            const columnLength = cell.value ? cell.value.toString().length : 10;
            if (columnLength > maxLength) {
              maxLength = columnLength;
            }
          });
          column.width = Math.min(maxLength + 2, 50); // Max width of 50
        });
      }

      // Freeze panes
      if (formatting.freezeRows || formatting.freezeCols) {
        const freezeRow = (formatting.freezeRows || 0) + 1;
        const freezeCol = (formatting.freezeCols || 0) + 1;
        worksheet.views = [{
          state: 'frozen',
          xSplit: formatting.freezeCols || 0,
          ySplit: formatting.freezeRows || 0,
          topLeftCell: this.getCellAddress('A', freezeRow, freezeCol - 1)
        }];
      }

      // Page setup
      if (formatting.pageSetup) {
        worksheet.pageSetup = {
          orientation: formatting.pageSetup.orientation || 'portrait',
          paperSize: formatting.pageSetup.paperSize || 9, // A4
          margins: formatting.pageSetup.margins || {
            left: 0.7, right: 0.7, top: 0.75, bottom: 0.75,
            header: 0.3, footer: 0.3
          }
        };
      }

    } catch (error) {
      logger.warn(`Sheet formatting failed: ${error.message}`);
    }
  }

  /**
   * Fill template with data
   * @param {string} templateName - Template filename
   * @param {Object} data - Data to fill
   * @returns {Promise<Buffer>} - Excel file buffer
   */
  static async fillTemplate(templateName, data) {
    try {
      const workbook = new ExcelJS.Workbook();
      await this.loadTemplate(workbook, templateName);

      // Process each worksheet in the template
      workbook.eachSheet((worksheet) => {
        this.fillWorksheetTemplate(worksheet, data);
      });

      const buffer = await workbook.xlsx.writeBuffer();
      logger.info(`Template filled successfully: ${templateName}`);
      
      return buffer;
    } catch (error) {
      logger.error('Template filling failed:', error);
      throw new APIError(`Template filling failed: ${error.message}`, 422);
    }
  }

  /**
   * Fill worksheet template with data
   * @param {ExcelJS.Worksheet} worksheet - Worksheet instance
   * @param {Object} data - Data to fill
   */
  static fillWorksheetTemplate(worksheet, data) {
    try {
      worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          if (cell.value && typeof cell.value === 'string') {
            // Replace single value placeholders like {{name}}
            let cellValue = cell.value;
            const singleValueRegex = /\{\{(\w+)\}\}/g;
            cellValue = cellValue.replace(singleValueRegex, (match, key) => {
              return data[key] !== undefined ? data[key] : match;
            });

            // Handle array data placeholders like {{#items}}
            const arrayRegex = /\{\{#(\w+)\}\}/g;
            const arrayMatch = arrayRegex.exec(cellValue);
            if (arrayMatch && data[arrayMatch[1]] && Array.isArray(data[arrayMatch[1]])) {
              const arrayData = data[arrayMatch[1]];
              const startRow = rowNumber;
              
              // Insert rows for array data
              arrayData.forEach((item, index) => {
                if (index > 0) {
                  worksheet.insertRow(startRow + index, []);
                }
                
                const currentRow = worksheet.getRow(startRow + index);
                currentRow.eachCell((arrayCell, arrayCellNumber) => {
                  if (arrayCell.value && typeof arrayCell.value === 'string') {
                    let arrayValue = arrayCell.value.replace(/\{\{#\w+\}\}/, '');
                    const itemRegex = /\{\{(\w+)\}\}/g;
                    arrayValue = arrayValue.replace(itemRegex, (itemMatch, itemKey) => {
                      return item[itemKey] !== undefined ? item[itemKey] : itemMatch;
                    });
                    arrayCell.value = arrayValue;
                  }
                });
              });
            } else {
              cell.value = cellValue;
            }
          }
        });
      });
    } catch (error) {
      logger.warn(`Worksheet template filling failed: ${error.message}`);
    }
  }

  /**
   * Get cell address from column letter and row number
   * @param {string} startCol - Starting column letter
   * @param {number} row - Row number
   * @param {number} colOffset - Column offset
   * @returns {string} - Cell address
   */
  static getCellAddress(startCol, row, colOffset = 0) {
    const colCode = startCol.charCodeAt(0) - 65 + colOffset;
    const colLetter = String.fromCharCode(65 + (colCode % 26));
    return `${colLetter}${row}`;
  }

  /**
   * Normalize color format
   * @param {string} color - Color value
   * @returns {string} - Normalized ARGB color
   */
  static normalizeColor(color) {
    if (!color) return 'FF000000';
    
    // Remove # if present
    color = color.replace('#', '');
    
    // Add alpha channel if not present
    if (color.length === 6) {
      color = 'FF' + color;
    }
    
    return color.toUpperCase();
  }

  /**
   * Generate Excel from query results
   * @param {Array} data - Query results
   * @param {Object} options - Generation options
   * @returns {Promise<Buffer>} - Excel file buffer
   */
  static async generateFromQueryResults(data, options = {}) {
    try {
      const payload = {
        sheets: [{
          name: options.sheetName || 'QueryResults',
          tables: [{
            name: options.tableName || 'DataTable',
            range: options.startCell || 'A1',
            data: data,
            style: options.style || {
              header: {
                font: { bold: true, color: { argb: 'FFFFFFFF' } },
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF366092' } }
              }
            }
          }]
        }]
      };

      return await this.createWorkbook(payload);
    } catch (error) {
      throw new APIError(`Query result Excel generation failed: ${error.message}`, 422);
    }
  }
}

module.exports = ExcelService;
