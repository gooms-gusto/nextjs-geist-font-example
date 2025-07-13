# Excel Generation API Documentation

## Overview

The Excel Generation API is a comprehensive Node.js backend service that allows users to generate Excel files from database queries with advanced features including:

- **Flexible Cell Addressing**: Address data to specific cells, ranges, columns, rows, or entire sheets
- **Formula Support**: Set Excel formulas in cells with automatic calculation
- **Database Integration**: Create tables directly from database query results
- **Template Support**: Use existing Excel templates with parameter filling
- **Advanced Styling**: Apply background colors, font colors, borders, and formatting
- **Data Type Configuration**: Set cell data types (number, currency, date, etc.) with custom formatting
- **Multi-sheet Support**: Generate workbooks with multiple sheets from different data sources

## Base URL

```
http://localhost:9000/api/excel
```

## Authentication

Currently, no authentication is required. This can be added based on your security requirements.

## Endpoints

### 1. Health Check

**GET** `/health`

Check if the API server is running.

**Response:**
```json
{
  "status": "OK",
  "message": "Excel API Server is running",
  "timestamp": "2024-01-15T10:30:00.000Z",
  "version": "1.0.0"
}
```

### 2. API Documentation

**GET** `/api/docs`

Get API endpoint information.

**Response:**
```json
{
  "title": "Excel Generation API",
  "version": "1.0.0",
  "description": "Advanced Excel generation API with database integration, templates, and styling",
  "endpoints": {
    "POST /api/excel/generate": "Generate Excel file from JSON payload",
    "POST /api/excel/template": "Upload Excel template",
    "GET /api/excel/templates": "List available templates"
  }
}
```

### 3. Generate Excel File

**POST** `/api/excel/generate`

Generate an Excel file from a comprehensive JSON payload with support for multiple sheets, cells, ranges, tables, formulas, and styling.

**Request Body:**
```json
{
  "template": "optional-template.xlsx",
  "filename": "my-report.xlsx",
  "sheets": [
    {
      "name": "Dashboard",
      "cells": [
        {
          "cell": "A1",
          "value": "Sales Report",
          "style": {
            "font": {
              "bold": true,
              "size": 16,
              "color": "#FFFFFF"
            },
            "fill": {
              "fgColor": "#366092"
            },
            "alignment": {
              "horizontal": "center",
              "vertical": "middle"
            }
          }
        },
        {
          "cell": "B2",
          "value": 1000,
          "dataType": "currency",
          "format": "$#,##0.00"
        },
        {
          "cell": "C2",
          "formula": "=SUM(D2:D10)",
          "dataType": "number"
        }
      ],
      "ranges": [
        {
          "range": "A3:C5",
          "data": [
            ["Product", "Quantity", "Price"],
            ["Widget A", 100, 25.50],
            ["Widget B", 150, 30.00]
          ],
          "style": {
            "border": {
              "top": {"style": "thin"},
              "bottom": {"style": "thin"},
              "left": {"style": "thin"},
              "right": {"style": "thin"}
            }
          }
        }
      ],
      "tables": [
        {
          "name": "SalesData",
          "range": "A10:D20",
          "data": [
            {"Product": "Widget A", "Q1": 1000, "Q2": 1200, "Q3": 1100},
            {"Product": "Widget B", "Q1": 800, "Q2": 900, "Q3": 950}
          ],
          "style": {
            "header": {
              "font": {"bold": true, "color": "#FFFFFF"},
              "fill": {"fgColor": "#366092"}
            },
            "rows": {
              "alignment": {"horizontal": "center"}
            },
            "alternateRows": {
              "fill": {"fgColor": "#F2F2F2"}
            }
          }
        }
      ],
      "formatting": {
        "autoWidth": true,
        "freezeRows": 1,
        "freezeCols": 1,
        "pageSetup": {
          "orientation": "landscape",
          "paperSize": 9
        }
      }
    }
  ]
}
```

**Response:**
- **Content-Type**: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
- **Content-Disposition**: `attachment; filename="my-report.xlsx"`
- Returns the generated Excel file as a binary download

### 4. Generate Excel from Database Query

**POST** `/api/excel/query`

Generate an Excel file directly from a database query with automatic table creation.

**Request Body:**
```json
{
  "query": "SELECT id, name, email, department, salary FROM employees WHERE department = ?",
  "params": ["Engineering"],
  "sheetName": "Engineering Team",
  "filename": "engineering-report.xlsx",
  "startCell": "A1",
  "includeHeaders": true,
  "style": {
    "header": {
      "font": {"bold": true, "color": "#FFFFFF"},
      "fill": {"fgColor": "#366092"},
      "alignment": {"horizontal": "center"}
    },
    "rows": {
      "alignment": {"horizontal": "left"}
    },
    "alternateRows": {
      "fill": {"fgColor": "#F8F9FA"}
    }
  }
}
```

**Response:**
Returns the generated Excel file with query results formatted as a table.

### 5. Upload Excel Template

**POST** `/api/excel/template`

Upload an Excel template file for later use in template-based generation.

**Request:**
- **Content-Type**: `multipart/form-data`
- **Field**: `template` (file upload)

**Response:**
```json
{
  "success": true,
  "message": "Template uploaded successfully",
  "template": {
    "filename": "template-1642234567890-123456789.xlsx",
    "originalName": "sales-template.xlsx",
    "size": 15360,
    "uploadDate": "2024-01-15T10:30:00.000Z",
    "path": "/templates/template-1642234567890-123456789.xlsx"
  }
}
```

### 6. List Available Templates

**GET** `/api/excel/templates`

Get a list of all available Excel templates.

**Response:**
```json
{
  "success": true,
  "count": 2,
  "templates": [
    {
      "filename": "sales-template.xlsx",
      "size": 15360,
      "created": "2024-01-15T10:00:00.000Z",
      "modified": "2024-01-15T10:00:00.000Z"
    },
    {
      "filename": "report-template.xlsx",
      "size": 23040,
      "created": "2024-01-14T15:30:00.000Z",
      "modified": "2024-01-14T15:30:00.000Z"
    }
  ]
}
```

### 7. Delete Template

**DELETE** `/api/excel/templates/:filename`

Delete a specific template file.

**Parameters:**
- `filename` (path parameter): Name of the template file to delete

**Response:**
```json
{
  "success": true,
  "message": "Template sales-template.xlsx deleted successfully"
}
```

### 8. Fill Template with Data

**POST** `/api/excel/template/fill`

Fill an existing template with data using placeholder replacement.

**Request Body:**
```json
{
  "template": "sales-template.xlsx",
  "filename": "filled-sales-report.xlsx",
  "data": {
    "companyName": "Acme Corporation",
    "reportDate": "2024-01-15",
    "totalSales": 125000,
    "salesRep": "John Doe",
    "items": [
      {"product": "Widget A", "quantity": 100, "price": 25.50, "total": 2550},
      {"product": "Widget B", "quantity": 75, "price": 30.00, "total": 2250}
    ]
  }
}
```

**Template Placeholders:**
- Single values: `{{companyName}}`, `{{reportDate}}`, `{{totalSales}}`
- Array data: `{{#items}}` with item properties like `{{product}}`, `{{quantity}}`

**Response:**
Returns the filled Excel template as a download.

### 9. Generate Multi-Sheet Excel

**POST** `/api/excel/multi-sheet`

Generate an Excel workbook with multiple sheets from different data sources or queries.

**Request Body:**
```json
{
  "filename": "multi-department-report.xlsx",
  "sheets": [
    {
      "name": "Engineering",
      "query": "SELECT * FROM employees WHERE department = ?",
      "params": ["Engineering"],
      "style": {
        "header": {"font": {"bold": true}, "fill": {"fgColor": "#366092"}}
      }
    },
    {
      "name": "Sales",
      "data": [
        {"name": "Alice", "sales": 50000, "target": 45000},
        {"name": "Bob", "sales": 60000, "target": 55000}
      ],
      "style": {
        "header": {"font": {"bold": true}, "fill": {"fgColor": "#28A745"}}
      }
    }
  ]
}
```

**Response:**
Returns a multi-sheet Excel workbook.

### 10. Generate Styled Excel

**POST** `/api/excel/styled`

Generate Excel with advanced styling options and formatting.

**Request Body:**
```json
{
  "filename": "styled-report.xlsx",
  "sheets": [
    {
      "name": "Styled Data",
      "data": [
        {"Product": "Widget A", "Sales": 1000, "Target": 900, "Performance": "110%"},
        {"Product": "Widget B", "Sales": 800, "Target": 850, "Performance": "94%"}
      ],
      "styles": {
        "header": {
          "font": {"bold": true, "size": 12, "color": "#FFFFFF"},
          "fill": {"fgColor": "#366092"},
          "alignment": {"horizontal": "center", "vertical": "middle"},
          "border": {
            "top": {"style": "thick"},
            "bottom": {"style": "thick"},
            "left": {"style": "thick"},
            "right": {"style": "thick"}
          }
        },
        "rows": {
          "font": {"size": 10},
          "alignment": {"horizontal": "left"},
          "border": {
            "top": {"style": "thin"},
            "bottom": {"style": "thin"},
            "left": {"style": "thin"},
            "right": {"style": "thin"}
          }
        },
        "alternateRows": {
          "fill": {"fgColor": "#F8F9FA"}
        }
      },
      "formatting": {
        "autoWidth": true,
        "freezeRows": 1,
        "pageSetup": {
          "orientation": "portrait",
          "margins": {
            "left": 0.7,
            "right": 0.7,
            "top": 0.75,
            "bottom": 0.75
          }
        }
      }
    }
  ]
}
```

**Response:**
Returns a styled Excel workbook.

### 11. Get Database Schema

**GET** `/api/excel/schema/:table?`

Get database schema information for all tables or a specific table.

**Parameters:**
- `table` (optional): Specific table name

**Response:**
```json
{
  "success": true,
  "table": "employees",
  "schema": [
    {"column_name": "id", "data_type": "integer", "is_nullable": "NO"},
    {"column_name": "name", "data_type": "varchar", "is_nullable": "NO"},
    {"column_name": "email", "data_type": "varchar", "is_nullable": "YES"},
    {"column_name": "department", "data_type": "varchar", "is_nullable": "YES"},
    {"column_name": "salary", "data_type": "numeric", "is_nullable": "YES"}
  ]
}
```

### 12. Validate SQL Query

**POST** `/api/excel/validate-query`

Validate a SQL query for safety without executing it.

**Request Body:**
```json
{
  "query": "SELECT id, name, email FROM employees WHERE department = 'Engineering'"
}
```

**Response:**
```json
{
  "success": true,
  "valid": true,
  "query": "SELECT id, name, email FROM employees WHERE department = 'Engineering'",
  "message": "Query is valid"
}
```

### 13. Get Sample Data

**GET** `/api/excel/sample-data`

Get sample data for testing Excel generation features.

**Response:**
```json
{
  "success": true,
  "count": 5,
  "data": [
    {"id": 1, "name": "John Doe", "email": "john@example.com", "department": "Engineering", "salary": 75000},
    {"id": 2, "name": "Jane Smith", "email": "jane@example.com", "department": "Marketing", "salary": 65000}
  ],
  "message": "Sample data for testing Excel generation"
}
```

## Data Types and Formatting

### Supported Data Types

- **text**: Plain text (format: `@`)
- **number**: Numeric values (format: `0.00`)
- **currency**: Currency values (format: `$#,##0.00`)
- **percentage**: Percentage values (format: `0.00%`)
- **date**: Date values (format: `mm/dd/yyyy`)
- **datetime**: Date and time values (format: `mm/dd/yyyy hh:mm:ss`)
- **time**: Time values (format: `hh:mm:ss`)

### Style Properties

#### Font Styling
```json
{
  "font": {
    "name": "Calibri",
    "size": 11,
    "bold": true,
    "italic": false,
    "underline": false,
    "color": "#000000"
  }
}
```

#### Fill/Background
```json
{
  "fill": {
    "type": "pattern",
    "pattern": "solid",
    "fgColor": "#FFFFFF"
  }
}
```

#### Alignment
```json
{
  "alignment": {
    "horizontal": "left|center|right",
    "vertical": "top|middle|bottom",
    "wrapText": true,
    "indent": 0
  }
}
```

#### Borders
```json
{
  "border": {
    "top": {"style": "thin|thick|medium"},
    "bottom": {"style": "thin|thick|medium"},
    "left": {"style": "thin|thick|medium"},
    "right": {"style": "thin|thick|medium"}
  }
}
```

## Error Handling

All endpoints return consistent error responses:

```json
{
  "success": false,
  "message": "Error description",
  "timestamp": "2024-01-15T10:30:00.000Z",
  "details": "Additional error details (in development mode)"
}
```

### Common HTTP Status Codes

- **200**: Success
- **201**: Created (for uploads)
- **400**: Bad Request (validation errors)
- **404**: Not Found (template/resource not found)
- **413**: Payload Too Large (file size exceeded)
- **422**: Unprocessable Entity (Excel processing errors)
- **500**: Internal Server Error

## Usage Examples

### Example 1: Simple Excel Generation

```javascript
const response = await fetch('http://localhost:9000/api/excel/generate', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json'
  },
  body: JSON.stringify({
    filename: 'simple-report.xlsx',
    sheets: [{
      name: 'Data',
      cells: [
        { cell: 'A1', value: 'Name', style: { font: { bold: true } } },
        { cell: 'B1', value: 'Score', style: { font: { bold: true } } },
        { cell: 'A2', value: 'John', dataType: 'text' },
        { cell: 'B2', value: 95, dataType: 'number' }
      ]
    }]
  })
});

const blob = await response.blob();
// Handle file download
```

### Example 2: Database Query to Excel

```javascript
const response = await fetch('http://localhost:9000/api/excel/query', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json'
  },
  body: JSON.stringify({
    query: 'SELECT * FROM employees WHERE salary > ?',
    params: [50000],
    filename: 'high-earners.xlsx',
    sheetName: 'High Earners',
    style: {
      header: {
        font: { bold: true, color: '#FFFFFF' },
        fill: { fgColor: '#366092' }
      }
    }
  })
});
```

### Example 3: Template Filling

```javascript
// First upload a template
const formData = new FormData();
formData.append('template', templateFile);

await fetch('http://localhost:9000/api/excel/template', {
  method: 'POST',
  body: formData
});

// Then fill the template
const response = await fetch('http://localhost:9000/api/excel/template/fill', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json'
  },
  body: JSON.stringify({
    template: 'uploaded-template.xlsx',
    data: {
      title: 'Monthly Report',
      date: '2024-01-15',
      items: [
        { product: 'Widget A', sales: 1000 },
        { product: 'Widget B', sales: 800 }
      ]
    }
  })
});
```

## Environment Configuration

Configure the API using environment variables in `.env`:

```env
# Server Configuration
PORT=9000
NODE_ENV=development

# Database Configuration
DATABASE_URL=postgresql://user:password@localhost:5432/database

# File Upload Configuration
MAX_FILE_SIZE=10485760
ALLOWED_FILE_TYPES=.xlsx,.xls

# CORS Configuration
CORS_ORIGIN=http://localhost:8000
```

## Security Considerations

1. **SQL Injection Prevention**: All database queries use parameterized statements
2. **File Upload Security**: File type and size validation
3. **Query Validation**: Dangerous SQL patterns are blocked
4. **Path Traversal Protection**: Template file paths are restricted
5. **CORS Configuration**: Configurable cross-origin access

## Performance Tips

1. **Large Datasets**: For large datasets, consider pagination or streaming
2. **Complex Styling**: Minimize complex styling for better performance
3. **Template Caching**: Templates are cached for repeated use
4. **Memory Management**: Large Excel files are streamed to prevent memory issues

## Support and Troubleshooting

### Common Issues

1. **Template Not Found**: Ensure template is uploaded and filename is correct
2. **Database Connection**: Check DATABASE_URL configuration
3. **File Size Limits**: Adjust MAX_FILE_SIZE for larger files
4. **Memory Issues**: For large Excel files, increase Node.js memory limit

### Logging

All operations are logged with different levels:
- **INFO**: Successful operations
- **WARN**: Non-critical issues
- **ERROR**: Critical errors with stack traces

Log files are stored in the `logs/` directory.
