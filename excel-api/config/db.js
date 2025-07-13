const { logger } = require('../middleware/errorHandler');

/**
 * Database configuration and connection management
 * Supports PostgreSQL, MySQL, and SQLite
 */

let dbConnection = null;

/**
 * Initialize database connection based on DATABASE_URL
 */
const initializeDatabase = async () => {
  const databaseUrl = process.env.DATABASE_URL;
  
  if (!databaseUrl) {
    logger.warn('No DATABASE_URL provided. Database features will be disabled.');
    return null;
  }

  try {
    if (databaseUrl.startsWith('postgresql://')) {
      // PostgreSQL connection
      const { Pool } = require('pg');
      dbConnection = new Pool({
        connectionString: databaseUrl,
        ssl: process.env.NODE_ENV === 'production' ? { rejectUnauthorized: false } : false
      });
      
      // Test connection
      await dbConnection.query('SELECT NOW()');
      logger.info('PostgreSQL database connected successfully');
      
    } else if (databaseUrl.startsWith('mysql://')) {
      // MySQL connection
      const mysql = require('mysql2/promise');
      dbConnection = mysql.createPool(databaseUrl);
      
      // Test connection
      const connection = await dbConnection.getConnection();
      await connection.ping();
      connection.release();
      logger.info('MySQL database connected successfully');
      
    } else if (databaseUrl.startsWith('sqlite:')) {
      // SQLite connection
      const sqlite3 = require('sqlite3').verbose();
      const dbPath = databaseUrl.replace('sqlite:', '');
      
      dbConnection = new sqlite3.Database(dbPath, (err) => {
        if (err) {
          throw err;
        }
        logger.info('SQLite database connected successfully');
      });
      
      // Promisify SQLite methods
      dbConnection.queryAsync = function(sql, params = []) {
        return new Promise((resolve, reject) => {
          this.all(sql, params, (err, rows) => {
            if (err) reject(err);
            else resolve(rows);
          });
        });
      };
    }
    
    return dbConnection;
  } catch (error) {
    logger.error('Database connection failed:', error);
    throw new Error(`Database connection failed: ${error.message}`);
  }
};

/**
 * Execute a database query with parameters
 * @param {string} sql - SQL query string
 * @param {Array} params - Query parameters
 * @returns {Promise<Array>} - Query results
 */
const executeQuery = async (sql, params = []) => {
  if (!dbConnection) {
    throw new Error('Database not initialized. Please check your DATABASE_URL configuration.');
  }

  try {
    logger.info(`Executing query: ${sql.substring(0, 100)}...`);
    
    if (process.env.DATABASE_URL.startsWith('postgresql://')) {
      const result = await dbConnection.query(sql, params);
      return result.rows;
      
    } else if (process.env.DATABASE_URL.startsWith('mysql://')) {
      const [rows] = await dbConnection.execute(sql, params);
      return rows;
      
    } else if (process.env.DATABASE_URL.startsWith('sqlite:')) {
      return await dbConnection.queryAsync(sql, params);
    }
    
  } catch (error) {
    logger.error('Query execution failed:', error);
    throw new Error(`Database query failed: ${error.message}`);
  }
};

/**
 * Get database schema information
 * @param {string} tableName - Optional table name to get specific table info
 * @returns {Promise<Array>} - Schema information
 */
const getSchema = async (tableName = null) => {
  if (!dbConnection) {
    throw new Error('Database not initialized');
  }

  try {
    let query;
    let params = [];

    if (process.env.DATABASE_URL.startsWith('postgresql://')) {
      query = tableName 
        ? `SELECT column_name, data_type, is_nullable FROM information_schema.columns WHERE table_name = $1`
        : `SELECT table_name FROM information_schema.tables WHERE table_schema = 'public'`;
      if (tableName) params = [tableName];
      
    } else if (process.env.DATABASE_URL.startsWith('mysql://')) {
      query = tableName
        ? `SELECT COLUMN_NAME as column_name, DATA_TYPE as data_type, IS_NULLABLE as is_nullable FROM information_schema.COLUMNS WHERE TABLE_NAME = ?`
        : `SELECT TABLE_NAME as table_name FROM information_schema.TABLES WHERE TABLE_SCHEMA = DATABASE()`;
      if (tableName) params = [tableName];
      
    } else if (process.env.DATABASE_URL.startsWith('sqlite:')) {
      query = tableName
        ? `PRAGMA table_info(${tableName})`
        : `SELECT name as table_name FROM sqlite_master WHERE type='table'`;
    }

    return await executeQuery(query, params);
  } catch (error) {
    logger.error('Schema query failed:', error);
    throw new Error(`Schema query failed: ${error.message}`);
  }
};

/**
 * Validate and sanitize SQL query to prevent injection
 * @param {string} sql - SQL query to validate
 * @returns {boolean} - Whether query is safe
 */
const validateQuery = (sql) => {
  // Basic SQL injection prevention
  const dangerousPatterns = [
    /;\s*(drop|delete|truncate|alter|create|insert|update)\s+/i,
    /union\s+select/i,
    /exec\s*\(/i,
    /script\s*>/i
  ];

  return !dangerousPatterns.some(pattern => pattern.test(sql));
};

/**
 * Execute a safe SELECT query with validation
 * @param {string} sql - SELECT query
 * @param {Array} params - Query parameters
 * @returns {Promise<Array>} - Query results
 */
const executeSafeQuery = async (sql, params = []) => {
  // Ensure it's a SELECT query
  if (!sql.trim().toLowerCase().startsWith('select')) {
    throw new Error('Only SELECT queries are allowed');
  }

  // Validate query for dangerous patterns
  if (!validateQuery(sql)) {
    throw new Error('Query contains potentially dangerous patterns');
  }

  return await executeQuery(sql, params);
};

/**
 * Close database connection
 */
const closeConnection = async () => {
  if (dbConnection) {
    try {
      if (process.env.DATABASE_URL.startsWith('postgresql://')) {
        await dbConnection.end();
      } else if (process.env.DATABASE_URL.startsWith('mysql://')) {
        await dbConnection.end();
      } else if (process.env.DATABASE_URL.startsWith('sqlite:')) {
        dbConnection.close();
      }
      logger.info('Database connection closed');
    } catch (error) {
      logger.error('Error closing database connection:', error);
    }
  }
};

// Sample data for testing (when no database is configured)
const getSampleData = () => {
  return [
    { id: 1, name: 'John Doe', email: 'john@example.com', department: 'Engineering', salary: 75000, hire_date: '2023-01-15' },
    { id: 2, name: 'Jane Smith', email: 'jane@example.com', department: 'Marketing', salary: 65000, hire_date: '2023-02-20' },
    { id: 3, name: 'Bob Johnson', email: 'bob@example.com', department: 'Sales', salary: 70000, hire_date: '2023-03-10' },
    { id: 4, name: 'Alice Brown', email: 'alice@example.com', department: 'HR', salary: 60000, hire_date: '2023-04-05' },
    { id: 5, name: 'Charlie Wilson', email: 'charlie@example.com', department: 'Engineering', salary: 80000, hire_date: '2023-05-12' }
  ];
};

module.exports = {
  initializeDatabase,
  executeQuery,
  executeSafeQuery,
  getSchema,
  validateQuery,
  closeConnection,
  getSampleData,
  getConnection: () => dbConnection
};
