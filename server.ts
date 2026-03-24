import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import cors from "cors";
import dotenv from "dotenv";
import oracledb from "oracledb";
import axios from "axios";
import multer from "multer";
import { fileURLToPath } from "url";

const result = dotenv.config();
if (result.error) {
  console.log("No .env file found, relying on environment variables.");
} else {
  console.log(".env file loaded successfully.");
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
// The platform requires port 3000 for external access.
const PORT = 3000;

app.use(cors());
app.use(express.json());

// Check if critical configuration is missing
const isConfigMissing = !process.env.DEV_DB_STRING_CONNECTION || 
                        !process.env.SHAREPOINT_API_BASE_URL || 
                        !process.env.SHAREPOINT_DRIVE_ID || 
                        !process.env.SHAREPOINT_PATH_XMLS_ID;

if (isConfigMissing) {
  console.warn("WARNING: Some environment variables are missing. The system will run in MOCK MODE for missing services.");
  console.warn("Missing variables:", [
    !process.env.DEV_DB_STRING_CONNECTION && "DEV_DB_STRING_CONNECTION",
    !process.env.SHAREPOINT_API_BASE_URL && "SHAREPOINT_API_BASE_URL",
    !process.env.SHAREPOINT_DRIVE_ID && "SHAREPOINT_DRIVE_ID",
    !process.env.SHAREPOINT_PATH_XMLS_ID && "SHAREPOINT_PATH_XMLS_ID"
  ].filter(Boolean).join(", "));
}

const upload = multer({ storage: multer.memoryStorage() });

// Database connection pool
let pool: oracledb.Pool | null = null;

async function getPool() {
  if (!process.env.DEV_DB_STRING_CONNECTION) {
    return null; // Signal mock mode for DB
  }
  if (!pool) {
    try {
      const dbConfig: oracledb.PoolAttributes = {
        user: process.env.DEV_DB_USER,
        password: process.env.DEV_DB_PASSWORD,
        connectString: process.env.DEV_DB_STRING_CONNECTION,
      };

      if (process.env.DEV_ORACLE_LIB_DIR) {
        oracledb.initOracleClient({ libDir: process.env.DEV_ORACLE_LIB_DIR });
      }

      pool = await oracledb.createPool(dbConfig);
      console.log("Oracle DB pool created successfully.");
    } catch (err: any) {
      console.error("Error creating database pool:", err.message);
      return null; // Fallback to mock mode
    }
  }
  return pool;
}

// SharePoint API Helpers
const SP_BASE_URL = process.env.SHAREPOINT_API_BASE_URL?.endsWith("/") 
  ? process.env.SHAREPOINT_API_BASE_URL 
  : (process.env.SHAREPOINT_API_BASE_URL ? `${process.env.SHAREPOINT_API_BASE_URL}/` : "");

const SP_DRIVE_ID = process.env.SHAREPOINT_DRIVE_ID;
const SP_PATH_XMLS_ID = process.env.SHAREPOINT_PATH_XMLS_ID;
const SP_BEARER_TOKEN = process.env.SHAREPOINT_BEARER_TOKEN;

const spApi = axios.create({
  baseURL: SP_BASE_URL,
  headers: {
    Authorization: `Bearer ${SP_BEARER_TOKEN || ""}`,
  },
});

// SharePoint API Routes
// 1. List files
app.get("/api/sharepoint/list", async (req, res) => {
  try {
    if (!SP_BASE_URL || !SP_DRIVE_ID || !SP_PATH_XMLS_ID) {
      console.log("Using MOCK SharePoint data (config missing)");
      return res.json([
        { id: "mock-1", name: "Mock_XML_1.xml", serverRelativeUrl: "mock-1", timeCreated: new Date().toISOString() },
        { id: "mock-2", name: "Mock_XML_2.xml", serverRelativeUrl: "mock-2", timeCreated: new Date().toISOString() }
      ]);
    }
    const url = `drives/listFileFolder/${SP_DRIVE_ID}/${SP_PATH_XMLS_ID}`;
    const response = await spApi.get(url);
    
    const items = response.data.value || response.data;
    const filtered = items
      .filter((item: any) => !item.folder && !item.name.toLowerCase().includes("validado_ok"))
      .map((item: any) => ({
        id: item.id,
        name: item.name,
        serverRelativeUrl: item.id,
        timeCreated: item.createdDateTime,
      }));

    res.json(filtered);
  } catch (error: any) {
    console.error("Error listing SharePoint files:", error.message);
    // Fallback to mock data on error
    res.json([
      { id: "mock-err", name: "Error_Mock.xml", serverRelativeUrl: "mock-err", timeCreated: new Date().toISOString() }
    ]);
  }
});

// 2. Download file
app.get("/api/sharepoint/download/:itemId", async (req, res) => {
  try {
    const { itemId } = req.params;
    if (!SP_BASE_URL || !SP_DRIVE_ID) {
      console.log("Using MOCK SharePoint download");
      res.setHeader("Content-Type", "text/xml");
      res.setHeader("Content-Disposition", `attachment; filename="${itemId}.xml"`);
      return res.send("<mock>This is mock XML content</mock>");
    }
    const url = `drives/downloadFile/${SP_DRIVE_ID}/${itemId}`;
    const response = await spApi.get(url, { responseType: "arraybuffer" });
    
    res.setHeader("Content-Type", response.headers["content-type"] || "application/octet-stream");
    res.setHeader("Content-Disposition", response.headers["content-disposition"] || `attachment; filename="${itemId}.xml"`);
    res.send(Buffer.from(response.data));
  } catch (error: any) {
    console.error("Error downloading SharePoint file:", error.message);
    res.status(500).json({ error: error.message });
  }
});

// 3. Upload file
app.post("/api/sharepoint/upload", upload.single("file"), async (req: any, res) => {
  try {
    const file = req.file;
    if (!file) return res.status(400).json({ error: "No file uploaded" });

    if (!SP_BASE_URL || !SP_DRIVE_ID || !SP_PATH_XMLS_ID) {
      console.log("Using MOCK SharePoint upload");
      return res.json({ id: "mock-upload-id", name: file.originalname });
    }

    const formData = new FormData();
    const blob = new Blob([file.buffer], { type: file.mimetype });
    formData.append("file", blob, file.originalname);

    const url = `upload/large?driveId=${SP_DRIVE_ID}&parentItemId=${SP_PATH_XMLS_ID}`;
    const response = await spApi.post(url, formData);

    res.json(response.data);
  } catch (error: any) {
    console.error("Error uploading SharePoint file:", error.message);
    res.status(500).json({ error: error.message });
  }
});

// 4. Delete file
app.delete("/api/sharepoint/delete/:itemId", async (req, res) => {
  try {
    const { itemId } = req.params;
    if (!SP_BASE_URL || !SP_DRIVE_ID) {
      console.log("Using MOCK SharePoint delete");
      return res.json({ success: true });
    }
    const url = `DeleteFile?driveId=${SP_DRIVE_ID}&itemId=${itemId}`;
    const response = await spApi.delete(url);
    res.json(response.data);
  } catch (error: any) {
    console.error("Error deleting SharePoint file:", error.message);
    res.status(500).json({ error: error.message });
  }
});

// 5. Rename file
app.post("/api/sharepoint/rename", async (req, res) => {
  try {
    const { itemId, newName } = req.body;
    if (!SP_BASE_URL || !SP_DRIVE_ID) {
      console.log("Using MOCK SharePoint rename");
      return res.json({ success: true, newName });
    }
    const url = `RenameFile?driveId=${SP_DRIVE_ID}&itemId=${itemId}&newName=${encodeURIComponent(newName)}`;
    const response = await spApi.post(url);
    res.json(response.data);
  } catch (error: any) {
    console.error("Error renaming SharePoint file:", error.message);
    res.status(500).json({ error: error.message });
  }
});

// Database API Routes
// 1. History
app.get("/api/db/history", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) {
      console.log("Using MOCK DB history");
      return res.json([
        { Title: "Mock History 1", Status: "Validado", nNF: "123", CNPJ: "12345678000199", ValidationDate: new Date().toISOString() }
      ]);
    }
    connection = await pool.getConnection();
    const result = await connection.execute(
      "SELECT * FROM DHL_FullHistory ORDER BY ValidationDate DESC FETCH FIRST 5000 ROWS ONLY",
      [],
      { outFormat: oracledb.OUT_FORMAT_OBJECT }
    );
    res.json(result.rows);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

app.post("/api/db/history", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) {
      console.log("Using MOCK DB history save");
      return res.json({ success: true });
    }
    connection = await pool.getConnection();
    const { Title, Status, ServerRelativeUrl, nNF, CNPJ, OS, NCM, xProd, UserEmail, Source, ValidationDate } = req.body;
    
    await connection.execute(
      `INSERT INTO DHL_FullHistory (Title, Status, ServerRelativeUrl, nNF, CNPJ, OS, NCM, xProd, UserEmail, Source, ValidationDate) 
       VALUES (:Title, :Status, :ServerRelativeUrl, :nNF, :CNPJ, :OS, :NCM, :xProd, :UserEmail, :Source, :ValidationDate)`,
      { Title, Status, ServerRelativeUrl, nNF, CNPJ, OS, NCM, xProd, UserEmail, Source, ValidationDate: new Date(ValidationDate) },
      { autoCommit: true }
    );
    res.json({ success: true });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

// 2. Recipients
app.get("/api/db/recipients", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) {
      console.log("Using MOCK DB recipients");
      return res.json([{ Title: "Mock Recipient 1" }, { Title: "Mock Recipient 2" }]);
    }
    connection = await pool.getConnection();
    const result = await connection.execute("SELECT * FROM DHL_Recipients", [], { outFormat: oracledb.OUT_FORMAT_OBJECT });
    res.json(result.rows);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

app.post("/api/db/recipients", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) return res.json({ success: true });
    connection = await pool.getConnection();
    const { Title } = req.body;
    await connection.execute("INSERT INTO DHL_Recipients (Title) VALUES (:Title)", { Title }, { autoCommit: true });
    res.json({ success: true });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

app.delete("/api/db/recipients/:title", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) return res.json({ success: true });
    connection = await pool.getConnection();
    const { title } = req.params;
    await connection.execute("DELETE FROM DHL_Recipients WHERE Title = :Title", { Title: title }, { autoCommit: true });
    res.json({ success: true });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

app.put("/api/db/recipients/:title", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) return res.json({ success: true });
    connection = await pool.getConnection();
    const oldTitle = req.params.title;
    const { Title } = req.body;
    await connection.execute("UPDATE DHL_Recipients SET Title = :Title WHERE Title = :oldTitle", { Title, oldTitle }, { autoCommit: true });
    res.json({ success: true });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

// 3. Mandatory Tags
app.get("/api/db/tags", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) {
      console.log("Using MOCK DB tags");
      return res.json([{ Title: "Mock Tag 1", TagRef: "MOCK_1" }]);
    }
    connection = await pool.getConnection();
    const result = await connection.execute("SELECT * FROM DHL_MandatoryTags", [], { outFormat: oracledb.OUT_FORMAT_OBJECT });
    res.json(result.rows);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

app.post("/api/db/tags", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) return res.json({ success: true });
    connection = await pool.getConnection();
    const { Title, TagRef } = req.body;
    await connection.execute("INSERT INTO DHL_MandatoryTags (Title, TagRef) VALUES (:Title, :TagRef)", { Title, TagRef }, { autoCommit: true });
    res.json({ success: true });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

app.delete("/api/db/tags/:tagRef", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) return res.json({ success: true });
    connection = await pool.getConnection();
    const { tagRef } = req.params;
    await connection.execute("DELETE FROM DHL_MandatoryTags WHERE TagRef = :TagRef", { TagRef: tagRef }, { autoCommit: true });
    res.json({ success: true });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

app.put("/api/db/tags/:tagRef", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) return res.json({ success: true });
    connection = await pool.getConnection();
    const { tagRef } = req.params;
    const { Title } = req.body;
    await connection.execute("UPDATE DHL_MandatoryTags SET Title = :Title WHERE TagRef = :TagRef", { Title, TagRef: tagRef }, { autoCommit: true });
    res.json({ success: true });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

// 4. OS Patterns
app.get("/api/db/os-patterns", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) {
      console.log("Using MOCK DB OS patterns");
      return res.json([{ Title: "Mock Pattern 1" }]);
    }
    connection = await pool.getConnection();
    const result = await connection.execute("SELECT * FROM DHL_OSPatterns", [], { outFormat: oracledb.OUT_FORMAT_OBJECT });
    res.json(result.rows);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

app.post("/api/db/os-patterns", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) return res.json({ success: true });
    connection = await pool.getConnection();
    const { Title } = req.body;
    await connection.execute("INSERT INTO DHL_OSPatterns (Title) VALUES (:Title)", { Title }, { autoCommit: true });
    res.json({ success: true });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

app.delete("/api/db/os-patterns/:title", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) return res.json({ success: true });
    connection = await pool.getConnection();
    const { title } = req.params;
    await connection.execute("DELETE FROM DHL_OSPatterns WHERE Title = :Title", { Title: title }, { autoCommit: true });
    res.json({ success: true });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

// 5. Config
app.get("/api/db/config", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) {
      console.log("Using MOCK DB config");
      return res.json([{ Title: "Mock Config", Value: "Mock Value" }]);
    }
    connection = await pool.getConnection();
    const result = await connection.execute("SELECT * FROM DHL_Config", [], { outFormat: oracledb.OUT_FORMAT_OBJECT });
    res.json(result.rows);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

app.post("/api/db/config", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) return res.json({ success: true });
    connection = await pool.getConnection();
    const { Title, Value } = req.body;
    await connection.execute(
      `MERGE INTO DHL_Config c
       USING (SELECT :Title as Title, :Value as Value FROM dual) src
       ON (c.Title = src.Title)
       WHEN MATCHED THEN UPDATE SET c.Value = src.Value
       WHEN NOT MATCHED THEN INSERT (Title, Value) VALUES (src.Title, src.Value)`,
      { Title, Value },
      { autoCommit: true }
    );
    res.json({ success: true });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

// 6. Validation History
app.get("/api/db/validation-history", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    connection = await pool.getConnection();
    const result = await connection.execute(
      "SELECT * FROM DHL_ValidationHistory ORDER BY ValidationDate DESC",
      [],
      { outFormat: oracledb.OUT_FORMAT_OBJECT }
    );
    res.json(result.rows);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

app.post("/api/db/validation-history", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) return res.json({ success: true });
    connection = await pool.getConnection();
    const { Title, ServerRelativeUrl, nNF, CNPJ, OS, NCM, xProd, Status, ValidationDate } = req.body;
    await connection.execute(
      `INSERT INTO DHL_ValidationHistory (Title, ServerRelativeUrl, nNF, CNPJ, OS, NCM, xProd, Status, ValidationDate) 
       VALUES (:Title, :ServerRelativeUrl, :nNF, :CNPJ, :OS, :NCM, :xProd, :Status, :ValidationDate)`,
      { Title, ServerRelativeUrl, nNF, CNPJ, OS, NCM, xProd, Status, ValidationDate: new Date(ValidationDate) },
      { autoCommit: true }
    );
    res.json({ success: true });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

app.delete("/api/db/validation-history/:id", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) return res.json({ success: true });
    connection = await pool.getConnection();
    const { id } = req.params;
    await connection.execute("DELETE FROM DHL_ValidationHistory WHERE ID = :ID", { ID: id }, { autoCommit: true });
    res.json({ success: true });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

// 7. Registered Products
app.get("/api/db/registered-products", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) {
      console.log("Using MOCK DB registered products");
      return res.json(["Mock Product 1", "Mock Product 2"]);
    }
    connection = await pool.getConnection();
    const result = await connection.execute(
      "SELECT PRODUCT_NAME FROM DHL_RegisteredProducts",
      [],
      { outFormat: oracledb.OUT_FORMAT_OBJECT }
    );
    res.json(result.rows.map((r: any) => (r as any).PRODUCT_NAME));
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

app.post("/api/db/registered-products", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) return res.json({ success: true });
    connection = await pool.getConnection();
    const { productName } = req.body;
    await connection.execute(
      "INSERT INTO DHL_RegisteredProducts (PRODUCT_NAME) VALUES (:productName)",
      { productName },
      { autoCommit: true }
    );
    res.json({ success: true });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

app.delete("/api/db/registered-products/:productName", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) return res.json({ success: true });
    connection = await pool.getConnection();
    const { productName } = req.params;
    await connection.execute(
      "DELETE FROM DHL_RegisteredProducts WHERE PRODUCT_NAME = :productName",
      { productName },
      { autoCommit: true }
    );
    res.json({ success: true });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

// 8. External DB Queries (NTV, OS, NCM)
app.post("/api/db/query/ntv", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) {
      console.log("Using MOCK DB query NTV");
      return res.json([{ PRTNUM: "MOCK-PRT-1", DESCR: "Mock Description" }]);
    }
    connection = await pool.getConnection();
    const { product } = req.body;
    const result = await connection.execute(
      `SELECT * FROM PRTMST WHERE UPPER(PRTNUM) LIKE UPPER(:product)`,
      { product: `%${product}%` },
      { outFormat: oracledb.OUT_FORMAT_OBJECT }
    );
    res.json(result.rows);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

app.post("/api/db/query/os", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) {
      console.log("Using MOCK DB query OS");
      return res.json([{ WAYBIL: "MOCK-OS-1", STATUS: "Mock Status" }]);
    }
    connection = await pool.getConnection();
    const { osNumber } = req.body;
    const result = await connection.execute(
      `SELECT * FROM RIMHDR WHERE WAYBIL = :osNumber`,
      { osNumber },
      { outFormat: oracledb.OUT_FORMAT_OBJECT }
    );
    res.json(result.rows);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

app.post("/api/db/query/ncm", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) {
      console.log("Using MOCK DB query NCM");
      return res.json([{ TYPCOD: "MOCK-NCM-1", DESCR: "Mock Description" }]);
    }
    connection = await pool.getConnection();
    const { ncm } = req.body;
    const result = await connection.execute(
      `SELECT * FROM PRTMST WHERE UPPER(TYPCOD) LIKE UPPER(:ncm)`,
      { ncm: `%${ncm}%` },
      { outFormat: oracledb.OUT_FORMAT_OBJECT }
    );
    res.json(result.rows);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

// 9. Initialize Database
app.post("/api/db/initialize", async (req, res) => {
  let connection;
  try {
    const pool = await getPool();
    if (!pool) {
      console.log("Using MOCK DB initialize");
      return res.json({ success: true });
    }
    connection = await pool.getConnection();
    
    const tables = [
      {
        name: 'DHL_Recipients',
        sql: `CREATE TABLE DHL_Recipients (Title VARCHAR2(255) PRIMARY KEY)`
      },
      {
        name: 'DHL_MandatoryTags',
        sql: `CREATE TABLE DHL_MandatoryTags (Title VARCHAR2(255), TagRef VARCHAR2(255) PRIMARY KEY)`
      },
      {
        name: 'DHL_OSForbiddenPatterns',
        sql: `CREATE TABLE DHL_OSForbiddenPatterns (Title VARCHAR2(255) PRIMARY KEY)`
      },
      {
        name: 'DHL_Config',
        sql: `CREATE TABLE DHL_Config (Title VARCHAR2(255) PRIMARY KEY, Value VARCHAR2(255))`
      },
      {
        name: 'DHL_FullHistory',
        sql: `CREATE TABLE DHL_FullHistory (
          Title VARCHAR2(255), 
          Status VARCHAR2(255), 
          ServerRelativeUrl VARCHAR2(1000), 
          nNF VARCHAR2(255), 
          CNPJ VARCHAR2(255), 
          OS VARCHAR2(255), 
          NCM VARCHAR2(255), 
          xProd VARCHAR2(1000), 
          UserEmail VARCHAR2(255), 
          Source VARCHAR2(255), 
          ValidationDate TIMESTAMP
        )`
      },
      {
        name: 'DHL_ValidationHistory',
        sql: `CREATE TABLE DHL_ValidationHistory (
          ID NUMBER GENERATED ALWAYS AS IDENTITY PRIMARY KEY, 
          Title VARCHAR2(255), 
          ServerRelativeUrl VARCHAR2(1000), 
          nNF VARCHAR2(255), 
          CNPJ VARCHAR2(255), 
          OS VARCHAR2(255), 
          NCM VARCHAR2(255), 
          xProd VARCHAR2(1000), 
          Status VARCHAR2(255), 
          ValidationDate TIMESTAMP
        )`
      },
      {
        name: 'DHL_RegisteredProducts',
        sql: `CREATE TABLE DHL_RegisteredProducts (PRODUCT_NAME VARCHAR2(255) PRIMARY KEY)`
      }
    ];

    for (const table of tables) {
      try {
        await connection.execute(table.sql);
        console.log(`Table ${table.name} created.`);
      } catch (err: any) {
        if (err.errorNum === 955) {
          console.log(`Table ${table.name} already exists.`);
        } else {
          console.error(`Error creating table ${table.name}:`, err.message);
        }
      }
    }

    res.json({ success: true });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (connection) await connection.close();
  }
});

// Vite middleware setup
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on port ${PORT}`);
    if (process.env.DEV_PORT) {
      console.log(`Note: DEV_PORT is set to ${process.env.DEV_PORT}, but the server is running on ${PORT} as required by the platform.`);
    }
  });
}

startServer();
