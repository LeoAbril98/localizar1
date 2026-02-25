import express from "express";
import { createServer as createViteServer } from "vite";
import Database from "better-sqlite3";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const db = new Database("inventory.db");

// Initialize database
db.exec(`
  CREATE TABLE IF NOT EXISTS inventory (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT,
    model TEXT,
    location TEXT,
    quantity TEXT
  );
  CREATE TABLE IF NOT EXISTS metadata (
    key TEXT PRIMARY KEY,
    value TEXT
  );
`);

// Insert default data if empty
const count = db.prepare("SELECT COUNT(*) as count FROM inventory").get() as { count: number };
if (count.count === 0) {
  const insertItem = db.prepare("INSERT INTO inventory (code, model, location, quantity) VALUES (?, ?, ?, ?)");
  const defaultItems = [
    { Código: "123456", Modelo: "Produto Exemplo A", Local: "Prateleira A1", Quantidade: 10 },
    { Código: "789012", Modelo: "Produto Exemplo B", Local: "Corredor B2", Quantidade: 5 }
  ];
  for (const item of defaultItems) {
    insertItem.run(item.Código, item.Modelo, item.Local, item.Quantidade);
  }
  db.prepare("INSERT OR REPLACE INTO metadata (key, value) VALUES (?, ?)").run('filename', 'estoque_padrao.xlsx');
}

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json({ limit: '50mb' }));

  // API Routes
  app.get("/api/inventory", (req, res) => {
    console.log("GET /api/inventory request received");
    try {
      const items = db.prepare("SELECT code as Código, model as Modelo, location as Local, quantity as Quantidade FROM inventory").all();
      const fileName = db.prepare("SELECT value FROM metadata WHERE key = 'filename'").get() as { value: string } | undefined;
      console.log(`Found ${items.length} items in database`);
      res.json({ items, fileName: fileName?.value || null });
    } catch (err) {
      console.error("Database error:", err);
      res.status(500).json({ error: "Failed to fetch inventory" });
    }
  });

  app.post("/api/inventory", (req, res) => {
    console.log("POST /api/inventory request received");
    const { items, fileName } = req.body;
    
    if (!Array.isArray(items)) {
      return res.status(400).json({ error: "Invalid data format" });
    }

    try {
      const deleteInventory = db.prepare("DELETE FROM inventory");
      const insertItem = db.prepare("INSERT INTO inventory (code, model, location, quantity) VALUES (?, ?, ?, ?)");
      const upsertMetadata = db.prepare("INSERT OR REPLACE INTO metadata (key, value) VALUES (?, ?)");

      const transaction = db.transaction((items: any[], fileName: string) => {
        deleteInventory.run();
        for (const item of items) {
          insertItem.run(String(item.Código), String(item.Modelo), String(item.Local), String(item.Quantidade));
        }
        if (fileName) {
          upsertMetadata.run('filename', fileName);
        }
      });

      transaction(items, fileName);
      res.json({ success: true });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: "Failed to save inventory" });
    }
  });

  app.delete("/api/inventory", (req, res) => {
    try {
      db.prepare("DELETE FROM inventory").run();
      db.prepare("DELETE FROM metadata WHERE key = 'filename'").run();
      res.json({ success: true });
    } catch (err) {
      res.status(500).json({ error: "Failed to clear inventory" });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.join(__dirname, "dist")));
    app.get("*", (req, res) => {
      res.sendFile(path.join(__dirname, "dist", "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
