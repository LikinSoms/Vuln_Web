const express = require("express");
const sqlite3 = require("sqlite3").verbose();
const crypto = require("crypto");
const path = require("path");

const app = express();
const db = new sqlite3.Database(process.env.DB_PATH || path.join(__dirname, "store.sqlite"));
const sessions = new Map();
const PORT = process.env.PORT || 3000;
const MAX_CART_QUANTITY = 25;

app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(express.static(path.join(__dirname, "public")));

function hashPassword(password) {
  return crypto.createHash("sha256").update(String(password)).digest("hex");
}

function positiveInteger(value) {
  const text = String(value ?? "").trim();
  if (!/^\d+$/.test(text)) return null;

  const parsed = Number(text);
  return Number.isSafeInteger(parsed) && parsed > 0 ? parsed : null;
}

function run(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.run(sql, params, function (err) {
      if (err) reject(err);
      else resolve(this);
    });
  });
}

function get(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.get(sql, params, (err, row) => (err ? reject(err) : resolve(row)));
  });
}

function all(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.all(sql, params, (err, rows) => (err ? reject(err) : resolve(rows)));
  });
}

function cookies(req) {
  return Object.fromEntries(
    (req.headers.cookie || "")
      .split(";")
      .filter(Boolean)
      .map((part) => {
        const pieces = part.trim().split("=");
        return [decodeURIComponent(pieces.shift()), decodeURIComponent(pieces.join("="))];
      })
  );
}

function currentUser(req) {
  return sessions.get(cookies(req).sid);
}

function requireUser(req, res, next) {
  const user = currentUser(req);
  if (!user) return res.status(401).json({ error: "Login required" });
  req.user = user;
  next();
}

function signInUser(res, user) {
  const token = crypto.randomBytes(24).toString("hex");
  sessions.set(token, {
    id: user.id,
    username: user.username,
    displayName: user.display_name || user.displayName
  });

  res.cookie("sid", token, {
    httpOnly: true,
    sameSite: "lax",
    maxAge: 1000 * 60 * 60 * 2
  });

  return sessions.get(token);
}

async function insertSeedOrder(userId, lines) {
  const total = Number(lines.reduce((sum, item) => sum + item.price * item.quantity, 0).toFixed(2));
  const order = await run(
    "INSERT INTO orders (user_id, created_at, total) VALUES (?, ?, ?)",
    [userId, new Date().toISOString(), total]
  );

  for (const item of lines) {
    await run(
      "INSERT INTO order_items (order_id, item_id, name, price, quantity, subtotal) VALUES (?, ?, ?, ?, ?, ?)",
      [order.lastID, item.id, item.name, item.price, item.quantity, Number((item.price * item.quantity).toFixed(2))]
    );
  }
}

async function initDb() {
  await run(`CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    display_name TEXT NOT NULL
  )`);

  await run(`CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    description TEXT NOT NULL,
    price REAL NOT NULL,
    image_url TEXT
  )`);

  const productColumns = await all("PRAGMA table_info(products)");
  if (!productColumns.some((column) => column.name === "image_url")) {
    await run("ALTER TABLE products ADD COLUMN image_url TEXT");
  }

  await run(`CREATE TABLE IF NOT EXISTS orders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    created_at TEXT NOT NULL,
    total REAL NOT NULL
  )`);

  await run(`CREATE TABLE IF NOT EXISTS order_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id INTEGER NOT NULL,
    item_id INTEGER NOT NULL,
    name TEXT NOT NULL,
    price REAL NOT NULL,
    quantity INTEGER NOT NULL,
    subtotal REAL NOT NULL
  )`);

  const userCount = await get("SELECT COUNT(*) AS count FROM users");
  if (!userCount.count) {
    const users = [
      ["alice", "password123", "Alice Carter"],
      ["bob", "password123", "Bob Singh"],
      ["admin", "admin123", "Store Admin"]
    ];

    for (const [username, password, displayName] of users) {
      await run(
        "INSERT INTO users (username, password_hash, display_name) VALUES (?, ?, ?)",
        [username, hashPassword(password), displayName]
      );
    }
  }

  const productCatalog = [
    ["Canvas Tote Bag", "Durable everyday tote with reinforced handles.", 18.99, "/images/canvas-tote-bag.png"],
    ["Desk Lamp", "Adjustable LED lamp for focused work.", 34.5, "/images/desk-lamp.png"],
    ["Wireless Mouse", "Compact mouse with quiet clicks.", 24.99, "/images/wireless-mouse.png"],
    ["Travel Mug", "Insulated stainless steel mug.", 16.75, "/images/travel-mug.png"],
    ["Notebook Set", "Three soft-cover notebooks with dotted pages.", 12.25, "/images/notebook-set.png"],
    ["Bluetooth Speaker", "Portable speaker with clear room-filling sound.", 49.99, "/images/bluetooth-speaker.png"]
  ];

  const productCount = await get("SELECT COUNT(*) AS count FROM products");
  if (!productCount.count) {
    for (const product of productCatalog) {
      await run(
        "INSERT INTO products (name, description, price, image_url) VALUES (?, ?, ?, ?)",
        product
      );
    }
  }

  for (const product of productCatalog) {
    await run(
      "UPDATE products SET image_url = ? WHERE name = ?",
      [product[3], product[0]]
    );
  }

  const orderCount = await get("SELECT COUNT(*) AS count FROM orders");
  if (!orderCount.count) {
    const alice = await get("SELECT id FROM users WHERE username = ?", ["alice"]);
    const bob = await get("SELECT id FROM users WHERE username = ?", ["bob"]);
    const products = await all("SELECT * FROM products ORDER BY id LIMIT 4");

    await insertSeedOrder(alice.id, [
      { ...products[0], quantity: 1 },
      { ...products[2], quantity: 2 }
    ]);

    await insertSeedOrder(bob.id, [
      { ...products[1], quantity: 1 },
      { ...products[3], quantity: 1 }
    ]);
  }
}

app.get("/api/products", async (req, res) => {
  res.json(await all("SELECT id, name, description, price, image_url AS imageUrl FROM products ORDER BY id"));
});

app.get("/api/me", (req, res) => {
  const user = currentUser(req);
  res.json({ user: user || null });
});

app.post("/api/login", (req, res) => {
  const username = req.body.username || "";
  const passwordHash = hashPassword(req.body.password || "");
  const sql =
    "SELECT id, username, display_name FROM users WHERE username = '" +
    username +
    "' AND password_hash = '" +
    passwordHash +
    "'";

  db.get(sql, (err, user) => {
    if (err) return res.status(500).json({ error: "Login failed" });
    if (!user) return res.status(401).json({ error: "Invalid credentials" });

    res.json({ user: signInUser(res, user) });
  });
});

app.post("/api/register", async (req, res) => {
  const username = String(req.body.username || "").trim();
  const password = String(req.body.password || "");
  const displayName = String(req.body.displayName || "").trim();

  if (!username || !password || !displayName) {
    return res.status(400).json({ error: "All fields are required" });
  }

  if (username.length < 3) {
    return res.status(400).json({ error: "Username must be at least 3 characters" });
  }

  if (password.length < 6) {
    return res.status(400).json({ error: "Password must be at least 6 characters" });
  }

  try {
    const result = await run(
      "INSERT INTO users (username, password_hash, display_name) VALUES (?, ?, ?)",
      [username, hashPassword(password), displayName]
    );

    res.status(201).json({
      user: signInUser(res, {
        id: result.lastID,
        username,
        displayName
      })
    });
  } catch (err) {
    if (err && err.code === "SQLITE_CONSTRAINT") {
      return res.status(409).json({ error: "Username is already taken" });
    }

    res.status(500).json({ error: "Registration failed" });
  }
});

app.post("/api/logout", (req, res) => {
  const sid = cookies(req).sid;
  if (sid) sessions.delete(sid);
  res.clearCookie("sid");
  res.json({ ok: true });
});

app.get("/api/orders", requireUser, async (req, res) => {
  const orders = await all(
    "SELECT id, created_at, total FROM orders WHERE user_id = ? ORDER BY id DESC",
    [req.user.id]
  );
  res.json({ orders });
});

app.post("/api/checkout", requireUser, async (req, res) => {
  const items = Array.isArray(req.body.items) ? req.body.items : [];
  if (!items.length) return res.status(400).json({ error: "Cart is empty" });

  try {
    const lines = [];
    let total = 0;

    for (const item of items) {
      const productId = positiveInteger(item.id);
      const quantity = positiveInteger(item.quantity);

      if (!productId || !quantity || quantity > MAX_CART_QUANTITY) {
        return res.status(400).json({ error: "Invalid cart item" });
      }

      const product = await get("SELECT id, name FROM products WHERE id = ?", [productId]);

      if (!product) return res.status(400).json({ error: "Unknown item" });

      const price = Number(item.price);
      const subtotal = Number((price * quantity).toFixed(2));

      lines.push({ product, price, quantity, subtotal });
      total += subtotal;
    }

    total = Number(total.toFixed(2));

    await run("BEGIN TRANSACTION");

    const order = await run(
      "INSERT INTO orders (user_id, created_at, total) VALUES (?, ?, ?)",
      [req.user.id, new Date().toISOString(), total]
    );

    for (const line of lines) {
      await run(
        "INSERT INTO order_items (order_id, item_id, name, price, quantity, subtotal) VALUES (?, ?, ?, ?, ?, ?)",
        [order.lastID, line.product.id, line.product.name, line.price, line.quantity, line.subtotal]
      );
    }

    await run("COMMIT");
    res.json({ orderId: order.lastID, total });
  } catch (err) {
    await run("ROLLBACK").catch(() => {});
    res.status(500).json({ error: "Checkout failed" });
  }
});

app.get("/api/orders/:id/invoice", requireUser, async (req, res) => {
  const orderId = positiveInteger(req.params.id);
  if (!orderId) return res.status(404).json({ error: "Invoice not found" });

  const order = await get(
    `SELECT orders.id, orders.created_at, orders.total, users.display_name, users.username
     FROM orders
     JOIN users ON users.id = orders.user_id
     WHERE orders.id = ?`,
    [orderId]
  );

  if (!order) return res.status(404).json({ error: "Invoice not found" });

  const items = await all(
    "SELECT item_id, name, price, quantity, subtotal FROM order_items WHERE order_id = ?",
    [order.id]
  );

  res.json({
    id: order.id,
    createdAt: order.created_at,
    total: order.total,
    customer: {
      name: order.display_name,
      username: order.username
    },
    items
  });
});

initDb()
  .then(() => {
    app.listen(PORT, () => {
      console.log(`Store running at http://localhost:${PORT}`);
    });
  })
  .catch((err) => {
    console.error(err);
    process.exit(1);
  });
