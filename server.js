const express = require("express");
const cors = require("cors");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = 3001;

// ─── Copy library dari node_modules ke public/libs saat server start ──────────
function copyLibs() {
  const libsDir = path.join(__dirname, "public", "libs");

  // Buat folder public/libs jika belum ada
  if (!fs.existsSync(libsDir)) {
    fs.mkdirSync(libsDir, { recursive: true });
    console.log("[libs] Folder public/libs dibuat");
  }

  const files = [
    {
      src: path.join(
        __dirname,
        "node_modules",
        "xlsx",
        "dist",
        "xlsx.full.min.js",
      ),
      dest: path.join(libsDir, "xlsx.full.min.js"),
      name: "SheetJS (xlsx)",
    },
    {
      src: path.join(
        __dirname,
        "node_modules",
        "jspdf",
        "dist",
        "jspdf.umd.min.js",
      ),
      dest: path.join(libsDir, "jspdf.umd.min.js"),
      name: "jsPDF",
    },
    {
      src: path.join(
        __dirname,
        "node_modules",
        "jspdf-autotable",
        "dist",
        "jspdf.plugin.autotable.min.js",
      ),
      dest: path.join(libsDir, "jspdf.plugin.autotable.min.js"),
      name: "jsPDF-AutoTable",
    },
  ];

  files.forEach(function (f) {
    if (!fs.existsSync(f.src)) {
      console.warn("[libs] TIDAK DITEMUKAN:", f.src);
      console.warn("[libs] Jalankan: npm install");
      return;
    }
    fs.copyFileSync(f.src, f.dest);
    console.log("[libs] OK:", f.name, "→", f.dest);
  });
}

copyLibs();

// ─── Middleware ────────────────────────────────────────────────────────────────
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

// ─── Proxy endpoint ke backend SAP ────────────────────────────────────────────
app.get("/api/sales-orders", async (req, res) => {
  let { iv_werks, iv_auart } = req.query;
  if (!iv_werks) iv_werks = "3000";
  if (!iv_auart) iv_auart = "ZOR2";

  const targetUrl = `https://backend-sap-ui-5.kayumebelsmg.net/api/local-so-data?iv_werks=${iv_werks}&iv_auart=${iv_auart}`;

  try {
    console.log(`[Proxy] Requesting: ${targetUrl}`);
    const response = await fetch(targetUrl, {
      headers: {
        "X-API-KEY": "Kmi3Seamrang123",
        accept: "application/json",
      },
    });

    if (!response.ok) throw new Error(`HTTP ${response.status}`);

    const result = await response.json();
    console.log("[Proxy] Response success:", result.success);

    let salesData = [];
    if (
      result.data &&
      result.data.t_data1 &&
      Array.isArray(result.data.t_data1)
    ) {
      salesData = result.data.t_data1;
    } else if (Array.isArray(result.data)) {
      salesData = result.data;
    } else if (Array.isArray(result)) {
      salesData = result;
    }

    res.json(salesData);
  } catch (error) {
    console.error("[Proxy] Error:", error.message);
    res.status(200).json([]);
  }
});

app.listen(PORT, () => {
  console.log(`\nServer running at http://localhost:${PORT}`);
  console.log(`Cek library: http://localhost:${PORT}/libs/xlsx.full.min.js`);
});
