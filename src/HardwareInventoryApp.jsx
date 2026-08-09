import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  LayoutDashboard, Package, Receipt, FileClock, Plus, Minus, Trash2,
  Search, X, Printer, AlertTriangle, Save, ChevronRight, ChevronLeft, Pencil,
  ShoppingCart, Settings as SettingsIcon, ArrowUpCircle, ArrowDownCircle,
  Check, Loader2, Wallet, CalendarDays, LogOut, Upload
} from "lucide-react";
import * as XLSX from "xlsx";

/* ---------------------------------------------------------
   PRODUCT CATEGORIZATION HELPER
   --------------------------------------------------------- */
const categorizeProductName = (name) => {
  const n = name.toLowerCase();
  const elecKws = ["wire", "cable", "switch", "socket", "tester", "holder", "smd", "bulb", "breaker", 
                   "db ", "meter box", "tape osaka", "insulation tape", "thimmal", "piano", "capacitor", "saddle 25mm", "saddle 32mm"];
  const plumbKws = ["pipe", "nipple", "elbow", "tee", "socket", "union", "valve", "cock", "shower", 
                    "bason", "basin", "flush", "waste", "jali", "tanki", "ppr", "pvc", "gi ", "bush", "connection pipe", "mixer pipe", "nrv", "saddle"];
  const toolKws = ["disc", "cutting", "grinding", "steel", "welding", "nail", "keel", "hinges", "qabza", 
                   "screw", "bolt", "ficher", "lock", "wahoo", "wahu", "key", "hammer", "wrench", "pliers", 
                   "cutter", "gainti", "bailcha", "chisel", "saw", "tape measuring", "measuring tape", "tape masking", "masking tape", 
                   "gloves", "glove", "brush", "dhaga", "sootar", "level", "karandi", "fara", "bracket", "clamp", "unifix", "doori"];
  const paintKws = ["paint", "oil paint", "enamel", "roller", "varnish", "colour", "putty", "solution", 
                    "samad", "elfi", "glue"];
  const cementKws = ["sand paper", "cement", "bond"];

  if (elecKws.some(kw => n.includes(kw))) return "Electrical";
  if (plumbKws.some(kw => n.includes(kw))) return "Plumbing & Sanitary";
  if (paintKws.some(kw => n.includes(kw))) return "Paint";
  if (cementKws.some(kw => n.includes(kw))) return "Cement & Aggregates";
  if (toolKws.some(kw => n.includes(kw))) return "Hardware & Tools";
  return "General";
};

const roundNum = (num, decimals = 2) => {
  const p = Math.pow(10, decimals);
  return Math.round((Number(num) || 0) * p) / p;
};

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
const fmtNum = (n) => (Math.round((Number(n) || 0) * 100) / 100).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtDate = (iso) => new Date(iso).toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
const fmtDateTime = (iso) => new Date(iso).toLocaleString(undefined, { year: "numeric", month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" });
const pad2 = (n) => String(n).padStart(2, "0");
const toDateKey = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const todayKey = () => toDateKey(new Date());
const fmtDayLabel = (key) => {
  const d = new Date(key + "T00:00:00");
  const today = todayKey();
  const yestKey = toDateKey(new Date(Date.now() - 86400000));
  if (key === today) return "Today";
  if (key === yestKey) return "Yesterday";
  return d.toLocaleDateString(undefined, { weekday: "short", year: "numeric", month: "short", day: "numeric" });
};

const DEFAULT_SETTINGS = { 
  shopName: "My Organization", 
  shopNameUrdu: "",
  phone: "",
  email: "",
  whatsapp: "",
  paymentDetails: "",
  bankDetails: "",
  address: "",
  logoUrl: "",
  quotationTitle: "MATERIAL REQUEST",
  invoiceTitle: "TAX INVOICE",
  currencySymbol: "Rs ", 
  invoiceCounter: 1, 
  lowStockDefault: 5 
};

const sellPrice = (p) => (Number(p.costPrice) || 0) * (1 + (Number(p.markup) || 0) / 100);

const SEED_PRODUCTS = [
  { id: "p0001", name: ".", category: "General", unit: "piece", quantity: 0, costPrice: 380.0, markup: 40.0, lowStock: 5 },
  { id: "p0002", name: "0-19 DB", category: "General", unit: "piece", quantity: 0, costPrice: 2600.0, markup: 40.0, lowStock: 5 },
  { id: "p0003", name: "0-4 DB", category: "General", unit: "piece", quantity: 0, costPrice: 1400.0, markup: 40.0, lowStock: 5 },
  { id: "p0004", name: "1\" CONDUIT PIPE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 88.09, markup: 40.0, lowStock: 5 },
  { id: "p0005", name: "2 CORE CABLE 40/76X90MTR", category: "Electrical", unit: "piece", quantity: 0, costPrice: 4950.0, markup: 40.0, lowStock: 5 },
  { id: "p0006", name: "2 CORE CABLE 40/76X90MTR  ORANGE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 4950.0, markup: 40.0, lowStock: 5 },
  { id: "p0007", name: "2 CORE CABLE 7/29X90MTR FINE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 11500.0, markup: 40.0, lowStock: 5 },
  { id: "p0008", name: "2 PIN  PLUG STONE", category: "General", unit: "piece", quantity: 0, costPrice: 24.17, markup: 40.0, lowStock: 5 },
  { id: "p0009", name: "2 PIN PLUG ROX", category: "General", unit: "piece", quantity: 0, costPrice: 83.33, markup: 30.0, lowStock: 5 },
  { id: "p0010", name: "2 PIN SHOE HMA", category: "General", unit: "piece", quantity: 0, costPrice: 20.0, markup: 30.0, lowStock: 5 },
  { id: "p0011", name: "2 PIN SHOE PATHAR", category: "General", unit: "piece", quantity: 0, costPrice: 30.0, markup: 40.0, lowStock: 5 },
  { id: "p0012", name: "2 PIN SHOE PATHAR GOOD", category: "General", unit: "piece", quantity: 0, costPrice: 12.0, markup: 40.0, lowStock: 5 },
  { id: "p0013", name: "2 PIN SWITCH PATHAR", category: "Electrical", unit: "piece", quantity: 0, costPrice: 30.0, markup: 30.0, lowStock: 5 },
  { id: "p0014", name: "23/76 CABLE BAREEK WHITE", category: "Electrical", unit: "roll", quantity: 0, costPrice: 800.0, markup: 40.0, lowStock: 5 },
  { id: "p0015", name: "2PIIN SHOE", category: "General", unit: "piece", quantity: 0, costPrice: 43.33, markup: 35.0, lowStock: 5 },
  { id: "p0016", name: "2PIN HOLDER", category: "Electrical", unit: "piece", quantity: 0, costPrice: 30.76, markup: 40.0, lowStock: 5 },
  { id: "p0017", name: "2PIN LIGHT MULTI", category: "General", unit: "piece", quantity: 0, costPrice: 110.0, markup: 35.0, lowStock: 5 },
  { id: "p0018", name: "2PIN SHOE PLASTIC", category: "General", unit: "piece", quantity: 0, costPrice: 13.75, markup: 40.0, lowStock: 5 },
  { id: "p0019", name: "2PIN SHOE STONE", category: "General", unit: "piece", quantity: 0, costPrice: 20.0, markup: 40.0, lowStock: 5 },
  { id: "p0020", name: "2X12 TOPLESS WOOD NAIL", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 420.0, markup: 30.0, lowStock: 5 },
  { id: "p0021", name: "3 PIN BREAKER SHOE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 550.0, markup: 30.0, lowStock: 5 },
  { id: "p0022", name: "3 PIN PLUG BIG", category: "General", unit: "piece", quantity: 0, costPrice: 82.08, markup: 40.0, lowStock: 5 },
  { id: "p0023", name: "3 PIN SHOE BIG POWER PLUG", category: "General", unit: "piece", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p0024", name: "3 PIN SHOE LIGHT PLUG", category: "General", unit: "piece", quantity: 0, costPrice: 51.5, markup: 40.0, lowStock: 5 },
  { id: "p0025", name: "3 PIN SHOE POWE PLUG", category: "General", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p0026", name: "3 PIN SHOE SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 53.0, markup: 40.0, lowStock: 5 },
  { id: "p0027", name: "3 PIN SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 45.83, markup: 30.0, lowStock: 5 },
  { id: "p0028", name: "3/29 CABLE PAK BILAL", category: "Electrical", unit: "piece", quantity: 0, costPrice: 3700.0, markup: 40.0, lowStock: 5 },
  { id: "p0029", name: "3/4\" CONDUIT PIPE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 66.93, markup: 40.0, lowStock: 5 },
  { id: "p0030", name: "3PIN LIGHT MULTI", category: "General", unit: "piece", quantity: 0, costPrice: 120.0, markup: 35.0, lowStock: 5 },
  { id: "p0031", name: "3PIN SHOE BIG", category: "General", unit: "piece", quantity: 0, costPrice: 95.0, markup: 30.0, lowStock: 5 },
  { id: "p0032", name: "3PIN SHOE HMA PATHAR", category: "General", unit: "piece", quantity: 0, costPrice: 35.0, markup: 35.0, lowStock: 5 },
  { id: "p0033", name: "3PIN SHOE SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 70.0, markup: 35.0, lowStock: 5 },
  { id: "p0034", name: "40/76 COIL WHITE", category: "General", unit: "piece", quantity: 0, costPrice: 2350.0, markup: 30.0, lowStock: 5 },
  { id: "p0035", name: "4w3ert8;l,", category: "General", unit: "piece", quantity: 0, costPrice: 198.0, markup: 40.0, lowStock: 5 },
  { id: "p0036", name: "5\" NAIL", category: "Hardware & Tools", unit: "kg", quantity: 0, costPrice: 310.0, markup: 40.0, lowStock: 5 },
  { id: "p0037", name: "6 HOLE EXTENTION BOARD", category: "General", unit: "piece", quantity: 0, costPrice: 65.0, markup: 40.0, lowStock: 5 },
  { id: "p0038", name: "6X3 BOX", category: "General", unit: "piece", quantity: 0, costPrice: 83.03, markup: 40.0, lowStock: 5 },
  { id: "p0039", name: "7", category: "General", unit: "piece", quantity: 0, costPrice: 123.0, markup: 30.0, lowStock: 5 },
  { id: "p0040", name: "7/29 CABLE PAK BILAL", category: "Electrical", unit: "piece", quantity: 0, costPrice: 7400.0, markup: 40.0, lowStock: 5 },
  { id: "p0041", name: "7/29 COIL DOUBLE", category: "General", unit: "piece", quantity: 0, costPrice: 9500.0, markup: 30.0, lowStock: 5 },
  { id: "p0042", name: "=", category: "General", unit: "piece", quantity: 0, costPrice: 570.0, markup: 40.0, lowStock: 5 },
  { id: "p0043", name: "AARI BLADE", category: "General", unit: "piece", quantity: 0, costPrice: 29.17, markup: 40.0, lowStock: 5 },
  { id: "p0044", name: "AC BREAKET", category: "General", unit: "pair", quantity: 0, costPrice: 520.0, markup: 40.0, lowStock: 5 },
  { id: "p0045", name: "ACCESSORIES WITH MUSLIM SHOER 3STAR", category: "General", unit: "piece", quantity: 0, costPrice: 6000.0, markup: 40.0, lowStock: 5 },
  { id: "p0046", name: "ACCESSORY SET CHINA", category: "General", unit: "piece", quantity: 0, costPrice: 750.0, markup: 40.0, lowStock: 5 },
  { id: "p0047", name: "ACID", category: "General", unit: "piece", quantity: 0, costPrice: 25.0, markup: 40.0, lowStock: 5 },
  { id: "p0048", name: "ADJUSTABLE SPANNER 12", category: "General", unit: "piece", quantity: 0, costPrice: 675.0, markup: 30.0, lowStock: 5 },
  { id: "p0049", name: "ADJUSTABLE SPANNER 6\"", category: "General", unit: "piece", quantity: 0, costPrice: 350.0, markup: 30.0, lowStock: 5 },
  { id: "p0050", name: "ADJUSTABLE SPANNER 8\"", category: "General", unit: "piece", quantity: 0, costPrice: 475.0, markup: 30.0, lowStock: 5 },
  { id: "p0051", name: "ADJUSTABLE WRENCH  12\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 620.0, markup: 40.0, lowStock: 5 },
  { id: "p0052", name: "ADJUSTABLE WRENCH  8\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 410.0, markup: 40.0, lowStock: 5 },
  { id: "p0053", name: "ADJUSTABLE WRENCH 10\" DIMOND", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 445.0, markup: 40.0, lowStock: 5 },
  { id: "p0054", name: "ADJUSTABLE WRENCH 12\" DIMOND", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 550.0, markup: 40.0, lowStock: 5 },
  { id: "p0055", name: "ADJUSTABLE WRENCH 8\" SMT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 320.0, markup: 40.0, lowStock: 5 },
  { id: "p0056", name: "AIR PUMP", category: "General", unit: "piece", quantity: 0, costPrice: 360.0, markup: 40.0, lowStock: 5 },
  { id: "p0057", name: "AIR PUMP BIG", category: "General", unit: "piece", quantity: 0, costPrice: 370.0, markup: 40.0, lowStock: 5 },
  { id: "p0058", name: "AIR PUMP SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 270.0, markup: 40.0, lowStock: 5 },
  { id: "p0059", name: "AKI DISC 4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 400.0, markup: 30.0, lowStock: 5 },
  { id: "p0060", name: "AKI DISC 5\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 500.0, markup: 30.0, lowStock: 5 },
  { id: "p0061", name: "ALEN KEY", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 210.0, markup: 40.0, lowStock: 5 },
  { id: "p0062", name: "ALMARI BREACKET", category: "General", unit: "piece", quantity: 0, costPrice: 24.0, markup: 40.0, lowStock: 5 },
  { id: "p0063", name: "AMBER", category: "General", unit: "piece", quantity: 0, costPrice: 68.33, markup: 30.0, lowStock: 5 },
  { id: "p0064", name: "ARI BLADE", category: "General", unit: "piece", quantity: 0, costPrice: 24.36, markup: 40.0, lowStock: 5 },
  { id: "p0065", name: "ARI FRAME", category: "General", unit: "piece", quantity: 0, costPrice: 280.0, markup: 40.0, lowStock: 5 },
  { id: "p0066", name: "ARI FRAME HEXSAW FRAME", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 500.0, markup: 30.0, lowStock: 5 },
  { id: "p0067", name: "BABY SOLUTION", category: "Paint", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p0068", name: "BACHAT BOND BIG", category: "Cement & Aggregates", unit: "piece", quantity: 0, costPrice: 680.0, markup: 30.0, lowStock: 5 },
  { id: "p0069", name: "BACHAT BOND SMALL", category: "Cement & Aggregates", unit: "piece", quantity: 0, costPrice: 220.0, markup: 30.0, lowStock: 5 },
  { id: "p0070", name: "BACK COVER 4X4", category: "General", unit: "piece", quantity: 0, costPrice: 23.0, markup: 40.0, lowStock: 5 },
  { id: "p0071", name: "BACK COVER 4X7", category: "General", unit: "piece", quantity: 0, costPrice: 38.0, markup: 40.0, lowStock: 5 },
  { id: "p0072", name: "BADBU DASTANA HAND GLOVES", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 98.33, markup: 40.0, lowStock: 5 },
  { id: "p0073", name: "BAILCHA  4#", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 400.0, markup: 40.0, lowStock: 5 },
  { id: "p0074", name: "BAILCHA  5#", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 417.0, markup: 40.0, lowStock: 5 },
  { id: "p0075", name: "BAILCHA  6#", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 433.0, markup: 40.0, lowStock: 5 },
  { id: "p0076", name: "BAILCHA SHAVOL", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 450.0, markup: 30.0, lowStock: 5 },
  { id: "p0077", name: "BAILCHA SHAVOL 5", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 429.17, markup: 40.0, lowStock: 5 },
  { id: "p0078", name: "BAILCHA SHAVOL 6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 446.33, markup: 40.0, lowStock: 5 },
  { id: "p0079", name: "BALL BALVE 1/2\" IA58", category: "General", unit: "piece", quantity: 0, costPrice: 545.0, markup: 40.0, lowStock: 5 },
  { id: "p0080", name: "BALL FOR FLOAT VALVE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 55.0, markup: 30.0, lowStock: 5 },
  { id: "p0081", name: "BALL VALE 3/4 FAZAL", category: "General", unit: "piece", quantity: 0, costPrice: 680.0, markup: 40.0, lowStock: 5 },
  { id: "p0082", name: "BALL VALVE  1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 275.0, markup: 40.0, lowStock: 5 },
  { id: "p0083", name: "BALL VALVE 1    \" I-A 58", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 975.0, markup: 40.0, lowStock: 5 },
  { id: "p0084", name: "BALL VALVE 1 FAZAL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 680.0, markup: 40.0, lowStock: 5 },
  { id: "p0085", name: "BALL VALVE 1 KITZ", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 780.0, markup: 30.0, lowStock: 5 },
  { id: "p0086", name: "BALL VALVE 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 800.0, markup: 35.0, lowStock: 5 },
  { id: "p0087", name: "BALL VALVE 1\" AFZAL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 680.0, markup: 40.0, lowStock: 5 },
  { id: "p0088", name: "BALL VALVE 1\" KTZ", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 800.0, markup: 35.0, lowStock: 5 },
  { id: "p0089", name: "BALL VALVE 1\" MASTER A58", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 940.0, markup: 40.0, lowStock: 5 },
  { id: "p0090", name: "BALL VALVE 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 390.0, markup: 40.0, lowStock: 5 },
  { id: "p0091", name: "BALL VALVE 1/2  1A58", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 550.0, markup: 40.0, lowStock: 5 },
  { id: "p0092", name: "BALL VALVE 1/2 AFZAL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 345.0, markup: 40.0, lowStock: 5 },
  { id: "p0093", name: "BALL VALVE 1/2 FAZAL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 350.0, markup: 40.0, lowStock: 5 },
  { id: "p0094", name: "BALL VALVE 1/2\"  I-A 58", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 540.0, markup: 30.0, lowStock: 5 },
  { id: "p0095", name: "BALL VALVE 1/2\" I-A 58", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 545.0, markup: 40.0, lowStock: 5 },
  { id: "p0096", name: "BALL VALVE 1/2\" IA58", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 545.0, markup: 40.0, lowStock: 5 },
  { id: "p0097", name: "BALL VALVE 1/2\" MAAZI", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 490.0, markup: 40.0, lowStock: 5 },
  { id: "p0098", name: "BALL VALVE 1/2\" MASTER", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 348.0, markup: 40.0, lowStock: 5 },
  { id: "p0099", name: "BALL VALVE 1/2'IA58", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 440.0, markup: 40.0, lowStock: 5 },
  { id: "p0100", name: "BALL VALVE 25MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p0101", name: "BALL VALVE 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 490.0, markup: 40.0, lowStock: 5 },
  { id: "p0102", name: "BALL VALVE 3/4  IA58", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 655.0, markup: 40.0, lowStock: 5 },
  { id: "p0103", name: "BALL VALVE 3/4 AFZAL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 445.0, markup: 40.0, lowStock: 5 },
  { id: "p0104", name: "BALL VALVE 3/4 AI 58", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 665.0, markup: 40.0, lowStock: 5 },
  { id: "p0105", name: "BALL VALVE 3/4 FAZAL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 450.0, markup: 40.0, lowStock: 5 },
  { id: "p0106", name: "BALL VALVE 3/4 IA58", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 650.0, markup: 40.0, lowStock: 5 },
  { id: "p0107", name: "BALL VALVE 3/4 RT", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 540.0, markup: 35.0, lowStock: 5 },
  { id: "p0108", name: "BALL VALVE 3/4\" I-A 58", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 655.0, markup: 40.0, lowStock: 5 },
  { id: "p0109", name: "BALL VALVE 3/4\" MAAZI", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 590.0, markup: 40.0, lowStock: 5 },
  { id: "p0110", name: "BALL VALVE 3/4\" MASTER", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 310.0, markup: 40.0, lowStock: 5 },
  { id: "p0111", name: "BALL VALVE 32MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 240.0, markup: 40.0, lowStock: 5 },
  { id: "p0112", name: "BALL VALVE A58   3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 610.0, markup: 40.0, lowStock: 5 },
  { id: "p0113", name: "BALL VALVE A58 MASTER 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 435.0, markup: 30.0, lowStock: 5 },
  { id: "p0114", name: "BALL VALVE A58 MASTER 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 480.0, markup: 30.0, lowStock: 5 },
  { id: "p0115", name: "BALL VALVE AA 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 395.0, markup: 30.0, lowStock: 5 },
  { id: "p0116", name: "BALL VALVE AFZAL 1/2'", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 380.0, markup: 40.0, lowStock: 5 },
  { id: "p0117", name: "BALL VALVE FAZAL 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 365.0, markup: 40.0, lowStock: 5 },
  { id: "p0118", name: "BALL VALVE FAZAL 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 430.0, markup: 40.0, lowStock: 5 },
  { id: "p0119", name: "BALL VALVE FAZAL NORMAL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 280.0, markup: 40.0, lowStock: 5 },
  { id: "p0120", name: "BALL VALVE GOLD 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 390.0, markup: 40.0, lowStock: 5 },
  { id: "p0121", name: "BALL VALVE GOLD 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 490.0, markup: 40.0, lowStock: 5 },
  { id: "p0122", name: "BALL VALVE RBS 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 380.0, markup: 40.0, lowStock: 5 },
  { id: "p0123", name: "BALL VALVE RBS 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 490.0, markup: 40.0, lowStock: 5 },
  { id: "p0124", name: "BALL VALVE SS 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 800.0, markup: 30.0, lowStock: 5 },
  { id: "p0125", name: "BALL VALVE SS 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 450.0, markup: 30.0, lowStock: 5 },
  { id: "p0126", name: "BALL VALVE SS 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 550.0, markup: 30.0, lowStock: 5 },
  { id: "p0127", name: "BALL VALVE TPS 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 730.0, markup: 40.0, lowStock: 5 },
  { id: "p0128", name: "BALL VALVE TPS 1/2\"", category: "Plumbing & Sanitary", unit: "370", quantity: 0, costPrice: 370.0, markup: 40.0, lowStock: 5 },
  { id: "p0129", name: "BALL VALVE TPS 3/4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 440.0, markup: 40.0, lowStock: 5 },
  { id: "p0130", name: "BASIN", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 2800.0, markup: 40.0, lowStock: 5 },
  { id: "p0131", name: "BASIN BIG", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 3600.0, markup: 40.0, lowStock: 5 },
  { id: "p0132", name: "BASIN NORMAL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 2800.0, markup: 40.0, lowStock: 5 },
  { id: "p0133", name: "BASIN WAIST", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 185.0, markup: 40.0, lowStock: 5 },
  { id: "p0134", name: "BASON MIXER", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1750.0, markup: 40.0, lowStock: 5 },
  { id: "p0135", name: "BASON MIXER CHINA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1250.0, markup: 30.0, lowStock: 5 },
  { id: "p0136", name: "BATH ROOM SET DOME 3STAR", category: "General", unit: "piece", quantity: 0, costPrice: 17000.0, markup: 40.0, lowStock: 5 },
  { id: "p0137", name: "BATHROOM LIGHT 1FT", category: "General", unit: "piece", quantity: 0, costPrice: 550.0, markup: 40.0, lowStock: 5 },
  { id: "p0138", name: "BATHROOM LIGHT 2FT", category: "General", unit: "piece", quantity: 0, costPrice: 600.0, markup: 40.0, lowStock: 5 },
  { id: "p0139", name: "BED SWITCH", category: "Electrical", unit: "piece", quantity: 0, costPrice: 45.0, markup: 35.0, lowStock: 5 },
  { id: "p0140", name: "BELL", category: "General", unit: "piece", quantity: 0, costPrice: 190.0, markup: 40.0, lowStock: 5 },
  { id: "p0141", name: "BELL BUTTON . PUSH BUTTON", category: "General", unit: "piece", quantity: 0, costPrice: 37.0, markup: 35.0, lowStock: 5 },
  { id: "p0142", name: "BELL PUSH DOUBLE", category: "General", unit: "piece", quantity: 0, costPrice: 170.0, markup: 40.0, lowStock: 5 },
  { id: "p0143", name: "BELL PUSH SINGLE", category: "General", unit: "piece", quantity: 0, costPrice: 140.0, markup: 40.0, lowStock: 5 },
  { id: "p0144", name: "BEND 1\"", category: "General", unit: "piece", quantity: 0, costPrice: 12.85, markup: 40.0, lowStock: 5 },
  { id: "p0145", name: "BEND 1-1/4\"", category: "General", unit: "piece", quantity: 0, costPrice: 85.0, markup: 40.0, lowStock: 5 },
  { id: "p0146", name: "BEND 3/4\"", category: "General", unit: "piece", quantity: 0, costPrice: 10.76, markup: 40.0, lowStock: 5 },
  { id: "p0147", name: "BEND 3GM 3/4", category: "General", unit: "piece", quantity: 0, costPrice: 19.79, markup: 40.0, lowStock: 5 },
  { id: "p0148", name: "BEND GM 1\"", category: "General", unit: "piece", quantity: 0, costPrice: 28.89, markup: 40.0, lowStock: 5 },
  { id: "p0149", name: "BIB COCK", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 690.0, markup: 30.0, lowStock: 5 },
  { id: "p0150", name: "BIB COCK 3STAR", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 950.0, markup: 40.0, lowStock: 5 },
  { id: "p0151", name: "BIB COCK AONE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 690.0, markup: 35.0, lowStock: 5 },
  { id: "p0152", name: "BIB COCK LIGHT", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 480.0, markup: 30.0, lowStock: 5 },
  { id: "p0153", name: "BIB COCK MASTER NICE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 580.0, markup: 40.0, lowStock: 5 },
  { id: "p0154", name: "BIB COCK PVC", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 280.0, markup: 40.0, lowStock: 5 },
  { id: "p0155", name: "BIB COCK RBS", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 810.0, markup: 40.0, lowStock: 5 },
  { id: "p0156", name: "BIB COCK RT", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 720.0, markup: 35.0, lowStock: 5 },
  { id: "p0157", name: "BIB COCK S ASIA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 435.0, markup: 40.0, lowStock: 5 },
  { id: "p0158", name: "BIB COCK SMALL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 520.0, markup: 30.0, lowStock: 5 },
  { id: "p0159", name: "BIB COCK STEEL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 420.0, markup: 40.0, lowStock: 5 },
  { id: "p0160", name: "BIB COCK STEEL ASIA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 630.0, markup: 40.0, lowStock: 5 },
  { id: "p0161", name: "BIB COCK TPS", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 665.0, markup: 40.0, lowStock: 5 },
  { id: "p0162", name: "BIB COCKASIA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 410.0, markup: 40.0, lowStock: 5 },
  { id: "p0163", name: "BIB COK HEAD", category: "General", unit: "piece", quantity: 0, costPrice: 48.0, markup: 40.0, lowStock: 5 },
  { id: "p0164", name: "BIB COK SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 540.0, markup: 30.0, lowStock: 5 },
  { id: "p0165", name: "BIB OCK ASIA", category: "General", unit: "piece", quantity: 0, costPrice: 430.0, markup: 40.0, lowStock: 5 },
  { id: "p0166", name: "BIBCOCK DOUBLE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1180.0, markup: 30.0, lowStock: 5 },
  { id: "p0167", name: "BIG 1/2 KG SOLUTION", category: "Paint", unit: "piece", quantity: 0, costPrice: 350.0, markup: 40.0, lowStock: 5 },
  { id: "p0168", name: "BIG 5 DISC 4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 320.0, markup: 40.0, lowStock: 5 },
  { id: "p0169", name: "BIG 5 DISC 7\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 640.0, markup: 40.0, lowStock: 5 },
  { id: "p0170", name: "BINDING WIRE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 350.0, markup: 30.0, lowStock: 5 },
  { id: "p0171", name: "BLACK AMBER", category: "General", unit: "piece", quantity: 0, costPrice: 95.8, markup: 40.0, lowStock: 5 },
  { id: "p0172", name: "BLACK SOCKET BIG", category: "Electrical", unit: "piece", quantity: 0, costPrice: 123.33, markup: 30.0, lowStock: 5 },
  { id: "p0173", name: "BLACK SOCKET SMALL", category: "Electrical", unit: "piece", quantity: 0, costPrice: 65.0, markup: 30.0, lowStock: 5 },
  { id: "p0174", name: "BLACK UMBER", category: "General", unit: "piece", quantity: 0, costPrice: 94.42, markup: 40.0, lowStock: 5 },
  { id: "p0175", name: "BLUE COUPLER 25MM", category: "General", unit: "piece", quantity: 0, costPrice: 210.0, markup: 35.0, lowStock: 5 },
  { id: "p0176", name: "BLUE COUPLER 32MM", category: "General", unit: "piece", quantity: 0, costPrice: 273.0, markup: 35.0, lowStock: 5 },
  { id: "p0177", name: "BLUE COUPLER 32X25MM", category: "General", unit: "piece", quantity: 0, costPrice: 280.0, markup: 35.0, lowStock: 5 },
  { id: "p0178", name: "BOARD 3 HOLE", category: "General", unit: "piece", quantity: 0, costPrice: 35.0, markup: 40.0, lowStock: 5 },
  { id: "p0179", name: "BOARD 3X3", category: "General", unit: "piece", quantity: 0, costPrice: 60.0, markup: 30.0, lowStock: 5 },
  { id: "p0180", name: "BOARD 4 HOLE", category: "General", unit: "piece", quantity: 0, costPrice: 60.0, markup: 40.0, lowStock: 5 },
  { id: "p0181", name: "BOLT KIT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 82.0, markup: 30.0, lowStock: 5 },
  { id: "p0182", name: "BOX 3X3", category: "General", unit: "6000", quantity: 0, costPrice: 56.4, markup: 30.0, lowStock: 5 },
  { id: "p0183", name: "BOX 6X3", category: "General", unit: "5000", quantity: 0, costPrice: 84.0, markup: 30.0, lowStock: 5 },
  { id: "p0184", name: "BOX GM 3X3", category: "General", unit: "piece", quantity: 0, costPrice: 47.92, markup: 40.0, lowStock: 5 },
  { id: "p0185", name: "BOX GM 6X3", category: "General", unit: "piece", quantity: 0, costPrice: 68.0, markup: 40.0, lowStock: 5 },
  { id: "p0186", name: "BOXES 3X3", category: "General", unit: "piece", quantity: 0, costPrice: 30.0, markup: 40.0, lowStock: 5 },
  { id: "p0187", name: "BRACKET 3\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 12.0, markup: 40.0, lowStock: 5 },
  { id: "p0188", name: "BREAKAR SINGLE", category: "General", unit: "piece", quantity: 0, costPrice: 275.0, markup: 40.0, lowStock: 5 },
  { id: "p0189", name: "BREAKER 10-16-20AMP JM", category: "Electrical", unit: "piece", quantity: 0, costPrice: 275.0, markup: 40.0, lowStock: 5 },
  { id: "p0190", name: "BREAKER 3 PHASE 63AMP", category: "Electrical", unit: "piece", quantity: 0, costPrice: 1200.0, markup: 40.0, lowStock: 5 },
  { id: "p0191", name: "BREAKER 63AMP DOUBLE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 560.0, markup: 30.0, lowStock: 5 },
  { id: "p0192", name: "BREAKER AMSONS 20AMP", category: "Electrical", unit: "piece", quantity: 0, costPrice: 380.0, markup: 40.0, lowStock: 5 },
  { id: "p0193", name: "BREAKER DABBI COVER", category: "Electrical", unit: "piece", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p0194", name: "BREAKER DOUBLE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 570.0, markup: 40.0, lowStock: 5 },
  { id: "p0195", name: "BREAKER KANGGI", category: "Electrical", unit: "piece", quantity: 0, costPrice: 280.0, markup: 40.0, lowStock: 5 },
  { id: "p0196", name: "BREAKER KANGI", category: "Electrical", unit: "piece", quantity: 0, costPrice: 440.0, markup: 40.0, lowStock: 5 },
  { id: "p0197", name: "BREAKER SINGLE 10 & 20 AMP", category: "Electrical", unit: "piece", quantity: 0, costPrice: 275.0, markup: 40.0, lowStock: 5 },
  { id: "p0198", name: "BRUSH ANGLE WALA", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 250.0, markup: 40.0, lowStock: 5 },
  { id: "p0199", name: "BULB 05W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p0200", name: "BULB 07W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p0201", name: "BULB 12W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 120.0, markup: 30.0, lowStock: 5 },
  { id: "p0202", name: "BULB 150W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 75.0, markup: 30.0, lowStock: 5 },
  { id: "p0203", name: "BULB 18W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 265.0, markup: 30.0, lowStock: 5 },
  { id: "p0204", name: "BULB 50W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 1030.0, markup: 30.0, lowStock: 5 },
  { id: "p0205", name: "BULB HOLDER", category: "Electrical", unit: "piece", quantity: 0, costPrice: 0.0, markup: 40.0, lowStock: 5 },
  { id: "p0206", name: "BULB HOLDER HANGUNG", category: "Electrical", unit: "piece", quantity: 0, costPrice: 30.0, markup: 40.0, lowStock: 5 },
  { id: "p0207", name: "BULB HOLDER WALL MOUNTED SMALL", category: "Electrical", unit: "piece", quantity: 0, costPrice: 40.0, markup: 40.0, lowStock: 5 },
  { id: "p0208", name: "BULB HOLER WALL MOUNTED SMALL", category: "Electrical", unit: "piece", quantity: 0, costPrice: 60.0, markup: 35.0, lowStock: 5 },
  { id: "p0209", name: "BUTTON BLACK BIG", category: "General", unit: "piece", quantity: 0, costPrice: 91.67, markup: 40.0, lowStock: 5 },
  { id: "p0210", name: "CABLE 2 CORE 3/29 ( 5500 & 5100) 90MTR", category: "Electrical", unit: "piece", quantity: 0, costPrice: 0.0, markup: 40.0, lowStock: 5 },
  { id: "p0211", name: "CABLE 2 CORE 7/29 ( 7000-8000 & 9800) 90MTR", category: "Electrical", unit: "piece", quantity: 0, costPrice: 0.0, markup: 40.0, lowStock: 5 },
  { id: "p0212", name: "CABLE 23/76", category: "Electrical", unit: "piece", quantity: 0, costPrice: 2650.0, markup: 40.0, lowStock: 5 },
  { id: "p0213", name: "CABLE 23/76 WHITE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 2500.0, markup: 30.0, lowStock: 5 },
  { id: "p0214", name: "CABLE 2COR 7/29", category: "Electrical", unit: "piece", quantity: 0, costPrice: 11500.0, markup: 35.0, lowStock: 5 },
  { id: "p0215", name: "CABLE 3/29 PAK BILAL", category: "Electrical", unit: "piece", quantity: 0, costPrice: 3400.0, markup: 30.0, lowStock: 5 },
  { id: "p0216", name: "CABLE 3/29 PAK ROSHAN", category: "Electrical", unit: "piece", quantity: 0, costPrice: 4500.0, markup: 35.0, lowStock: 5 },
  { id: "p0217", name: "CABLE 3/29 SINGLE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 4450.0, markup: 30.0, lowStock: 5 },
  { id: "p0218", name: "CABLE 40/76 YELLOW", category: "Electrical", unit: "piece", quantity: 0, costPrice: 5500.0, markup: 30.0, lowStock: 5 },
  { id: "p0219", name: "CABLE 47/29 PAK ROSHAN", category: "Electrical", unit: "piece", quantity: 0, costPrice: 9200.0, markup: 35.0, lowStock: 5 },
  { id: "p0220", name: "CABLE 7/29 PAK BILAL", category: "Electrical", unit: "piece", quantity: 0, costPrice: 7500.0, markup: 30.0, lowStock: 5 },
  { id: "p0221", name: "CABLE 7/29 SINGLE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 9000.0, markup: 30.0, lowStock: 5 },
  { id: "p0222", name: "CABLE CUTTER 8", category: "Electrical", unit: "piece", quantity: 0, costPrice: 580.0, markup: 30.0, lowStock: 5 },
  { id: "p0223", name: "CABLE ORANGE 40/76", category: "Electrical", unit: "piece", quantity: 0, costPrice: 4850.0, markup: 40.0, lowStock: 5 },
  { id: "p0224", name: "CABLE SINGLE 3/29", category: "Electrical", unit: "piece", quantity: 0, costPrice: 4400.0, markup: 35.0, lowStock: 5 },
  { id: "p0225", name: "CABLE TIE   10\"", category: "Electrical", unit: "piece", quantity: 0, costPrice: 240.0, markup: 40.0, lowStock: 5 },
  { id: "p0226", name: "CABLE TIE   6\"", category: "Electrical", unit: "piece", quantity: 0, costPrice: 130.0, markup: 40.0, lowStock: 5 },
  { id: "p0227", name: "CABLE TIE   8\"", category: "Electrical", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p0228", name: "CABLE TIE  10", category: "Electrical", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p0229", name: "CABLE TIE  12", category: "Electrical", unit: "piece", quantity: 0, costPrice: 260.0, markup: 40.0, lowStock: 5 },
  { id: "p0230", name: "CABLE TIE  6", category: "Electrical", unit: "piece", quantity: 0, costPrice: 130.0, markup: 40.0, lowStock: 5 },
  { id: "p0231", name: "CABLE TIE  8", category: "Electrical", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p0232", name: "CABLE TIE 10", category: "Electrical", unit: "piece", quantity: 0, costPrice: 250.0, markup: 35.0, lowStock: 5 },
  { id: "p0233", name: "CABLE TIE 10#", category: "Electrical", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p0234", name: "CABLE TIE 12", category: "Electrical", unit: "piece", quantity: 0, costPrice: 270.0, markup: 35.0, lowStock: 5 },
  { id: "p0235", name: "CABLE TIE 14", category: "Electrical", unit: "piece", quantity: 0, costPrice: 330.0, markup: 35.0, lowStock: 5 },
  { id: "p0236", name: "CABLE TIE 16", category: "Electrical", unit: "piece", quantity: 0, costPrice: 420.0, markup: 35.0, lowStock: 5 },
  { id: "p0237", name: "CABLE TIE 4", category: "Electrical", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p0238", name: "CABLE TIE 6\"", category: "Electrical", unit: "piece", quantity: 0, costPrice: 130.0, markup: 40.0, lowStock: 5 },
  { id: "p0239", name: "CABLE TIE 8\"", category: "Electrical", unit: "piece", quantity: 0, costPrice: 150.0, markup: 35.0, lowStock: 5 },
  { id: "p0240", name: "CABLE TIE ORANGE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 350.0, markup: 40.0, lowStock: 5 },
  { id: "p0241", name: "CABLE TIE YELLOW", category: "Electrical", unit: "piece", quantity: 0, costPrice: 290.0, markup: 40.0, lowStock: 5 },
  { id: "p0242", name: "CAPACITOR  3.5", category: "Electrical", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p0243", name: "CAPACITOR 138-182", category: "Electrical", unit: "piece", quantity: 0, costPrice: 190.0, markup: 40.0, lowStock: 5 },
  { id: "p0244", name: "CAPACITOR 200-250", category: "Electrical", unit: "piece", quantity: 0, costPrice: 230.0, markup: 40.0, lowStock: 5 },
  { id: "p0245", name: "CAPACITOR 250-300", category: "Electrical", unit: "piece", quantity: 0, costPrice: 280.0, markup: 40.0, lowStock: 5 },
  { id: "p0246", name: "CAPACITOR 35UF", category: "Electrical", unit: "piece", quantity: 0, costPrice: 400.0, markup: 35.0, lowStock: 5 },
  { id: "p0247", name: "CAPACITOR 80-110", category: "Electrical", unit: "piece", quantity: 0, costPrice: 140.0, markup: 40.0, lowStock: 5 },
  { id: "p0248", name: "CAPACITOR CF", category: "Electrical", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p0249", name: "CAPASITOR    25FCU", category: "General", unit: "piece", quantity: 0, costPrice: 340.0, markup: 30.0, lowStock: 5 },
  { id: "p0250", name: "CAPASITOR    30FCU", category: "General", unit: "piece", quantity: 0, costPrice: 360.0, markup: 30.0, lowStock: 5 },
  { id: "p0251", name: "CAPASITOR    35FCU", category: "General", unit: "piece", quantity: 0, costPrice: 410.0, markup: 30.0, lowStock: 5 },
  { id: "p0252", name: "CAPASITOR    40FCU", category: "General", unit: "piece", quantity: 0, costPrice: 440.0, markup: 30.0, lowStock: 5 },
  { id: "p0253", name: "CAPASITOR    45FCU", category: "General", unit: "piece", quantity: 0, costPrice: 480.0, markup: 30.0, lowStock: 5 },
  { id: "p0254", name: "CAPASITOR    50FCU", category: "General", unit: "piece", quantity: 0, costPrice: 520.0, markup: 30.0, lowStock: 5 },
  { id: "p0255", name: "CAPITAL CABLE 3/29.", category: "Electrical", unit: "piece", quantity: 0, costPrice: 4450.0, markup: 30.0, lowStock: 5 },
  { id: "p0256", name: "CAPITAL CABLE 7/29.", category: "Electrical", unit: "piece", quantity: 0, costPrice: 9000.0, markup: 30.0, lowStock: 5 },
  { id: "p0257", name: "CAR POLISH", category: "General", unit: "piece", quantity: 0, costPrice: 250.0, markup: 40.0, lowStock: 5 },
  { id: "p0258", name: "CARBON", category: "General", unit: "piece", quantity: 0, costPrice: 33.0, markup: 40.0, lowStock: 5 },
  { id: "p0259", name: "CARBON 104", category: "General", unit: "piece", quantity: 0, costPrice: 47.5, markup: 40.0, lowStock: 5 },
  { id: "p0260", name: "CARBON 21 N 50", category: "General", unit: "piece", quantity: 0, costPrice: 42.5, markup: 40.0, lowStock: 5 },
  { id: "p0261", name: "CARBON BRUSHES", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 600.0, markup: 30.0, lowStock: 5 },
  { id: "p0262", name: "CARBON BUSH", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 25.0, markup: 40.0, lowStock: 5 },
  { id: "p0263", name: "CARBON HITACHI 152-44", category: "General", unit: "piece", quantity: 0, costPrice: 50.0, markup: 40.0, lowStock: 5 },
  { id: "p0264", name: "CARBON HITACHI 21-50NO", category: "General", unit: "piece", quantity: 0, costPrice: 45.0, markup: 40.0, lowStock: 5 },
  { id: "p0265", name: "CC TV CABLE  1X90 MTR", category: "Electrical", unit: "piece", quantity: 0, costPrice: 2050.0, markup: 40.0, lowStock: 5 },
  { id: "p0266", name: "CELIND HOLDER", category: "Electrical", unit: "piece", quantity: 0, costPrice: 60.0, markup: 40.0, lowStock: 5 },
  { id: "p0267", name: "CELING EXHAUST FAN", category: "General", unit: "piece", quantity: 0, costPrice: 3150.0, markup: 40.0, lowStock: 5 },
  { id: "p0268", name: "CENTRO BOND", category: "Cement & Aggregates", unit: "piece", quantity: 0, costPrice: 480.0, markup: 30.0, lowStock: 5 },
  { id: "p0269", name: "CHALK", category: "General", unit: "piece", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p0270", name: "CHAND TARA SINGLE", category: "General", unit: "piece", quantity: 0, costPrice: 110.0, markup: 40.0, lowStock: 5 },
  { id: "p0271", name: "CHANGE OVER", category: "General", unit: "piece", quantity: 0, costPrice: 975.0, markup: 35.0, lowStock: 5 },
  { id: "p0272", name: "CHECK NUT 1\" TANKI BUSH", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 165.0, markup: 40.0, lowStock: 5 },
  { id: "p0273", name: "CHECK NUT 3/4\" TANKI BUSH", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p0274", name: "CHECK NUT TANKI BUSH", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 190.0, markup: 40.0, lowStock: 5 },
  { id: "p0275", name: "CHECK VALVE 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 360.0, markup: 40.0, lowStock: 5 },
  { id: "p0276", name: "CHECK VALVE BRASS", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 380.0, markup: 40.0, lowStock: 5 },
  { id: "p0277", name: "CHECK VALVE BRASS 1", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 475.0, markup: 35.0, lowStock: 5 },
  { id: "p0278", name: "CHESIL", category: "General", unit: "piece", quantity: 0, costPrice: 242.5, markup: 40.0, lowStock: 5 },
  { id: "p0279", name: "CHESIL 10\"", category: "General", unit: "piece", quantity: 0, costPrice: 300.0, markup: 40.0, lowStock: 5 },
  { id: "p0280", name: "CHESIL 12\"", category: "General", unit: "piece", quantity: 0, costPrice: 362.0, markup: 40.0, lowStock: 5 },
  { id: "p0281", name: "CHESIL CHAINI DEGI", category: "General", unit: "piece", quantity: 0, costPrice: 187.0, markup: 40.0, lowStock: 5 },
  { id: "p0282", name: "CHESIL WITH GUARD 10\"", category: "General", unit: "piece", quantity: 0, costPrice: 350.0, markup: 40.0, lowStock: 5 },
  { id: "p0283", name: "CHESIL WITH GUARD 8\"", category: "General", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p0284", name: "CHINA MAGIC", category: "General", unit: "piece", quantity: 0, costPrice: 55.0, markup: 30.0, lowStock: 5 },
  { id: "p0285", name: "CHOORI HOOCK HEAVY", category: "General", unit: "piece", quantity: 0, costPrice: 64.61, markup: 40.0, lowStock: 5 },
  { id: "p0286", name: "CHOORI HOOCK NORMAL", category: "General", unit: "piece", quantity: 0, costPrice: 45.0, markup: 40.0, lowStock: 5 },
  { id: "p0287", name: "CHUTKI VALVE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 380.0, markup: 35.0, lowStock: 5 },
  { id: "p0288", name: "CLAW HAMMER AKI", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 800.0, markup: 40.0, lowStock: 5 },
  { id: "p0289", name: "CLAW HAMMER AKT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 730.0, markup: 40.0, lowStock: 5 },
  { id: "p0290", name: "CLAW HAMMER ATC", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 520.0, markup: 30.0, lowStock: 5 },
  { id: "p0291", name: "CLAW HAMMER CHROME", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 300.0, markup: 40.0, lowStock: 5 },
  { id: "p0292", name: "CLAW HAMMER HTC", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 800.0, markup: 40.0, lowStock: 5 },
  { id: "p0293", name: "CLEANING BRUSH", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 230.0, markup: 40.0, lowStock: 5 },
  { id: "p0294", name: "CLEANING CLOTH", category: "General", unit: "piece", quantity: 0, costPrice: 37.5, markup: 40.0, lowStock: 5 },
  { id: "p0295", name: "CLEANING CLOTH BIG", category: "General", unit: "piece", quantity: 0, costPrice: 38.33, markup: 40.0, lowStock: 5 },
  { id: "p0296", name: "CLEANING CLOTH SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 23.33, markup: 40.0, lowStock: 5 },
  { id: "p0297", name: "CLEANING CLOTH yellow", category: "General", unit: "piece", quantity: 0, costPrice: 41.67, markup: 30.0, lowStock: 5 },
  { id: "p0298", name: "CLEANING COTH COLOURED", category: "Paint", unit: "piece", quantity: 0, costPrice: 37.5, markup: 30.0, lowStock: 5 },
  { id: "p0299", name: "CLEAR TAPE 2", category: "General", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p0300", name: "CLOTH  TAPE DOUBLE", category: "General", unit: "piece", quantity: 0, costPrice: 120.0, markup: 35.0, lowStock: 5 },
  { id: "p0301", name: "CLOTH DOORI", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p0302", name: "CLOTH ROPE", category: "General", unit: "piece", quantity: 0, costPrice: 180.0, markup: 35.0, lowStock: 5 },
  { id: "p0303", name: "CLOTH TAPE", category: "General", unit: "piece", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p0304", name: "CNC 6X6 JALI HEAVY DUTY", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 580.0, markup: 30.0, lowStock: 5 },
  { id: "p0305", name: "COLOUR 1/4", category: "Paint", unit: "piece", quantity: 0, costPrice: 130.0, markup: 40.0, lowStock: 5 },
  { id: "p0306", name: "COLOUR TUBE 687", category: "Paint", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p0307", name: "COMBINATION SPANNER  19X19", category: "General", unit: "piece", quantity: 0, costPrice: 300.0, markup: 40.0, lowStock: 5 },
  { id: "p0308", name: "COMBINATION SPANNER 10MM", category: "General", unit: "piece", quantity: 0, costPrice: 125.0, markup: 40.0, lowStock: 5 },
  { id: "p0309", name: "COMBINATION SPANNER 11", category: "General", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p0310", name: "COMBINATION SPANNER 13", category: "General", unit: "piece", quantity: 0, costPrice: 143.0, markup: 40.0, lowStock: 5 },
  { id: "p0311", name: "COMBINATION SPANNER 13X13", category: "General", unit: "piece", quantity: 0, costPrice: 170.0, markup: 40.0, lowStock: 5 },
  { id: "p0312", name: "COMBINATION SPANNER 14", category: "General", unit: "piece", quantity: 0, costPrice: 173.33, markup: 40.0, lowStock: 5 },
  { id: "p0313", name: "COMBINATION SPANNER 14X15", category: "General", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p0314", name: "COMBINATION SPANNER 15", category: "General", unit: "piece", quantity: 0, costPrice: 185.0, markup: 40.0, lowStock: 5 },
  { id: "p0315", name: "COMBINATION SPANNER 16X16", category: "General", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p0316", name: "COMMOD KIT", category: "General", unit: "piece", quantity: 0, costPrice: 40.0, markup: 40.0, lowStock: 5 },
  { id: "p0317", name: "COMMOD WASHER WHITE", category: "General", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p0318", name: "COMMODE 101 3STAR", category: "General", unit: "piece", quantity: 0, costPrice: 18500.0, markup: 40.0, lowStock: 5 },
  { id: "p0319", name: "COMMODE KIT", category: "General", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p0320", name: "COMMODE WASHER BLACK", category: "General", unit: "piece", quantity: 0, costPrice: 25.0, markup: 40.0, lowStock: 5 },
  { id: "p0321", name: "CONCRETE BIT 1/4X4", category: "General", unit: "piece", quantity: 0, costPrice: 40.0, markup: 40.0, lowStock: 5 },
  { id: "p0322", name: "CONCRETE BIT 3/16 DREXEL", category: "General", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p0323", name: "CONCRETE BIT 3/8", category: "General", unit: "piece", quantity: 0, costPrice: 40.0, markup: 40.0, lowStock: 5 },
  { id: "p0324", name: "CONCRETE DRILL BIT 8MM", category: "General", unit: "piece", quantity: 0, costPrice: 40.0, markup: 40.0, lowStock: 5 },
  { id: "p0325", name: "CONCRETE GRINDING DISC 4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 480.0, markup: 35.0, lowStock: 5 },
  { id: "p0326", name: "CONCRETE GRINDING DISC 5\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 580.0, markup: 35.0, lowStock: 5 },
  { id: "p0327", name: "CONDIIT PIPE 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p0328", name: "CONDIIT PIPE 3/4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 66.27, markup: 40.0, lowStock: 5 },
  { id: "p0329", name: "CONDUIT PIIPE 3/4\" 28.700", category: "General", unit: "piece", quantity: 0, costPrice: 68.89, markup: 40.0, lowStock: 5 },
  { id: "p0330", name: "CONDUIT PIPE   1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 87.62, markup: 40.0, lowStock: 5 },
  { id: "p0331", name: "CONDUIT PIPE 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 86.7, markup: 40.0, lowStock: 5 },
  { id: "p0332", name: "CONDUIT PIPE 1\"     20.500", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 0.0, markup: 40.0, lowStock: 5 },
  { id: "p0333", name: "CONDUIT PIPE 1\"     VICTOR", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 280.0, markup: 40.0, lowStock: 5 },
  { id: "p0334", name: "CONDUIT PIPE 1\" 22KG ROOF    240", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 105.0, markup: 40.0, lowStock: 5 },
  { id: "p0335", name: "CONDUIT PIPE 1\" 33 OR 34 KG FOR WALL @ 245", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 0.0, markup: 40.0, lowStock: 5 },
  { id: "p0336", name: "CONDUIT PIPE 1\" 37 OR 38 KG FOR ROOF @ 245", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 0.0, markup: 40.0, lowStock: 5 },
  { id: "p0337", name: "CONDUIT PIPE 1\" DEEVAR", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 86.0, markup: 40.0, lowStock: 5 },
  { id: "p0338", name: "CONDUIT PIPE 1\" HEAVY", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 98.0, markup: 40.0, lowStock: 5 },
  { id: "p0339", name: "CONDUIT PIPE 1\" WHITE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 250.0, markup: 35.0, lowStock: 5 },
  { id: "p0340", name: "CONDUIT PIPE 3/4      30.70KG=100 PCS", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 85.0, markup: 40.0, lowStock: 5 },
  { id: "p0341", name: "CONDUIT PIPE 3/4 BLUE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 51.0, markup: 40.0, lowStock: 5 },
  { id: "p0342", name: "CONDUIT PIPE 3/4 HEAVY", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 62.21, markup: 40.0, lowStock: 5 },
  { id: "p0343", name: "CONDUIT PIPE 3/4 RED", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 72.62, markup: 40.0, lowStock: 5 },
  { id: "p0344", name: "CONDUIT PIPE 3/4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 58.14, markup: 40.0, lowStock: 5 },
  { id: "p0345", name: "CONDUIT PIPE 3/4\"      VICTOR", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 195.0, markup: 40.0, lowStock: 5 },
  { id: "p0346", name: "CONDUIT PIPE 3/4\" 24 OR 25 KG FOR WALL @ 245", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 0.0, markup: 40.0, lowStock: 5 },
  { id: "p0347", name: "CONDUIT PIPE 3/4\" 27 OR 28 KG FOR ROOF @ 245", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 0.0, markup: 40.0, lowStock: 5 },
  { id: "p0348", name: "CONDUIT PIPE 3/4\" 29.600 ROOF 240", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 68.88, markup: 40.0, lowStock: 5 },
  { id: "p0349", name: "CONDUIT PIPE 3/4\" DEEVAR", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 56.0, markup: 40.0, lowStock: 5 },
  { id: "p0350", name: "CONDUIT PIPE 3/4\" WHITE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 290.0, markup: 35.0, lowStock: 5 },
  { id: "p0351", name: "CONDUIT PIPE DEEVAR 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 75.0, markup: 40.0, lowStock: 5 },
  { id: "p0352", name: "CONDUIT PIPE DEEVAR 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 60.0, markup: 40.0, lowStock: 5 },
  { id: "p0353", name: "CONDUIT PIPIE 1\"       40 KG    =  100 PCS", category: "General", unit: "piece", quantity: 0, costPrice: 110.0, markup: 40.0, lowStock: 5 },
  { id: "p0354", name: "CONNECTION PIIPE 1MTR", category: "General", unit: "piece", quantity: 0, costPrice: 125.0, markup: 30.0, lowStock: 5 },
  { id: "p0355", name: "CONNECTION PIIPE 2 FT", category: "General", unit: "piece", quantity: 0, costPrice: 125.0, markup: 30.0, lowStock: 5 },
  { id: "p0356", name: "CONNECTION PIPE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 135.0, markup: 35.0, lowStock: 5 },
  { id: "p0357", name: "CONNECTION PIPE  24\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p0358", name: "CONNECTION PIPE 18\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 70.0, markup: 30.0, lowStock: 5 },
  { id: "p0359", name: "CONNECTION PIPE 1MTR", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 120.0, markup: 30.0, lowStock: 5 },
  { id: "p0360", name: "CONNECTION PIPE 24\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 95.0, markup: 30.0, lowStock: 5 },
  { id: "p0361", name: "CONNECTION PIPE 2FT", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p0362", name: "CONNECTION PIPE 3FT", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 115.0, markup: 40.0, lowStock: 5 },
  { id: "p0363", name: "COTON HAND GLOVES", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 20.0, markup: 40.0, lowStock: 5 },
  { id: "p0364", name: "CP BODY MUSLIM SHOWER BODY", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 165.0, markup: 40.0, lowStock: 5 },
  { id: "p0365", name: "CP CHAIN", category: "General", unit: "piece", quantity: 0, costPrice: 680.0, markup: 30.0, lowStock: 5 },
  { id: "p0366", name: "CP CHAIN  H DUTY", category: "General", unit: "piece", quantity: 0, costPrice: 490.0, markup: 40.0, lowStock: 5 },
  { id: "p0367", name: "CP CHAIN BHARI 1MTR", category: "General", unit: "piece", quantity: 0, costPrice: 210.0, markup: 40.0, lowStock: 5 },
  { id: "p0368", name: "CP CHAIN CHAINA", category: "General", unit: "piece", quantity: 0, costPrice: 225.0, markup: 40.0, lowStock: 5 },
  { id: "p0369", name: "CP CHAIN FOR MUSLIM SHOWE", category: "General", unit: "piece", quantity: 0, costPrice: 240.0, markup: 40.0, lowStock: 5 },
  { id: "p0370", name: "CP CHAIN HEAVY DUTY", category: "General", unit: "piece", quantity: 0, costPrice: 490.0, markup: 40.0, lowStock: 5 },
  { id: "p0371", name: "CP CHAIN HTC", category: "General", unit: "piece", quantity: 0, costPrice: 280.0, markup: 40.0, lowStock: 5 },
  { id: "p0372", name: "CP CHAIN SINGLE P MASTER", category: "General", unit: "piece", quantity: 0, costPrice: 135.0, markup: 40.0, lowStock: 5 },
  { id: "p0373", name: "CP CUP", category: "General", unit: "piece", quantity: 0, costPrice: 7.5, markup: 40.0, lowStock: 5 },
  { id: "p0374", name: "CP ELBOW   1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 330.0, markup: 30.0, lowStock: 5 },
  { id: "p0375", name: "CP MUSLIM SHOWER BODY", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 165.0, markup: 40.0, lowStock: 5 },
  { id: "p0376", name: "CP NIPPLE 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 70.0, markup: 35.0, lowStock: 5 },
  { id: "p0377", name: "CP NIPPLE 1-1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 100.0, markup: 35.0, lowStock: 5 },
  { id: "p0378", name: "CP NIPPLE 1-1/2\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 98.0, markup: 30.0, lowStock: 5 },
  { id: "p0379", name: "CP NIPPLE 12", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 75.0, markup: 30.0, lowStock: 5 },
  { id: "p0380", name: "CP NOZAL 1", category: "General", unit: "piece", quantity: 0, costPrice: 60.0, markup: 40.0, lowStock: 5 },
  { id: "p0381", name: "CP NOZAL 1\"", category: "General", unit: "piece", quantity: 0, costPrice: 98.0, markup: 40.0, lowStock: 5 },
  { id: "p0382", name: "CP NOZAL 1-1/2\"", category: "General", unit: "piece", quantity: 0, costPrice: 155.0, markup: 40.0, lowStock: 5 },
  { id: "p0383", name: "CP NOZAL 2\"", category: "General", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p0384", name: "CP NOZAL AI-58    1", category: "General", unit: "piece", quantity: 0, costPrice: 115.0, markup: 40.0, lowStock: 5 },
  { id: "p0385", name: "CP NOZAL AI-58    1-1/2", category: "General", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p0386", name: "CP NOZAL AI-58    2", category: "General", unit: "piece", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p0387", name: "CP NOZZAL 1", category: "General", unit: "piece", quantity: 0, costPrice: 60.0, markup: 40.0, lowStock: 5 },
  { id: "p0388", name: "CP NOZZAL 1\"", category: "General", unit: "piece", quantity: 0, costPrice: 68.0, markup: 40.0, lowStock: 5 },
  { id: "p0389", name: "CP NOZZAL 1\" BHARI", category: "General", unit: "piece", quantity: 0, costPrice: 87.0, markup: 40.0, lowStock: 5 },
  { id: "p0390", name: "CP NOZZAL 1\" FI", category: "General", unit: "piece", quantity: 0, costPrice: 97.5, markup: 40.0, lowStock: 5 },
  { id: "p0391", name: "CP NOZZAL 1\" IA", category: "General", unit: "piece", quantity: 0, costPrice: 107.0, markup: 40.0, lowStock: 5 },
  { id: "p0392", name: "CP NOZZAL 1-1/2", category: "General", unit: "piece", quantity: 0, costPrice: 75.0, markup: 40.0, lowStock: 5 },
  { id: "p0393", name: "CP NOZZAL 1-1/2 IA", category: "General", unit: "piece", quantity: 0, costPrice: 140.0, markup: 40.0, lowStock: 5 },
  { id: "p0394", name: "CP NOZZAL 1-1/2\"", category: "General", unit: "piece", quantity: 0, costPrice: 92.0, markup: 40.0, lowStock: 5 },
  { id: "p0395", name: "CP NOZZAL 1-1/2\" FI", category: "General", unit: "piece", quantity: 0, costPrice: 115.0, markup: 40.0, lowStock: 5 },
  { id: "p0396", name: "CP NOZZAL 1/2", category: "General", unit: "piece", quantity: 0, costPrice: 85.0, markup: 40.0, lowStock: 5 },
  { id: "p0397", name: "CP NOZZAL 2\"", category: "General", unit: "piece", quantity: 0, costPrice: 115.0, markup: 40.0, lowStock: 5 },
  { id: "p0398", name: "CP NOZZAL 2\" FI", category: "General", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p0399", name: "CP NOZZAL 3\"", category: "General", unit: "piece", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p0400", name: "CP NOZZAL 4\"", category: "General", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p0401", name: "CP NOZZAL 6\"", category: "General", unit: "piece", quantity: 0, costPrice: 248.0, markup: 40.0, lowStock: 5 },
  { id: "p0402", name: "CP NOZZAL I58       1\"", category: "General", unit: "piece", quantity: 0, costPrice: 130.0, markup: 40.0, lowStock: 5 },
  { id: "p0403", name: "CP NOZZAL I58       1-1/2\"", category: "General", unit: "piece", quantity: 0, costPrice: 165.0, markup: 40.0, lowStock: 5 },
  { id: "p0404", name: "CP NOZZAL I58       2\"", category: "General", unit: "piece", quantity: 0, costPrice: 230.0, markup: 40.0, lowStock: 5 },
  { id: "p0405", name: "CP NOZZAL1/2", category: "General", unit: "piece", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p0406", name: "CP SOCKET 1/2", category: "Electrical", unit: "piece", quantity: 0, costPrice: 190.0, markup: 30.0, lowStock: 5 },
  { id: "p0407", name: "CP TEE 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 360.0, markup: 30.0, lowStock: 5 },
  { id: "p0408", name: "CSK 3/4X8", category: "General", unit: "piece", quantity: 0, costPrice: 191.0, markup: 40.0, lowStock: 5 },
  { id: "p0409", name: "CSK SCREW 8X1/2", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 135.0, markup: 40.0, lowStock: 5 },
  { id: "p0410", name: "CUP GRINDING DISC 4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 470.0, markup: 30.0, lowStock: 5 },
  { id: "p0411", name: "CUP GRINDING DISC 5\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 630.0, markup: 30.0, lowStock: 5 },
  { id: "p0412", name: "CYCLE LOCK", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 360.0, markup: 40.0, lowStock: 5 },
  { id: "p0413", name: "CYCLE SCREW 1", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 270.0, markup: 35.0, lowStock: 5 },
  { id: "p0414", name: "CYCLE SCREW 1-1/2", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 380.0, markup: 35.0, lowStock: 5 },
  { id: "p0415", name: "CYCLE SCREW 3/4", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 250.0, markup: 35.0, lowStock: 5 },
  { id: "p0416", name: "CYLENDER LOCK", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 300.0, markup: 30.0, lowStock: 5 },
  { id: "p0417", name: "Cconduit pipe 3/4    29.600 @242", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 0.0, markup: 40.0, lowStock: 5 },
  { id: "p0418", name: "DAATRI", category: "General", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p0419", name: "DABBRA SILVER 4.900KG/KG 420", category: "General", unit: "piece", quantity: 0, costPrice: 353.29, markup: 40.0, lowStock: 5 },
  { id: "p0420", name: "DABRA", category: "General", unit: "piece", quantity: 0, costPrice: 275.0, markup: 40.0, lowStock: 5 },
  { id: "p0421", name: "DABRA 3KG/460", category: "General", unit: "piece", quantity: 0, costPrice: 345.0, markup: 40.0, lowStock: 5 },
  { id: "p0422", name: "DABRA 460/KG", category: "General", unit: "piece", quantity: 0, costPrice: 345.0, markup: 40.0, lowStock: 5 },
  { id: "p0423", name: "DABRA BIG", category: "General", unit: "piece", quantity: 0, costPrice: 337.0, markup: 40.0, lowStock: 5 },
  { id: "p0424", name: "DABRA PLASTIC", category: "General", unit: "piece", quantity: 0, costPrice: 257.5, markup: 40.0, lowStock: 5 },
  { id: "p0425", name: "DABRA SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 310.0, markup: 40.0, lowStock: 5 },
  { id: "p0426", name: "DABRA STEEL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 327.0, markup: 40.0, lowStock: 5 },
  { id: "p0427", name: "DABRA TANGARI", category: "General", unit: "piece", quantity: 0, costPrice: 307.0, markup: 40.0, lowStock: 5 },
  { id: "p0428", name: "DAIGI JALI  6", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 510.0, markup: 30.0, lowStock: 5 },
  { id: "p0429", name: "DAIGI JALI 9X9", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 930.0, markup: 40.0, lowStock: 5 },
  { id: "p0430", name: "DAISI LOCK MUHAFIZ", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 580.0, markup: 40.0, lowStock: 5 },
  { id: "p0431", name: "DASTA BAILCHA", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 175.0, markup: 40.0, lowStock: 5 },
  { id: "p0432", name: "DASTA GAINTI", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 100.0, markup: 40.0, lowStock: 5 },
  { id: "p0433", name: "DATRI", category: "General", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p0434", name: "DB 0-10", category: "Electrical", unit: "piece", quantity: 0, costPrice: 1760.0, markup: 40.0, lowStock: 5 },
  { id: "p0435", name: "DB 0-13", category: "Electrical", unit: "piece", quantity: 0, costPrice: 1600.0, markup: 35.0, lowStock: 5 },
  { id: "p0436", name: "DB 0-15", category: "Electrical", unit: "piece", quantity: 0, costPrice: 1550.0, markup: 30.0, lowStock: 5 },
  { id: "p0437", name: "DB 0-19", category: "Electrical", unit: "piece", quantity: 0, costPrice: 2850.0, markup: 35.0, lowStock: 5 },
  { id: "p0438", name: "DB 0-4", category: "Electrical", unit: "piece", quantity: 0, costPrice: 1100.0, markup: 35.0, lowStock: 5 },
  { id: "p0439", name: "DB 0-4 EICO", category: "Electrical", unit: "piece", quantity: 0, costPrice: 1000.0, markup: 30.0, lowStock: 5 },
  { id: "p0440", name: "DB FALCON PVC 4-6 WAY", category: "Electrical", unit: "piece", quantity: 0, costPrice: 290.0, markup: 30.0, lowStock: 5 },
  { id: "p0441", name: "DB-0-19 EICO", category: "General", unit: "piece", quantity: 0, costPrice: 2850.0, markup: 30.0, lowStock: 5 },
  { id: "p0442", name: "DB-013 EICO", category: "General", unit: "piece", quantity: 0, costPrice: 1650.0, markup: 30.0, lowStock: 5 },
  { id: "p0443", name: "DECOSET SET", category: "General", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p0444", name: "DEEVAR DAABI GM 3/4", category: "General", unit: "piece", quantity: 0, costPrice: 24.0, markup: 40.0, lowStock: 5 },
  { id: "p0445", name: "DEEVAR DABBI 1\"", category: "General", unit: "piece", quantity: 0, costPrice: 23.33, markup: 40.0, lowStock: 5 },
  { id: "p0446", name: "DEEVAR DABBI 3/4", category: "General", unit: "piece", quantity: 0, costPrice: 29.17, markup: 30.0, lowStock: 5 },
  { id: "p0447", name: "DEEVAR DABBI 3/4\"", category: "General", unit: "piece", quantity: 0, costPrice: 21.67, markup: 40.0, lowStock: 5 },
  { id: "p0448", name: "DEEVAR LIGHT", category: "General", unit: "piece", quantity: 0, costPrice: 1050.0, markup: 40.0, lowStock: 5 },
  { id: "p0449", name: "DEEVAR TIKI", category: "General", unit: "packet", quantity: 0, costPrice: 800.0, markup: 35.0, lowStock: 5 },
  { id: "p0450", name: "DEEVAR TIKKI", category: "General", unit: "piece", quantity: 0, costPrice: 12.0, markup: 40.0, lowStock: 5 },
  { id: "p0451", name: "DEEWAR DABBI 1\"", category: "General", unit: "piece", quantity: 0, costPrice: 34.0, markup: 40.0, lowStock: 5 },
  { id: "p0452", name: "DHAGA", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 75.0, markup: 35.0, lowStock: 5 },
  { id: "p0453", name: "DIMMER CHINA FITTING", category: "General", unit: "piece", quantity: 0, costPrice: 125.0, markup: 40.0, lowStock: 5 },
  { id: "p0454", name: "DIMMER PAKISTANI", category: "General", unit: "piece", quantity: 0, costPrice: 105.0, markup: 40.0, lowStock: 5 },
  { id: "p0455", name: "DOOR  LOK", category: "General", unit: "piece", quantity: 0, costPrice: 1500.0, markup: 40.0, lowStock: 5 },
  { id: "p0456", name: "DOOR BELL", category: "General", unit: "piece", quantity: 0, costPrice: 285.0, markup: 35.0, lowStock: 5 },
  { id: "p0457", name: "DOOR BELL DING DONG", category: "General", unit: "piece", quantity: 0, costPrice: 250.0, markup: 35.0, lowStock: 5 },
  { id: "p0458", name: "DOOR BELL SPAROW", category: "General", unit: "piece", quantity: 0, costPrice: 210.0, markup: 40.0, lowStock: 5 },
  { id: "p0459", name: "DOOR BELL SPARROW", category: "General", unit: "piece", quantity: 0, costPrice: 210.0, markup: 30.0, lowStock: 5 },
  { id: "p0460", name: "DOOR CATCHER", category: "General", unit: "piece", quantity: 0, costPrice: 150.0, markup: 30.0, lowStock: 5 },
  { id: "p0461", name: "DOOR CLOSER PASTOL", category: "General", unit: "pair", quantity: 0, costPrice: 710.0, markup: 40.0, lowStock: 5 },
  { id: "p0462", name: "DOOR FARA", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 1300.0, markup: 40.0, lowStock: 5 },
  { id: "p0463", name: "DOOR LOCK", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 1150.0, markup: 40.0, lowStock: 5 },
  { id: "p0464", name: "DOUBLE BELL BUSH", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 145.0, markup: 40.0, lowStock: 5 },
  { id: "p0465", name: "DOUBLE BELL PUSH", category: "General", unit: "piece", quantity: 0, costPrice: 160.0, markup: 30.0, lowStock: 5 },
  { id: "p0466", name: "DOUBLE BIB COCK", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1300.0, markup: 30.0, lowStock: 5 },
  { id: "p0467", name: "DOUBLE BIB COCK SONEX", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1050.0, markup: 40.0, lowStock: 5 },
  { id: "p0468", name: "DOUBLE LIGHT FOR WALL", category: "General", unit: "piece", quantity: 0, costPrice: 1350.0, markup: 40.0, lowStock: 5 },
  { id: "p0469", name: "DOUBLE PUSH BUTTON", category: "General", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p0470", name: "DOUBLE TAPE", category: "General", unit: "piece", quantity: 0, costPrice: 70.0, markup: 30.0, lowStock: 5 },
  { id: "p0471", name: "DRAIN OPENER", category: "General", unit: "piece", quantity: 0, costPrice: 110.0, markup: 35.0, lowStock: 5 },
  { id: "p0472", name: "DRAWER LOCK", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 290.0, markup: 40.0, lowStock: 5 },
  { id: "p0473", name: "DREAM LIGHT BULB 12W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 130.0, markup: 35.0, lowStock: 5 },
  { id: "p0474", name: "DREAM LIGHT BULB 18W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 260.0, markup: 35.0, lowStock: 5 },
  { id: "p0475", name: "DREAM LIGHT BULB 30 W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 410.0, markup: 35.0, lowStock: 5 },
  { id: "p0476", name: "DREAM LIGHT BULB 30W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 410.0, markup: 30.0, lowStock: 5 },
  { id: "p0477", name: "DREAM LIGHT BULB 50W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 750.0, markup: 35.0, lowStock: 5 },
  { id: "p0478", name: "DREAM LIGHT BULB 5W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 105.0, markup: 35.0, lowStock: 5 },
  { id: "p0479", name: "DREAM LIGHT COB LIGHT 5W", category: "General", unit: "piece", quantity: 0, costPrice: 180.0, markup: 35.0, lowStock: 5 },
  { id: "p0480", name: "DREAM LIGHT MOON LIGHT 12W", category: "General", unit: "piece", quantity: 0, costPrice: 310.0, markup: 30.0, lowStock: 5 },
  { id: "p0481", name: "DREAM LIGHT MOON LIGHT 18W", category: "General", unit: "piece", quantity: 0, costPrice: 380.0, markup: 30.0, lowStock: 5 },
  { id: "p0482", name: "DREAM LIGHT ROPE LIGHT", category: "General", unit: "piece", quantity: 0, costPrice: 5600.0, markup: 35.0, lowStock: 5 },
  { id: "p0483", name: "DREAM LIGHT SMD 7W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 155.0, markup: 30.0, lowStock: 5 },
  { id: "p0484", name: "DREAM LIGHT TUBELGHT 4FT 40W", category: "General", unit: "piece", quantity: 0, costPrice: 850.0, markup: 30.0, lowStock: 5 },
  { id: "p0485", name: "DRY ALL SCREW 6X1-1/2", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 149.0, markup: 40.0, lowStock: 5 },
  { id: "p0486", name: "DRY WALL 6X1", category: "General", unit: "piece", quantity: 0, costPrice: 115.0, markup: 40.0, lowStock: 5 },
  { id: "p0487", name: "DRY WALL 6X2", category: "General", unit: "piece", quantity: 0, costPrice: 198.0, markup: 40.0, lowStock: 5 },
  { id: "p0488", name: "DRY WALL 8X1-1/4", category: "General", unit: "piece", quantity: 0, costPrice: 178.0, markup: 40.0, lowStock: 5 },
  { id: "p0489", name: "DRY WALL SACREW   6X1-1/2", category: "General", unit: "piece", quantity: 0, costPrice: 149.0, markup: 40.0, lowStock: 5 },
  { id: "p0490", name: "DRY WALL SACREW   6X2-1/2", category: "General", unit: "piece", quantity: 0, costPrice: 276.0, markup: 40.0, lowStock: 5 },
  { id: "p0491", name: "DRY WALL SACREW   8X2", category: "General", unit: "piece", quantity: 0, costPrice: 298.0, markup: 40.0, lowStock: 5 },
  { id: "p0492", name: "DRY WALL SACREW   8X2-1/2", category: "General", unit: "piece", quantity: 0, costPrice: 365.0, markup: 40.0, lowStock: 5 },
  { id: "p0493", name: "DRY WALL SCREW          2X10", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 398.0, markup: 40.0, lowStock: 5 },
  { id: "p0494", name: "DRY WALL SCREW          3X10", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 590.0, markup: 40.0, lowStock: 5 },
  { id: "p0495", name: "DRY WALL SCREW        1X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 120.0, markup: 35.0, lowStock: 5 },
  { id: "p0496", name: "DRY WALL SCREW        1X8", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 150.0, markup: 35.0, lowStock: 5 },
  { id: "p0497", name: "DRY WALL SCREW     1X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 105.0, markup: 30.0, lowStock: 5 },
  { id: "p0498", name: "DRY WALL SCREW     1X8", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 120.0, markup: 30.0, lowStock: 5 },
  { id: "p0499", name: "DRY WALL SCREW   1-1/2X10", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 342.0, markup: 40.0, lowStock: 5 },
  { id: "p0500", name: "DRY WALL SCREW   1-1/2X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 155.0, markup: 40.0, lowStock: 5 },
  { id: "p0501", name: "DRY WALL SCREW   1-1/2X8", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 198.0, markup: 40.0, lowStock: 5 },
  { id: "p0502", name: "DRY WALL SCREW   1-1/4X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 132.0, markup: 40.0, lowStock: 5 },
  { id: "p0503", name: "DRY WALL SCREW   1/2X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 75.0, markup: 40.0, lowStock: 5 },
  { id: "p0504", name: "DRY WALL SCREW   1X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 104.0, markup: 40.0, lowStock: 5 },
  { id: "p0505", name: "DRY WALL SCREW   3/4X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 96.0, markup: 40.0, lowStock: 5 },
  { id: "p0506", name: "DRY WALL SCREW   3X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 297.0, markup: 40.0, lowStock: 5 },
  { id: "p0507", name: "DRY WALL SCREW  1-1/2X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 155.0, markup: 30.0, lowStock: 5 },
  { id: "p0508", name: "DRY WALL SCREW  1-1/2X8", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 198.0, markup: 30.0, lowStock: 5 },
  { id: "p0509", name: "DRY WALL SCREW  1-1/4X8", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 188.0, markup: 30.0, lowStock: 5 },
  { id: "p0510", name: "DRY WALL SCREW  1X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 118.0, markup: 30.0, lowStock: 5 },
  { id: "p0511", name: "DRY WALL SCREW  2X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 198.0, markup: 40.0, lowStock: 5 },
  { id: "p0512", name: "DRY WALL SCREW  2X8", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 298.0, markup: 40.0, lowStock: 5 },
  { id: "p0513", name: "DRY WALL SCREW  3X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 333.0, markup: 40.0, lowStock: 5 },
  { id: "p0514", name: "DRY WALL SCREW  3X8", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 420.0, markup: 40.0, lowStock: 5 },
  { id: "p0515", name: "DRY WALL SCREW  5/8X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 84.0, markup: 40.0, lowStock: 5 },
  { id: "p0516", name: "DRY WALL SCREW 1-1/2X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 0.0, markup: 40.0, lowStock: 5 },
  { id: "p0517", name: "DRY WALL SCREW 1-1/2X8", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 210.0, markup: 35.0, lowStock: 5 },
  { id: "p0518", name: "DRY WALL SCREW 1-1/4X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 120.0, markup: 30.0, lowStock: 5 },
  { id: "p0519", name: "DRY WALL SCREW 1-1/4X8", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 150.0, markup: 30.0, lowStock: 5 },
  { id: "p0520", name: "DRY WALL SCREW 10X2", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 420.0, markup: 40.0, lowStock: 5 },
  { id: "p0521", name: "DRY WALL SCREW 1X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 118.0, markup: 40.0, lowStock: 5 },
  { id: "p0522", name: "DRY WALL SCREW 1X8", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 152.0, markup: 40.0, lowStock: 5 },
  { id: "p0523", name: "DRY WALL SCREW 2-1/2X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 260.0, markup: 30.0, lowStock: 5 },
  { id: "p0524", name: "DRY WALL SCREW 2-1/2X8", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 306.0, markup: 30.0, lowStock: 5 },
  { id: "p0525", name: "DRY WALL SCREW 2X10", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 398.0, markup: 40.0, lowStock: 5 },
  { id: "p0526", name: "DRY WALL SCREW 2X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 186.0, markup: 30.0, lowStock: 5 },
  { id: "p0527", name: "DRY WALL SCREW 2X8", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 240.0, markup: 30.0, lowStock: 5 },
  { id: "p0528", name: "DRY WALL SCREW 3/4X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 85.0, markup: 30.0, lowStock: 5 },
  { id: "p0529", name: "DRY WALL SCREW 3/4X8", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 105.0, markup: 30.0, lowStock: 5 },
  { id: "p0530", name: "DRY WALL SCREW 5/8X6", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 90.0, markup: 30.0, lowStock: 5 },
  { id: "p0531", name: "DRY WALL SCREW 6X 2", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 198.0, markup: 40.0, lowStock: 5 },
  { id: "p0532", name: "DRY WALL SCREW 6X1", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 104.0, markup: 80.0, lowStock: 5 },
  { id: "p0533", name: "DRY WALL SCREW 6X1-1/2", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 155.0, markup: 40.0, lowStock: 5 },
  { id: "p0534", name: "DRY WALL SCREW 6X1-1/4", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 132.0, markup: 80.0, lowStock: 5 },
  { id: "p0535", name: "DRY WALL SCREW 6X2", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 198.0, markup: 40.0, lowStock: 5 },
  { id: "p0536", name: "DRY WALL SCREW 6X3/4", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 83.0, markup: 80.0, lowStock: 5 },
  { id: "p0537", name: "DRY WALL SCREW 8X1-1/2", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 198.0, markup: 40.0, lowStock: 5 },
  { id: "p0538", name: "DRY WALL SCREW 8X3", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 420.0, markup: 40.0, lowStock: 5 },
  { id: "p0539", name: "DRY WALL SCREW 8X3/4", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 143.0, markup: 40.0, lowStock: 5 },
  { id: "p0540", name: "DRY WALLSCREW 4X10", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 720.0, markup: 30.0, lowStock: 5 },
  { id: "p0541", name: "DUCT 40X40", category: "General", unit: "piece", quantity: 0, costPrice: 750.0, markup: 35.0, lowStock: 5 },
  { id: "p0542", name: "DUCT MEDIUM", category: "General", unit: "piece", quantity: 0, costPrice: 235.0, markup: 35.0, lowStock: 5 },
  { id: "p0543", name: "DUCT PATTI  1 \" 15pcs paking", category: "General", unit: "piece", quantity: 0, costPrice: 235.0, markup: 30.0, lowStock: 5 },
  { id: "p0544", name: "DUCT PATTI 1/2", category: "General", unit: "piece", quantity: 0, costPrice: 110.0, markup: 35.0, lowStock: 5 },
  { id: "p0545", name: "DUCT PATTI 1/2\" 40pcs paking", category: "General", unit: "piece", quantity: 0, costPrice: 125.0, markup: 30.0, lowStock: 5 },
  { id: "p0546", name: "DUCT PATTI 3/4", category: "General", unit: "piece", quantity: 0, costPrice: 215.0, markup: 40.0, lowStock: 5 },
  { id: "p0547", name: "DUCT PATTI 3/4\"", category: "General", unit: "piece", quantity: 0, costPrice: 220.0, markup: 35.0, lowStock: 5 },
  { id: "p0548", name: "DUCT PATTI 3/4\" 25pcs paking", category: "General", unit: "piece", quantity: 0, costPrice: 185.0, markup: 30.0, lowStock: 5 },
  { id: "p0549", name: "DUCT PATTI 40X40", category: "General", unit: "piece", quantity: 0, costPrice: 600.0, markup: 40.0, lowStock: 5 },
  { id: "p0550", name: "DUCT PATTI CLIP", category: "General", unit: "piece", quantity: 0, costPrice: 1.35, markup: 40.0, lowStock: 5 },
  { id: "p0551", name: "DUCT PATTI NAIL", category: "Hardware & Tools", unit: "packet", quantity: 0, costPrice: 165.0, markup: 35.0, lowStock: 5 },
  { id: "p0552", name: "DUCT PUTTY 1\"", category: "Paint", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p0553", name: "DUCT PUTTY 3/4\"", category: "Paint", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p0554", name: "EARTH WIRE WITH COULE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 380.0, markup: 40.0, lowStock: 5 },
  { id: "p0555", name: "ELBI BABY", category: "General", unit: "piece", quantity: 0, costPrice: 17.0, markup: 40.0, lowStock: 5 },
  { id: "p0556", name: "ELBOW     3X45", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 117.0, markup: 40.0, lowStock: 5 },
  { id: "p0557", name: "ELBOW 2\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 58.0, markup: 40.0, lowStock: 5 },
  { id: "p0558", name: "ELBOW 3\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 175.9, markup: 40.0, lowStock: 5 },
  { id: "p0559", name: "ELBOW 3X45", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 152.4, markup: 40.0, lowStock: 5 },
  { id: "p0560", name: "ELECTRIC PLIER CUTTER", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 680.0, markup: 30.0, lowStock: 5 },
  { id: "p0561", name: "ELFI", category: "Paint", unit: "piece", quantity: 0, costPrice: 17.0, markup: 40.0, lowStock: 5 },
  { id: "p0562", name: "ELFI 20 GRM", category: "Paint", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p0563", name: "ELFI 20GRM", category: "Paint", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p0564", name: "ELFI 2O GRM", category: "Paint", unit: "piece", quantity: 0, costPrice: 105.0, markup: 40.0, lowStock: 5 },
  { id: "p0565", name: "ELFI 50GRM", category: "Paint", unit: "piece", quantity: 0, costPrice: 246.67, markup: 40.0, lowStock: 5 },
  { id: "p0566", name: "ELFI BIG", category: "Paint", unit: "piece", quantity: 0, costPrice: 260.0, markup: 40.0, lowStock: 5 },
  { id: "p0567", name: "ELFI MEDIUM", category: "Paint", unit: "piece", quantity: 0, costPrice: 170.0, markup: 40.0, lowStock: 5 },
  { id: "p0568", name: "ELFI WARNISH", category: "Paint", unit: "piece", quantity: 0, costPrice: 550.0, markup: 30.0, lowStock: 5 },
  { id: "p0569", name: "EMULSION 1/4", category: "General", unit: "piece", quantity: 0, costPrice: 350.0, markup: 40.0, lowStock: 5 },
  { id: "p0570", name: "ENAMEL OIL PAINT  QRT", category: "Paint", unit: "piece", quantity: 0, costPrice: 495.0, markup: 35.0, lowStock: 5 },
  { id: "p0571", name: "ENAMEL PAINT DAYAR COLOUR", category: "Paint", unit: "gal", quantity: 0, costPrice: 1550.0, markup: 30.0, lowStock: 5 },
  { id: "p0572", name: "ENAMEL PAINT QUARTER", category: "Paint", unit: "piece", quantity: 0, costPrice: 490.0, markup: 30.0, lowStock: 5 },
  { id: "p0573", name: "ENAMEL PAINT SMALL", category: "Paint", unit: "piece", quantity: 0, costPrice: 165.0, markup: 30.0, lowStock: 5 },
  { id: "p0574", name: "ENAMEL PAINT SMALL SILVER", category: "Paint", unit: "piece", quantity: 0, costPrice: 175.0, markup: 30.0, lowStock: 5 },
  { id: "p0575", name: "END CAP 3\"", category: "General", unit: "piece", quantity: 0, costPrice: 98.0, markup: 40.0, lowStock: 5 },
  { id: "p0576", name: "END CAP 4\"", category: "General", unit: "piece", quantity: 0, costPrice: 113.0, markup: 40.0, lowStock: 5 },
  { id: "p0577", name: "ENEMAL PAINT QRTR", category: "Paint", unit: "piece", quantity: 0, costPrice: 490.0, markup: 35.0, lowStock: 5 },
  { id: "p0578", name: "ENEMEL PAINT", category: "Paint", unit: "490", quantity: 0, costPrice: 490.0, markup: 35.0, lowStock: 5 },
  { id: "p0579", name: "ENGLIS SEAT COMMOD", category: "General", unit: "piece", quantity: 0, costPrice: 700.0, markup: 30.0, lowStock: 5 },
  { id: "p0580", name: "ENGLISH WC WHITE", category: "General", unit: "piece", quantity: 0, costPrice: 7500.0, markup: 40.0, lowStock: 5 },
  { id: "p0581", name: "EXCEL", category: "General", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p0582", name: "EXHAUST FAN 10", category: "General", unit: "piece", quantity: 0, costPrice: 3150.0, markup: 30.0, lowStock: 5 },
  { id: "p0583", name: "EXHUAST FAN 10 PVC", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 2800.0, markup: 40.0, lowStock: 5 },
  { id: "p0584", name: "EXHUAST FAN 12 METAL REAL", category: "General", unit: "piece", quantity: 0, costPrice: 2650.0, markup: 40.0, lowStock: 5 },
  { id: "p0585", name: "EXHUST FAN 2\"", category: "General", unit: "piece", quantity: 0, costPrice: 1950.0, markup: 40.0, lowStock: 5 },
  { id: "p0586", name: "EXHUST FAN 8\"", category: "General", unit: "piece", quantity: 0, costPrice: 2050.0, markup: 40.0, lowStock: 5 },
  { id: "p0587", name: "EXTENSION BOARD MT", category: "General", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p0588", name: "EXTENSION BOARD SA 012", category: "General", unit: "piece", quantity: 0, costPrice: 170.0, markup: 40.0, lowStock: 5 },
  { id: "p0589", name: "EXTENSION BOARD SA 212", category: "General", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p0590", name: "EXTENSION CABLE  WHITE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 360.0, markup: 35.0, lowStock: 5 },
  { id: "p0591", name: "EXTENSION CABLE ORANGE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 240.0, markup: 35.0, lowStock: 5 },
  { id: "p0592", name: "EXTENSION CABLE ROUND", category: "Electrical", unit: "piece", quantity: 0, costPrice: 330.0, markup: 35.0, lowStock: 5 },
  { id: "p0593", name: "EXTENSION LEAD GREEN COLOUR", category: "Paint", unit: "piece", quantity: 0, costPrice: 380.0, markup: 40.0, lowStock: 5 },
  { id: "p0594", name: "EXTENSION LEAD JAHAZ BLUE COLOUR", category: "Paint", unit: "piece", quantity: 0, costPrice: 360.0, markup: 40.0, lowStock: 5 },
  { id: "p0595", name: "EXTENSION LEAD ORANGE COLOUR", category: "Paint", unit: "piece", quantity: 0, costPrice: 410.0, markup: 40.0, lowStock: 5 },
  { id: "p0596", name: "EXTENSION LEAD PLATE", category: "General", unit: "piece", quantity: 0, costPrice: 310.0, markup: 40.0, lowStock: 5 },
  { id: "p0597", name: "FAJAR LOCK", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 520.0, markup: 30.0, lowStock: 5 },
  { id: "p0598", name: "FAN BOX", category: "General", unit: "145", quantity: 0, costPrice: 136.3, markup: 30.0, lowStock: 5 },
  { id: "p0599", name: "FAN BOX GM", category: "General", unit: "piece", quantity: 0, costPrice: 185.0, markup: 40.0, lowStock: 5 },
  { id: "p0600", name: "FAN BOX GRACE", category: "General", unit: "piece", quantity: 0, costPrice: 125.0, markup: 30.0, lowStock: 5 },
  { id: "p0601", name: "FAN CAPACITOR 3.5", category: "Electrical", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p0602", name: "FAN DIMMER AMSON", category: "General", unit: "piece", quantity: 0, costPrice: 145.0, markup: 35.0, lowStock: 5 },
  { id: "p0603", name: "FAN DIMMERCHINA", category: "General", unit: "piece", quantity: 0, costPrice: 160.0, markup: 35.0, lowStock: 5 },
  { id: "p0604", name: "FAN HOLDER ROUND", category: "Electrical", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p0605", name: "FAN HOOCK RAWAL BOLT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 57.5, markup: 30.0, lowStock: 5 },
  { id: "p0606", name: "FAN HOOCK ROUND", category: "General", unit: "piece", quantity: 0, costPrice: 170.0, markup: 30.0, lowStock: 5 },
  { id: "p0607", name: "FAN ROD 1-1/2FT", category: "General", unit: "piece", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p0608", name: "FAN ROD 2 FT", category: "General", unit: "piece", quantity: 0, costPrice: 230.0, markup: 40.0, lowStock: 5 },
  { id: "p0609", name: "FAN SHEET", category: "General", unit: "piece", quantity: 0, costPrice: 44.17, markup: 35.0, lowStock: 5 },
  { id: "p0610", name: "FANOOS CHAIN", category: "General", unit: "piece", quantity: 0, costPrice: 150.0, markup: 30.0, lowStock: 5 },
  { id: "p0611", name: "FANOOS CHAN ORDINORY", category: "General", unit: "ft", quantity: 0, costPrice: 35.5, markup: 30.0, lowStock: 5 },
  { id: "p0612", name: "FARA LOCK", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 1250.0, markup: 35.0, lowStock: 5 },
  { id: "p0613", name: "FINE BASIN MIXER PVC", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 720.0, markup: 30.0, lowStock: 5 },
  { id: "p0614", name: "FISH DOORI", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 45.0, markup: 35.0, lowStock: 5 },
  { id: "p0615", name: "FLEXBLE PIPE    1\" ROLL 300FT", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 6.17, markup: 30.0, lowStock: 5 },
  { id: "p0616", name: "FLEXBLE PIPE 3/4\" ROLL 300FT", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 5.5, markup: 30.0, lowStock: 5 },
  { id: "p0617", name: "FLEXIBLE PIPE 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1850.0, markup: 35.0, lowStock: 5 },
  { id: "p0618", name: "FLEXIBLE PIPE 1-1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 2225.0, markup: 40.0, lowStock: 5 },
  { id: "p0619", name: "FLEXIBLE PIPE 2\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 2100.0, markup: 40.0, lowStock: 5 },
  { id: "p0620", name: "FLEXIBLE PIPE 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1650.0, markup: 35.0, lowStock: 5 },
  { id: "p0621", name: "FLOAT VALVE  3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 640.0, markup: 30.0, lowStock: 5 },
  { id: "p0622", name: "FLOAT VALVE /FLUSH TAN FITTING", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 330.0, markup: 30.0, lowStock: 5 },
  { id: "p0623", name: "FLOOD LIGHT 50W", category: "General", unit: "piece", quantity: 0, costPrice: 580.0, markup: 30.0, lowStock: 5 },
  { id: "p0624", name: "FLOOD LIGHT CHINA", category: "General", unit: "piece", quantity: 0, costPrice: 260.0, markup: 30.0, lowStock: 5 },
  { id: "p0625", name: "FLOOR JALI 4X4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 190.0, markup: 35.0, lowStock: 5 },
  { id: "p0626", name: "FLOOR MESH 4x4", category: "General", unit: "piece", quantity: 0, costPrice: 270.0, markup: 40.0, lowStock: 5 },
  { id: "p0627", name: "FLOOR MESH 6x6", category: "General", unit: "piece", quantity: 0, costPrice: 380.0, markup: 40.0, lowStock: 5 },
  { id: "p0628", name: "FLUSH TANK 3STAR", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 3000.0, markup: 40.0, lowStock: 5 },
  { id: "p0629", name: "FLUSH TANK SYPHAN", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 480.0, markup: 40.0, lowStock: 5 },
  { id: "p0630", name: "FLUSH TANK SYPHAN ABS", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 580.0, markup: 40.0, lowStock: 5 },
  { id: "p0631", name: "FLUSH TANKI", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1600.0, markup: 40.0, lowStock: 5 },
  { id: "p0632", name: "FOAM BIG", category: "General", unit: "piece", quantity: 0, costPrice: 47.22, markup: 30.0, lowStock: 5 },
  { id: "p0633", name: "FOAM SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 33.33, markup: 30.0, lowStock: 5 },
  { id: "p0634", name: "FOAM SPRAY", category: "General", unit: "piece", quantity: 0, costPrice: 625.0, markup: 30.0, lowStock: 5 },
  { id: "p0635", name: "FOLDING NECK  SINK LONG HEAD", category: "General", unit: "piece", quantity: 0, costPrice: 370.0, markup: 40.0, lowStock: 5 },
  { id: "p0636", name: "GAINER LOCK  30MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 135.0, markup: 30.0, lowStock: 5 },
  { id: "p0637", name: "GAINER LOCK  40MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 250.0, markup: 30.0, lowStock: 5 },
  { id: "p0638", name: "GAINER LOCK  50MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 290.0, markup: 30.0, lowStock: 5 },
  { id: "p0639", name: "GANTI DASTA HANDLE PICK AX HANDLE", category: "General", unit: "piece", quantity: 0, costPrice: 120.0, markup: 30.0, lowStock: 5 },
  { id: "p0640", name: "GARDEN PIPE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 21.28, markup: 40.0, lowStock: 5 },
  { id: "p0641", name: "GARDEN PIPE 16KG X610=9760", category: "Plumbing & Sanitary", unit: "ft", quantity: 0, costPrice: 35.49, markup: 30.0, lowStock: 5 },
  { id: "p0642", name: "GARDEN PIPE 3/4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 550.0, markup: 40.0, lowStock: 5 },
  { id: "p0643", name: "GASS BALL VALVE 1/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p0644", name: "GASS NOZAL", category: "General", unit: "piece", quantity: 0, costPrice: 20.0, markup: 30.0, lowStock: 5 },
  { id: "p0645", name: "GASS NOZAL;", category: "General", unit: "piece", quantity: 0, costPrice: 18.0, markup: 40.0, lowStock: 5 },
  { id: "p0646", name: "GASS TEE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 10.0, markup: 40.0, lowStock: 5 },
  { id: "p0647", name: "GATE VALVE 25MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 285.0, markup: 40.0, lowStock: 5 },
  { id: "p0648", name: "GATE VALVE 32MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 380.0, markup: 35.0, lowStock: 5 },
  { id: "p0649", name: "GATTA", category: "General", unit: "piece", quantity: 0, costPrice: 220.0, markup: 30.0, lowStock: 5 },
  { id: "p0650", name: "GI BUSH 1-1/2X3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 148.0, markup: 40.0, lowStock: 5 },
  { id: "p0651", name: "GI BUSH 1-1/4X1", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 90.0, markup: 35.0, lowStock: 5 },
  { id: "p0652", name: "GI BUSH 1-1/4X3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p0653", name: "GI BUSH 1X3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 99.0, markup: 30.0, lowStock: 5 },
  { id: "p0654", name: "GI BUSH 3/4X1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 48.33, markup: 35.0, lowStock: 5 },
  { id: "p0655", name: "GI BUSH NORMAL 1-1/4X1", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 90.0, markup: 30.0, lowStock: 5 },
  { id: "p0656", name: "GI BUSH NORMAL 1X1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 50.0, markup: 30.0, lowStock: 5 },
  { id: "p0657", name: "GI ELBOW 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 141.0, markup: 40.0, lowStock: 5 },
  { id: "p0658", name: "GI ELBOW 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 60.0, markup: 40.0, lowStock: 5 },
  { id: "p0659", name: "GI ELBOW 1/2\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 46.0, markup: 40.0, lowStock: 5 },
  { id: "p0660", name: "GI ELBOW 1X3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 207.0, markup: 30.0, lowStock: 5 },
  { id: "p0661", name: "GI ELBOW 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 95.0, markup: 30.0, lowStock: 5 },
  { id: "p0662", name: "GI ELBOW CHINA 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 130.0, markup: 30.0, lowStock: 5 },
  { id: "p0663", name: "GI FM ELBOW", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 100.0, markup: 40.0, lowStock: 5 },
  { id: "p0664", name: "GI FM ELBOW 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 113.33, markup: 40.0, lowStock: 5 },
  { id: "p0665", name: "GI FM ELBOW 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 135.0, markup: 30.0, lowStock: 5 },
  { id: "p0666", name: "GI HEX NIPPLE 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 65.0, markup: 30.0, lowStock: 5 },
  { id: "p0667", name: "GI HEX NIPPLE 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 92.0, markup: 30.0, lowStock: 5 },
  { id: "p0668", name: "GI M FEMALE ELBOW 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 148.33, markup: 40.0, lowStock: 5 },
  { id: "p0669", name: "GI M/FEMALE ELBOW 1/2\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 72.0, markup: 40.0, lowStock: 5 },
  { id: "p0670", name: "GI M/FEMALE ELBOW 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 0.0, markup: 40.0, lowStock: 5 },
  { id: "p0671", name: "GI NAIL WASHER", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 420.0, markup: 40.0, lowStock: 5 },
  { id: "p0672", name: "GI NIPPLE 1", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 123.0, markup: 40.0, lowStock: 5 },
  { id: "p0673", name: "GI NIPPLE 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 140.0, markup: 35.0, lowStock: 5 },
  { id: "p0674", name: "GI NIPPLE 1\" CHINA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 125.0, markup: 30.0, lowStock: 5 },
  { id: "p0675", name: "GI NIPPLE 1-1/4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 265.0, markup: 35.0, lowStock: 5 },
  { id: "p0676", name: "GI NIPPLE 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 65.0, markup: 35.0, lowStock: 5 },
  { id: "p0677", name: "GI NIPPLE 1/2\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 38.0, markup: 40.0, lowStock: 5 },
  { id: "p0678", name: "GI NIPPLE 1/2\" ORIGNAL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 66.67, markup: 40.0, lowStock: 5 },
  { id: "p0679", name: "GI NIPPLE 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 82.5, markup: 35.0, lowStock: 5 },
  { id: "p0680", name: "GI NIPPLE 3/4 CHINA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 76.67, markup: 30.0, lowStock: 5 },
  { id: "p0681", name: "GI NIPPLE 3/4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p0682", name: "GI NIPPLE 3/4\" ORIGNAL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 100.0, markup: 40.0, lowStock: 5 },
  { id: "p0683", name: "GI NIPPLE CHAINA 1-1/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 190.83, markup: 30.0, lowStock: 5 },
  { id: "p0684", name: "GI NIPPLE CHINA 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 52.5, markup: 30.0, lowStock: 5 },
  { id: "p0685", name: "GI NIPPLE CHINA 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 77.5, markup: 30.0, lowStock: 5 },
  { id: "p0686", name: "GI NIPPLE HE 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p0687", name: "GI NIPPLE HE 1/2\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p0688", name: "GI NIPPLE HE 3/4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 100.0, markup: 40.0, lowStock: 5 },
  { id: "p0689", name: "GI PIPE NIPPLE 1-1/4X12", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 295.0, markup: 30.0, lowStock: 5 },
  { id: "p0690", name: "GI PIPE NIPPLE 1/2X4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 38.0, markup: 30.0, lowStock: 5 },
  { id: "p0691", name: "GI PIPE NIPPLE 3/4X6\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 75.0, markup: 30.0, lowStock: 5 },
  { id: "p0692", name: "GI PIPIE NIPPLE 1/2X12\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 110.0, markup: 30.0, lowStock: 5 },
  { id: "p0693", name: "GI REDUCER SOCKET 1X3/4", category: "Electrical", unit: "piece", quantity: 0, costPrice: 113.0, markup: 40.0, lowStock: 5 },
  { id: "p0694", name: "GI REDUCER SOCKET 3/4X1/2", category: "Electrical", unit: "piece", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p0695", name: "GI SOCKET  1/2", category: "Electrical", unit: "piece", quantity: 0, costPrice: 55.0, markup: 40.0, lowStock: 5 },
  { id: "p0696", name: "GI SOCKET 1", category: "Electrical", unit: "piece", quantity: 0, costPrice: 110.0, markup: 40.0, lowStock: 5 },
  { id: "p0697", name: "GI SOCKET 1\" PK C", category: "Electrical", unit: "piece", quantity: 0, costPrice: 140.0, markup: 40.0, lowStock: 5 },
  { id: "p0698", name: "GI SOCKET 1/2\"", category: "Electrical", unit: "piece", quantity: 0, costPrice: 70.0, markup: 35.0, lowStock: 5 },
  { id: "p0699", name: "GI SOCKET 1X1/2", category: "Electrical", unit: "piece", quantity: 0, costPrice: 160.0, markup: 30.0, lowStock: 5 },
  { id: "p0700", name: "GI SOCKET 1X1/2 CHINA", category: "Electrical", unit: "piece", quantity: 0, costPrice: 101.17, markup: 30.0, lowStock: 5 },
  { id: "p0701", name: "GI SOCKET 3/4", category: "Electrical", unit: "piece", quantity: 0, costPrice: 88.0, markup: 30.0, lowStock: 5 },
  { id: "p0702", name: "GI SOCKET 3/4\"", category: "Electrical", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p0703", name: "GI SOCKET 3/4. CHINA", category: "Electrical", unit: "piece", quantity: 0, costPrice: 85.0, markup: 30.0, lowStock: 5 },
  { id: "p0704", name: "GI SOCKET 3/4X1/2", category: "Electrical", unit: "piece", quantity: 0, costPrice: 131.67, markup: 30.0, lowStock: 5 },
  { id: "p0705", name: "GI SOCKET CHINA 1-1/4", category: "Electrical", unit: "piece", quantity: 0, costPrice: 206.0, markup: 30.0, lowStock: 5 },
  { id: "p0706", name: "GI SOCKET CHINA 1-1/4X1", category: "Electrical", unit: "piece", quantity: 0, costPrice: 250.0, markup: 30.0, lowStock: 5 },
  { id: "p0707", name: "GI SOCKET CHINA 1/2", category: "Electrical", unit: "piece", quantity: 0, costPrice: 60.0, markup: 30.0, lowStock: 5 },
  { id: "p0708", name: "GI TEE 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 230.0, markup: 40.0, lowStock: 5 },
  { id: "p0709", name: "GI TEE 1-1/4.", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 390.0, markup: 30.0, lowStock: 5 },
  { id: "p0710", name: "GI TEE 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 105.0, markup: 35.0, lowStock: 5 },
  { id: "p0711", name: "GI TEE 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 165.0, markup: 35.0, lowStock: 5 },
  { id: "p0712", name: "GI TEE 3/4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 127.0, markup: 40.0, lowStock: 5 },
  { id: "p0713", name: "GI TEE CHINA 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 85.0, markup: 30.0, lowStock: 5 },
  { id: "p0714", name: "GI THREADED ROD 10MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 420.0, markup: 40.0, lowStock: 5 },
  { id: "p0715", name: "GI U CLAMP", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 20.83, markup: 30.0, lowStock: 5 },
  { id: "p0716", name: "GI U CLAMP   2\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 13.75, markup: 40.0, lowStock: 5 },
  { id: "p0717", name: "GI U CLAMP  1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 3.75, markup: 40.0, lowStock: 5 },
  { id: "p0718", name: "GI U CLAMP  2\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 190.0, markup: 40.0, lowStock: 5 },
  { id: "p0719", name: "GI U CLAMP  3\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 15.83, markup: 40.0, lowStock: 5 },
  { id: "p0720", name: "GI U CLAMP  4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 16.67, markup: 40.0, lowStock: 5 },
  { id: "p0721", name: "GI U CLAMP 1", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 5.0, markup: 40.0, lowStock: 5 },
  { id: "p0722", name: "GI U CLAMP 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 12.0, markup: 40.0, lowStock: 5 },
  { id: "p0723", name: "GI U CLAMP 1-1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 10.83, markup: 40.0, lowStock: 5 },
  { id: "p0724", name: "GI U CLAMP 1-1/2\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 11.67, markup: 40.0, lowStock: 5 },
  { id: "p0725", name: "GI U CLAMP 2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 16.23, markup: 40.0, lowStock: 5 },
  { id: "p0726", name: "GI U CLAMP 2\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 15.83, markup: 40.0, lowStock: 5 },
  { id: "p0727", name: "GI U CLAMP 3", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 17.08, markup: 40.0, lowStock: 5 },
  { id: "p0728", name: "GI U CLAMP 3\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 19.33, markup: 30.0, lowStock: 5 },
  { id: "p0729", name: "GI U CLAMP 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 4.17, markup: 40.0, lowStock: 5 },
  { id: "p0730", name: "GI U CLAMP 4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 19.17, markup: 40.0, lowStock: 5 },
  { id: "p0731", name: "GI UNION", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 350.0, markup: 30.0, lowStock: 5 },
  { id: "p0732", name: "GI UNION 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 250.0, markup: 40.0, lowStock: 5 },
  { id: "p0733", name: "GI UNION 1-1/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 680.0, markup: 30.0, lowStock: 5 },
  { id: "p0734", name: "GI UNION 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 335.0, markup: 35.0, lowStock: 5 },
  { id: "p0735", name: "GI UNION 3/4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 308.0, markup: 30.0, lowStock: 5 },
  { id: "p0736", name: "GI UNION 384", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 275.0, markup: 35.0, lowStock: 5 },
  { id: "p0737", name: "GI UNION CHINA 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 340.0, markup: 40.0, lowStock: 5 },
  { id: "p0738", name: "GI UNION PAK CHINA 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 248.0, markup: 40.0, lowStock: 5 },
  { id: "p0739", name: "GI WASHER", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 430.0, markup: 30.0, lowStock: 5 },
  { id: "p0740", name: "GI WASHER FOR RAWAL BOLT", category: "Plumbing & Sanitary", unit: "kg", quantity: 0, costPrice: 340.0, markup: 40.0, lowStock: 5 },
  { id: "p0741", name: "GINIPPLE 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 96.0, markup: 30.0, lowStock: 5 },
  { id: "p0742", name: "GLASSI", category: "General", unit: "3360", quantity: 0, costPrice: 33.0, markup: 30.0, lowStock: 5 },
  { id: "p0743", name: "GLASSI GM", category: "General", unit: "piece", quantity: 0, costPrice: 39.0, markup: 40.0, lowStock: 5 },
  { id: "p0744", name: "GLASSI GRACE", category: "General", unit: "piece", quantity: 0, costPrice: 27.69, markup: 30.0, lowStock: 5 },
  { id: "p0745", name: "GLASSY", category: "General", unit: "piece", quantity: 0, costPrice: 39.06, markup: 40.0, lowStock: 5 },
  { id: "p0746", name: "GLOBE", category: "General", unit: "piece", quantity: 0, costPrice: 650.0, markup: 40.0, lowStock: 5 },
  { id: "p0747", name: "GLOVES 300", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 105.0, markup: 40.0, lowStock: 5 },
  { id: "p0748", name: "GLOVES B#300", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 80.0, markup: 35.0, lowStock: 5 },
  { id: "p0749", name: "GLOVES BADBU", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p0750", name: "GLOVES BLUE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 175.0, markup: 35.0, lowStock: 5 },
  { id: "p0751", name: "GLOVES BUDBO DASTANA", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 95.0, markup: 40.0, lowStock: 5 },
  { id: "p0752", name: "GLOVES LONG PERPAL", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 100.0, markup: 40.0, lowStock: 5 },
  { id: "p0753", name: "GM CABLE 10MM", category: "Electrical", unit: "roll", quantity: 0, costPrice: 34800.0, markup: 40.0, lowStock: 5 },
  { id: "p0754", name: "GO RASSI", category: "General", unit: "piece", quantity: 0, costPrice: 27.59, markup: 40.0, lowStock: 5 },
  { id: "p0755", name: "GO RASSI / PCS 25", category: "General", unit: "kg", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p0756", name: "GO RASSI 65/PCS 1/2KG EACH", category: "General", unit: "piece", quantity: 0, costPrice: 135.0, markup: 40.0, lowStock: 5 },
  { id: "p0757", name: "GODA JORI", category: "General", unit: "piece", quantity: 0, costPrice: 300.0, markup: 40.0, lowStock: 5 },
  { id: "p0758", name: "GOLA SOOTAR", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 29.17, markup: 30.0, lowStock: 5 },
  { id: "p0759", name: "GOLDEN COLOUR", category: "Paint", unit: "piece", quantity: 0, costPrice: 210.0, markup: 40.0, lowStock: 5 },
  { id: "p0760", name: "GOVES GULABI", category: "General", unit: "piece", quantity: 0, costPrice: 95.0, markup: 40.0, lowStock: 5 },
  { id: "p0761", name: "GRECE", category: "General", unit: "piece", quantity: 0, costPrice: 75.0, markup: 40.0, lowStock: 5 },
  { id: "p0762", name: "GREES", category: "General", unit: "piece", quantity: 0, costPrice: 70.0, markup: 30.0, lowStock: 5 },
  { id: "p0763", name: "GRINDIND DISC 5\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 130.0, markup: 40.0, lowStock: 5 },
  { id: "p0764", name: "GRINDING CUP 5\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 590.0, markup: 40.0, lowStock: 5 },
  { id: "p0765", name: "GRINDING NUT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p0766", name: "GUNYA", category: "General", unit: "piece", quantity: 0, costPrice: 195.0, markup: 40.0, lowStock: 5 },
  { id: "p0767", name: "GUNYA SMALL BLACK", category: "General", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p0768", name: "GUNYA SMALL WHITE", category: "General", unit: "piece", quantity: 0, costPrice: 350.0, markup: 40.0, lowStock: 5 },
  { id: "p0769", name: "GURMALA", category: "General", unit: "piece", quantity: 0, costPrice: 80.0, markup: 30.0, lowStock: 5 },
  { id: "p0770", name: "GURMALA SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 70.0, markup: 30.0, lowStock: 5 },
  { id: "p0771", name: "GYZER PIIPE 1-1/2FT", category: "General", unit: "piece", quantity: 0, costPrice: 115.0, markup: 40.0, lowStock: 5 },
  { id: "p0772", name: "GYZER PIIPE 1MTR", category: "General", unit: "piece", quantity: 0, costPrice: 248.0, markup: 40.0, lowStock: 5 },
  { id: "p0773", name: "GYZER PIIPE 2FT", category: "General", unit: "piece", quantity: 0, costPrice: 158.0, markup: 40.0, lowStock: 5 },
  { id: "p0774", name: "GYZER PIIPE 3FT", category: "General", unit: "piece", quantity: 0, costPrice: 325.0, markup: 35.0, lowStock: 5 },
  { id: "p0775", name: "GYZER PIPE 11-1/2FT", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 200.0, markup: 35.0, lowStock: 5 },
  { id: "p0776", name: "GYZER PIPE 12\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 130.0, markup: 40.0, lowStock: 5 },
  { id: "p0777", name: "GYZER PIPE 18\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 150.0, markup: 40.0, lowStock: 5 },
  { id: "p0778", name: "GYZER PIPE 1FT", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 190.0, markup: 35.0, lowStock: 5 },
  { id: "p0779", name: "GYZER PIPE 1MTR", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 260.0, markup: 40.0, lowStock: 5 },
  { id: "p0780", name: "GYZER PIPE 24\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 165.0, markup: 40.0, lowStock: 5 },
  { id: "p0781", name: "GYZER PIPE 2FT", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 215.0, markup: 35.0, lowStock: 5 },
  { id: "p0782", name: "HAMMER 1 KG", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 310.0, markup: 40.0, lowStock: 5 },
  { id: "p0783", name: "HAMMER 1-1/2KG", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 410.0, markup: 40.0, lowStock: 5 },
  { id: "p0784", name: "HAMMER 1.5KG", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 450.0, markup: 40.0, lowStock: 5 },
  { id: "p0785", name: "HAMMER 1KG", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 350.0, markup: 40.0, lowStock: 5 },
  { id: "p0786", name: "HAMMER HANDLE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 37.5, markup: 35.0, lowStock: 5 },
  { id: "p0787", name: "HAMMER WOODEN HATHORI", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 176.67, markup: 40.0, lowStock: 5 },
  { id: "p0788", name: "HAND DRILL BIT", category: "General", unit: "piece", quantity: 0, costPrice: 280.0, markup: 30.0, lowStock: 5 },
  { id: "p0789", name: "HAND GLOVES", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 96.67, markup: 40.0, lowStock: 5 },
  { id: "p0790", name: "HAND GLOVES 300", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 105.0, markup: 40.0, lowStock: 5 },
  { id: "p0791", name: "HAND GLOVES BADBU", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 95.83, markup: 35.0, lowStock: 5 },
  { id: "p0792", name: "HAND GLOVES LONG", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 95.83, markup: 40.0, lowStock: 5 },
  { id: "p0793", name: "HAND GLOVES LONG PINK", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 87.5, markup: 30.0, lowStock: 5 },
  { id: "p0794", name: "HAND GLOVES LONG YELLOW", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 150.0, markup: 30.0, lowStock: 5 },
  { id: "p0795", name: "HAND GLOVES PERPAL", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 93.33, markup: 35.0, lowStock: 5 },
  { id: "p0796", name: "HAND LEVEL PRIDE 12", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 740.0, markup: 30.0, lowStock: 5 },
  { id: "p0797", name: "HAND LEVEL YELLOE 12\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 200.0, markup: 35.0, lowStock: 5 },
  { id: "p0798", name: "HAND LEVEL ZEROX", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 420.0, markup: 40.0, lowStock: 5 },
  { id: "p0799", name: "HAND SAW 14\" ORANGE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 303.0, markup: 40.0, lowStock: 5 },
  { id: "p0800", name: "HAND SAW 16\" ORANGE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 270.0, markup: 40.0, lowStock: 5 },
  { id: "p0801", name: "HAND SAW 18\" ORANGE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 360.0, markup: 40.0, lowStock: 5 },
  { id: "p0802", name: "HAND SAW AARI MALTA", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 270.0, markup: 30.0, lowStock: 5 },
  { id: "p0803", name: "HAND SAW ORANGE SMALL", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 210.0, markup: 40.0, lowStock: 5 },
  { id: "p0804", name: "HAND SAW WOODEN HANDLE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 150.0, markup: 40.0, lowStock: 5 },
  { id: "p0805", name: "HANDLE BRUSH", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p0806", name: "HEATER MIXER", category: "General", unit: "piece", quantity: 0, costPrice: 2400.0, markup: 40.0, lowStock: 5 },
  { id: "p0807", name: "HEATING ROD F8 2000W", category: "General", unit: "piece", quantity: 0, costPrice: 360.0, markup: 80.6, lowStock: 5 },
  { id: "p0808", name: "HEX NIPPLE 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 42.0, markup: 40.0, lowStock: 5 },
  { id: "p0809", name: "HEX NIPPLE 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 95.0, markup: 35.0, lowStock: 5 },
  { id: "p0810", name: "HEXSAW FRAME", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 460.0, markup: 40.0, lowStock: 5 },
  { id: "p0811", name: "HILTI 10MMX1FT", category: "General", unit: "piece", quantity: 0, costPrice: 140.0, markup: 40.0, lowStock: 5 },
  { id: "p0812", name: "HILTI 12MMX1FT", category: "General", unit: "piece", quantity: 0, costPrice: 190.0, markup: 40.0, lowStock: 5 },
  { id: "p0813", name: "HILTI 13MMX460MM", category: "General", unit: "piece", quantity: 0, costPrice: 420.0, markup: 40.0, lowStock: 5 },
  { id: "p0814", name: "HILTI 14MMX1FT", category: "General", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p0815", name: "HILTI 16MMX1FT", category: "General", unit: "piece", quantity: 0, costPrice: 240.0, markup: 40.0, lowStock: 5 },
  { id: "p0816", name: "HILTI BIT 10MMX6\"", category: "General", unit: "piece", quantity: 0, costPrice: 110.0, markup: 40.0, lowStock: 5 },
  { id: "p0817", name: "HILTI BIT 12MM", category: "General", unit: "piece", quantity: 0, costPrice: 110.0, markup: 40.0, lowStock: 5 },
  { id: "p0818", name: "HILTI BIT 12X350", category: "General", unit: "piece", quantity: 0, costPrice: 230.0, markup: 60.0, lowStock: 5 },
  { id: "p0819", name: "HILTI BIT 14MM", category: "General", unit: "piece", quantity: 0, costPrice: 135.0, markup: 40.0, lowStock: 5 },
  { id: "p0820", name: "HILTI BIT 14X350", category: "General", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p0821", name: "HILTI BIT 14XFT", category: "General", unit: "piece", quantity: 0, costPrice: 250.0, markup: 40.0, lowStock: 5 },
  { id: "p0822", name: "HILTI BIT 16MMX6\"", category: "General", unit: "piece", quantity: 0, costPrice: 150.0, markup: 40.0, lowStock: 5 },
  { id: "p0823", name: "HILTI BIT 18X350", category: "General", unit: "piece", quantity: 0, costPrice: 310.0, markup: 40.0, lowStock: 5 },
  { id: "p0824", name: "HILTI BIT 6MM", category: "General", unit: "piece", quantity: 0, costPrice: 65.0, markup: 40.0, lowStock: 5 },
  { id: "p0825", name: "HILTI BIT 6MM SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p0826", name: "HILTI BIT 6MMX160MM", category: "General", unit: "piece", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p0827", name: "HILTI BIT 6MMX4\"", category: "General", unit: "piece", quantity: 0, costPrice: 100.0, markup: 35.0, lowStock: 5 },
  { id: "p0828", name: "HILTI BIT 6MMX6\"", category: "General", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p0829", name: "HILTI BIT 6X4", category: "General", unit: "piece", quantity: 0, costPrice: 60.0, markup: 40.0, lowStock: 5 },
  { id: "p0830", name: "HILTI BIT 6X6", category: "General", unit: "piece", quantity: 0, costPrice: 540.0, markup: 40.0, lowStock: 5 },
  { id: "p0831", name: "HILTI BIT 8MM", category: "General", unit: "piece", quantity: 0, costPrice: 110.0, markup: 40.0, lowStock: 5 },
  { id: "p0832", name: "HILTI BIT LACILA 10MM", category: "General", unit: "piece", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p0833", name: "HILTI BIT LACILA 12MM", category: "General", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p0834", name: "HILTI BIT LASILA", category: "General", unit: "piece", quantity: 0, costPrice: 170.0, markup: 40.0, lowStock: 5 },
  { id: "p0835", name: "HILTI BIT LASILA 12MM", category: "General", unit: "piece", quantity: 0, costPrice: 210.0, markup: 40.0, lowStock: 5 },
  { id: "p0836", name: "HILTI BIT LASILA 1MM", category: "General", unit: "piece", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p0837", name: "HILTI BIT NPI 12MM", category: "General", unit: "piece", quantity: 0, costPrice: 210.0, markup: 40.0, lowStock: 5 },
  { id: "p0838", name: "HILTI BIT NPI 6MM", category: "General", unit: "piece", quantity: 0, costPrice: 125.0, markup: 40.0, lowStock: 5 },
  { id: "p0839", name: "HILTI BIT NPI 8MM", category: "General", unit: "piece", quantity: 0, costPrice: 170.0, markup: 40.0, lowStock: 5 },
  { id: "p0840", name: "HILTI BIT ZEROX 6X110MM", category: "General", unit: "piece", quantity: 0, costPrice: 135.0, markup: 40.0, lowStock: 5 },
  { id: "p0841", name: "HILTI BIT6X4\"", category: "General", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p0842", name: "HILTI CHESIL 14MM", category: "General", unit: "piece", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p0843", name: "HILTI DRILL BIT 14MM", category: "General", unit: "piece", quantity: 0, costPrice: 125.0, markup: 40.0, lowStock: 5 },
  { id: "p0844", name: "HILTI DRILL BIT 16MM", category: "General", unit: "piece", quantity: 0, costPrice: 155.0, markup: 40.0, lowStock: 5 },
  { id: "p0845", name: "HILTI DRILL BIT 6MM NPI", category: "General", unit: "piece", quantity: 0, costPrice: 170.0, markup: 40.0, lowStock: 5 },
  { id: "p0846", name: "HILTI HIT 6X6", category: "General", unit: "piece", quantity: 0, costPrice: 85.0, markup: 40.0, lowStock: 5 },
  { id: "p0847", name: "HILTI IT 6X4", category: "General", unit: "piece", quantity: 0, costPrice: 65.0, markup: 40.0, lowStock: 5 },
  { id: "p0848", name: "HINGES", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 10.0, markup: 40.0, lowStock: 5 },
  { id: "p0849", name: "HINGES QABZA 4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 70.83, markup: 30.0, lowStock: 5 },
  { id: "p0850", name: "HINGES QABZA 4\" PISTOL", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 71.67, markup: 35.0, lowStock: 5 },
  { id: "p0851", name: "HINGES QABZA IMRAN 4\"\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 60.0, markup: 35.0, lowStock: 5 },
  { id: "p0852", name: "HOCKEY PIPE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 100.0, markup: 40.0, lowStock: 5 },
  { id: "p0853", name: "HOLDER DEEVAR", category: "Electrical", unit: "m plus 76%", quantity: 0, costPrice: 46.67, markup: 30.0, lowStock: 5 },
  { id: "p0854", name: "HOLDER HANGING", category: "Electrical", unit: "piece", quantity: 0, costPrice: 31.67, markup: 30.0, lowStock: 5 },
  { id: "p0855", name: "HOLE SAW SET 11 PCS", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 660.0, markup: 35.0, lowStock: 5 },
  { id: "p0856", name: "HOLE SAW SET 8 PCS", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 760.0, markup: 35.0, lowStock: 5 },
  { id: "p0857", name: "HOME PROTRCT ( GAL 6KG)", category: "General", unit: "piece", quantity: 0, costPrice: 3300.0, markup: 30.0, lowStock: 5 },
  { id: "p0858", name: "HOME PROTRCT 1KG", category: "General", unit: "piece", quantity: 0, costPrice: 600.0, markup: 30.0, lowStock: 5 },
  { id: "p0859", name: "HOOCK", category: "General", unit: "piece", quantity: 0, costPrice: 1150.0, markup: 40.0, lowStock: 5 },
  { id: "p0860", name: "HOOCK PATTI", category: "General", unit: "piece", quantity: 0, costPrice: 15.0, markup: 35.0, lowStock: 5 },
  { id: "p0861", name: "HOSE LEVEL PIPE5.600KG@700", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 700.0, markup: 30.0, lowStock: 5 },
  { id: "p0862", name: "HRS 4\"", category: "General", unit: "piece", quantity: 0, costPrice: 410.0, markup: 40.0, lowStock: 5 },
  { id: "p0863", name: "HRS 5\"", category: "General", unit: "piece", quantity: 0, costPrice: 510.0, markup: 40.0, lowStock: 5 },
  { id: "p0864", name: "HRS 7", category: "General", unit: "piece", quantity: 0, costPrice: 960.0, markup: 40.0, lowStock: 5 },
  { id: "p0865", name: "HRS 7\"", category: "General", unit: "piece", quantity: 0, costPrice: 940.0, markup: 40.0, lowStock: 5 },
  { id: "p0866", name: "HRS DISC", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 900.0, markup: 40.0, lowStock: 5 },
  { id: "p0867", name: "HRS DISC 4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 430.0, markup: 40.0, lowStock: 5 },
  { id: "p0868", name: "HRS DISC 5\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 540.0, markup: 40.0, lowStock: 5 },
  { id: "p0869", name: "HRS DISC 7\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 960.0, markup: 40.0, lowStock: 5 },
  { id: "p0870", name: "HRS DISK 4\"", category: "General", unit: "piece", quantity: 0, costPrice: 400.0, markup: 40.0, lowStock: 5 },
  { id: "p0871", name: "HRS DISK 5\"", category: "General", unit: "piece", quantity: 0, costPrice: 500.0, markup: 40.0, lowStock: 5 },
  { id: "p0872", name: "INDIAN WC", category: "General", unit: "piece", quantity: 0, costPrice: 1400.0, markup: 30.0, lowStock: 5 },
  { id: "p0873", name: "INDIAN WC BABY", category: "General", unit: "piece", quantity: 0, costPrice: 1300.0, markup: 30.0, lowStock: 5 },
  { id: "p0874", name: "INDIAN WC BIG", category: "General", unit: "piece", quantity: 0, costPrice: 1400.0, markup: 40.0, lowStock: 5 },
  { id: "p0875", name: "INDIAN WC SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 1350.0, markup: 40.0, lowStock: 5 },
  { id: "p0876", name: "INSULATION TAPE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 41.5, markup: 30.0, lowStock: 5 },
  { id: "p0877", name: "INVERTOR", category: "General", unit: "piece", quantity: 0, costPrice: 3200.0, markup: 35.0, lowStock: 5 },
  { id: "p0878", name: "JALI  HEAVY 6X6", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 350.0, markup: 40.0, lowStock: 5 },
  { id: "p0879", name: "JALI 3 STAR", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 195.0, markup: 40.0, lowStock: 5 },
  { id: "p0880", name: "JALI 4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 165.0, markup: 40.0, lowStock: 5 },
  { id: "p0881", name: "JALI 4X4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 440.0, markup: 30.0, lowStock: 5 },
  { id: "p0882", name: "JALI 6X6", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 175.0, markup: 40.0, lowStock: 5 },
  { id: "p0883", name: "JALI 6X6 MINHAS", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p0884", name: "JALI APPLE 6X6", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 620.0, markup: 30.0, lowStock: 5 },
  { id: "p0885", name: "JALI EXHUST 12X12", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 370.0, markup: 40.0, lowStock: 5 },
  { id: "p0886", name: "JALI EXHUST 15X15", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 470.0, markup: 40.0, lowStock: 5 },
  { id: "p0887", name: "JALI EXHUST 18X18", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 600.0, markup: 40.0, lowStock: 5 },
  { id: "p0888", name: "JALI H/DUTY", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 270.0, markup: 40.0, lowStock: 5 },
  { id: "p0889", name: "JALI HOLE WALI 9X9", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 395.0, markup: 30.0, lowStock: 5 },
  { id: "p0890", name: "JALI NORMAL 6X6", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p0891", name: "JALI STEEL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 290.0, markup: 40.0, lowStock: 5 },
  { id: "p0892", name: "JALI STEEL HALKI", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p0893", name: "JAMOOR KALBA SMT", category: "General", unit: "piece", quantity: 0, costPrice: 620.0, markup: 40.0, lowStock: 5 },
  { id: "p0894", name: "JHARU BIG", category: "General", unit: "piece", quantity: 0, costPrice: 104.0, markup: 40.0, lowStock: 5 },
  { id: "p0895", name: "JIMSA SAMAD TUBE", category: "Paint", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p0896", name: "JM PIPE 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 240.0, markup: 40.0, lowStock: 5 },
  { id: "p0897", name: "JM PIPE 3/4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 185.0, markup: 40.0, lowStock: 5 },
  { id: "p0898", name: "JOINT BOX 1\"", category: "General", unit: "piece", quantity: 0, costPrice: 28.33, markup: 40.0, lowStock: 5 },
  { id: "p0899", name: "JOINT BOX 3/4", category: "General", unit: "piece", quantity: 0, costPrice: 27.5, markup: 40.0, lowStock: 5 },
  { id: "p0900", name: "JOINT BOX 3/4\"", category: "General", unit: "piece", quantity: 0, costPrice: 26.67, markup: 40.0, lowStock: 5 },
  { id: "p0901", name: "JOINTER", category: "General", unit: "piece", quantity: 0, costPrice: 11.28, markup: 30.0, lowStock: 5 },
  { id: "p0902", name: "JUBLEE CLAMP", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 13.33, markup: 30.0, lowStock: 5 },
  { id: "p0903", name: "JUBLEE CLAMP 1", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 8.25, markup: 40.0, lowStock: 5 },
  { id: "p0904", name: "JUBLEE CLAMP 1\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 10.0, markup: 40.0, lowStock: 5 },
  { id: "p0905", name: "JUBLEE CLAMP 1-1/2", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 27.5, markup: 40.0, lowStock: 5 },
  { id: "p0906", name: "JUBLEE CLAMP 1-1/4", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 18.5, markup: 40.0, lowStock: 5 },
  { id: "p0907", name: "JUBLEE CLAMP 2'", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 40.0, markup: 40.0, lowStock: 5 },
  { id: "p0908", name: "JUBLEE CLAMP 3/4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 20.0, markup: 40.0, lowStock: 5 },
  { id: "p0909", name: "JUBLEE CLIP", category: "General", unit: "piece", quantity: 0, costPrice: 30.0, markup: 40.0, lowStock: 5 },
  { id: "p0910", name: "JUBLEE CLIP 1\"", category: "General", unit: "piece", quantity: 0, costPrice: 8.0, markup: 40.0, lowStock: 5 },
  { id: "p0911", name: "JUBLEE CLIP 1*1/4", category: "General", unit: "piece", quantity: 0, costPrice: 11.67, markup: 40.0, lowStock: 5 },
  { id: "p0912", name: "JUBLEE CLIP 1-1/2\"", category: "General", unit: "piece", quantity: 0, costPrice: 22.5, markup: 40.0, lowStock: 5 },
  { id: "p0913", name: "JUBLEE CLIP 1-1/4", category: "General", unit: "piece", quantity: 0, costPrice: 20.5, markup: 40.0, lowStock: 5 },
  { id: "p0914", name: "JUBLEE CLIP GASS", category: "General", unit: "packet", quantity: 0, costPrice: 1250.0, markup: 40.0, lowStock: 5 },
  { id: "p0915", name: "JUBLI CLIP GASS CLIP", category: "General", unit: "piece", quantity: 0, costPrice: 5.5, markup: 30.0, lowStock: 5 },
  { id: "p0916", name: "JUNCTION BOX 8X10", category: "General", unit: "piece", quantity: 0, costPrice: 133.0, markup: 40.0, lowStock: 5 },
  { id: "p0917", name: "KALBA JAMOOR 6\"", category: "General", unit: "piece", quantity: 0, costPrice: 240.0, markup: 40.0, lowStock: 5 },
  { id: "p0918", name: "KARABDI CHORI SQUARE", category: "General", unit: "piece", quantity: 0, costPrice: 186.67, markup: 40.0, lowStock: 5 },
  { id: "p0919", name: "KARANDI  CHOORI", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 110.0, markup: 40.0, lowStock: 5 },
  { id: "p0920", name: "KARANDI GOL", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 203.33, markup: 40.0, lowStock: 5 },
  { id: "p0921", name: "KARANDI GOL AND SQUARE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 130.0, markup: 40.0, lowStock: 5 },
  { id: "p0922", name: "KARANDI GOL BIG", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 185.0, markup: 40.0, lowStock: 5 },
  { id: "p0923", name: "KARANDI JAMBU BIG", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 220.0, markup: 30.0, lowStock: 5 },
  { id: "p0924", name: "KARANDI JUMBU SMALL GOAL", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 185.0, markup: 30.0, lowStock: 5 },
  { id: "p0925", name: "KARANDI SQUARE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 141.0, markup: 30.0, lowStock: 5 },
  { id: "p0926", name: "KEEL 5", category: "Hardware & Tools", unit: "kg", quantity: 0, costPrice: 310.0, markup: 40.0, lowStock: 5 },
  { id: "p0927", name: "KITCHEN MIXER LONG NECK CHINA", category: "General", unit: "piece", quantity: 0, costPrice: 1400.0, markup: 30.0, lowStock: 5 },
  { id: "p0928", name: "KITCHEN SINK MIXER sonex", category: "General", unit: "piece", quantity: 0, costPrice: 6500.0, markup: 40.0, lowStock: 5 },
  { id: "p0929", name: "KNIFE BLADE", category: "General", unit: "pkts", quantity: 0, costPrice: 75.0, markup: 40.0, lowStock: 5 },
  { id: "p0930", name: "KNIFE BLADE CUTTER", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 27.5, markup: 40.0, lowStock: 5 },
  { id: "p0931", name: "KNIFE BLADE CUTTER H DUTY", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p0932", name: "KNIFE CUTTER PRIDE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 310.0, markup: 40.0, lowStock: 5 },
  { id: "p0933", name: "KNIFE CUTTER ZEROX", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 110.0, markup: 40.0, lowStock: 5 },
  { id: "p0934", name: "KNIFE RED", category: "General", unit: "piece", quantity: 0, costPrice: 110.0, markup: 30.0, lowStock: 5 },
  { id: "p0935", name: "L BRACKET", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 8.0, markup: 30.0, lowStock: 5 },
  { id: "p0936", name: "L BRACKET 1\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 4.0, markup: 30.0, lowStock: 5 },
  { id: "p0937", name: "L BRACKET 1-1/2", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 6.0, markup: 30.0, lowStock: 5 },
  { id: "p0938", name: "L BRAKET 1\"", category: "General", unit: "piece", quantity: 0, costPrice: 210.0, markup: 40.0, lowStock: 5 },
  { id: "p0939", name: "L BRAKET 2\"", category: "General", unit: "piece", quantity: 0, costPrice: 8.02, markup: 40.0, lowStock: 5 },
  { id: "p0940", name: "L KEY SET", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 470.0, markup: 35.0, lowStock: 5 },
  { id: "p0941", name: "L KEY SET STAR", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 560.0, markup: 30.0, lowStock: 5 },
  { id: "p0942", name: "LADDER 7 STEP", category: "General", unit: "piece", quantity: 0, costPrice: 3300.0, markup: 40.0, lowStock: 5 },
  { id: "p0943", name: "LED LIGHT TUFF 7W", category: "General", unit: "piece", quantity: 0, costPrice: 190.0, markup: 40.0, lowStock: 5 },
  { id: "p0944", name: "LEVEL  PRIDE 12\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 680.0, markup: 30.0, lowStock: 5 },
  { id: "p0945", name: "LEVEL 12 OSKAR", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 800.0, markup: 40.0, lowStock: 5 },
  { id: "p0946", name: "LEVEL 12\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 430.0, markup: 40.0, lowStock: 5 },
  { id: "p0947", name: "LEVEL 12\" NEWTON", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 360.0, markup: 40.0, lowStock: 5 },
  { id: "p0948", name: "LEVEL NPT 1FT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 380.0, markup: 40.0, lowStock: 5 },
  { id: "p0949", name: "LEVEL PIPE  4.45KGX605", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 600.0, markup: 40.0, lowStock: 5 },
  { id: "p0950", name: "LEVEL PIPE  4.65KGX620", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 9.65, markup: 40.0, lowStock: 5 },
  { id: "p0951", name: "LGHT BOX 3X3", category: "General", unit: "piece", quantity: 0, costPrice: 38.0, markup: 30.0, lowStock: 5 },
  { id: "p0952", name: "LIGHT BOX", category: "General", unit: "piece", quantity: 0, costPrice: 38.51, markup: 40.0, lowStock: 5 },
  { id: "p0953", name: "LIGHT PLUG", category: "General", unit: "piece", quantity: 0, costPrice: 265.0, markup: 30.0, lowStock: 5 },
  { id: "p0954", name: "LIGHT PLUG 5-IN1", category: "General", unit: "piece", quantity: 0, costPrice: 160.0, markup: 35.0, lowStock: 5 },
  { id: "p0955", name: "LIGHT PLUG 9IN1 PATHAR", category: "General", unit: "piece", quantity: 0, costPrice: 200.0, markup: 35.0, lowStock: 5 },
  { id: "p0956", name: "LIGHT PLUG CHINA FITTING", category: "General", unit: "piece", quantity: 0, costPrice: 328.0, markup: 40.0, lowStock: 5 },
  { id: "p0957", name: "LIGHT PLUG DOUBLE", category: "General", unit: "piece", quantity: 0, costPrice: 370.0, markup: 35.0, lowStock: 5 },
  { id: "p0958", name: "LIGHT PLUG HMA", category: "General", unit: "piece", quantity: 0, costPrice: 230.0, markup: 30.0, lowStock: 5 },
  { id: "p0959", name: "LIGHT PLUG PATHAR", category: "General", unit: "piece", quantity: 0, costPrice: 145.0, markup: 30.0, lowStock: 5 },
  { id: "p0960", name: "LIGHT PLUG PATHAR 9IN ONE", category: "General", unit: "piece", quantity: 0, costPrice: 190.0, markup: 35.0, lowStock: 5 },
  { id: "p0961", name: "LIGHT PLUG PATHER 5IN ONE", category: "General", unit: "piece", quantity: 0, costPrice: 180.0, markup: 30.0, lowStock: 5 },
  { id: "p0962", name: "LIGHT PLUG STONE", category: "General", unit: "piece", quantity: 0, costPrice: 165.0, markup: 40.0, lowStock: 5 },
  { id: "p0963", name: "LIGHT PLUG STONE 9 IN 1", category: "General", unit: "piece", quantity: 0, costPrice: 165.0, markup: 40.0, lowStock: 5 },
  { id: "p0964", name: "LIGHT PLUG STONE SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 170.0, markup: 40.0, lowStock: 5 },
  { id: "p0965", name: "LOCK 60MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 290.0, markup: 40.0, lowStock: 5 },
  { id: "p0966", name: "LOCK BLACK CHINA 50MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 195.0, markup: 40.0, lowStock: 5 },
  { id: "p0967", name: "LOCK DAISI", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 670.0, markup: 40.0, lowStock: 5 },
  { id: "p0968", name: "LOCK DAISI 2KEY", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 300.0, markup: 30.0, lowStock: 5 },
  { id: "p0969", name: "LOCK DAISI 3 KEY", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 350.0, markup: 30.0, lowStock: 5 },
  { id: "p0970", name: "LOCK EAHO 50MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 280.0, markup: 40.0, lowStock: 5 },
  { id: "p0971", name: "LOCK UNION", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 270.0, markup: 40.0, lowStock: 5 },
  { id: "p0972", name: "LOCK WAHU 30MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 190.0, markup: 40.0, lowStock: 5 },
  { id: "p0973", name: "LOCK WAHU 4\0MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 290.0, markup: 40.0, lowStock: 5 },
  { id: "p0974", name: "LOCK WAHU 50MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 360.0, markup: 40.0, lowStock: 5 },
  { id: "p0975", name: "LOCK WEISHAN 30MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p0976", name: "LOCK WEISHAN 40MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 170.0, markup: 40.0, lowStock: 5 },
  { id: "p0977", name: "LOCK WEISHAN 50MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 240.0, markup: 40.0, lowStock: 5 },
  { id: "p0978", name: "LOCK WEISHAN 60MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 290.0, markup: 40.0, lowStock: 5 },
  { id: "p0979", name: "LOCK WEISHAN 70MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 390.0, markup: 40.0, lowStock: 5 },
  { id: "p0980", name: "LOVER BRACKET FAN", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 5781.0, markup: 40.0, lowStock: 5 },
  { id: "p0981", name: "MACHINE SCREW M4X25", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 201.0, markup: 40.0, lowStock: 5 },
  { id: "p0982", name: "MACHINE SCREW M4X35", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 340.0, markup: 40.0, lowStock: 5 },
  { id: "p0983", name: "MACHINE SCREW M4X40", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 261.0, markup: 40.0, lowStock: 5 },
  { id: "p0984", name: "MAGIC", category: "General", unit: "piece", quantity: 0, costPrice: 58.0, markup: 35.0, lowStock: 5 },
  { id: "p0985", name: "MAIZAIL BUTTON", category: "General", unit: "piece", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p0986", name: "MAIZAIL FUSE", category: "General", unit: "piece", quantity: 0, costPrice: 60.0, markup: 40.0, lowStock: 5 },
  { id: "p0987", name: "MAJOOLA CHORAS", category: "General", unit: "piece", quantity: 0, costPrice: 75.0, markup: 40.0, lowStock: 5 },
  { id: "p0988", name: "MAJOOLA KARANDI", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 90.0, markup: 35.0, lowStock: 5 },
  { id: "p0989", name: "MALMAL THAN", category: "General", unit: "piece", quantity: 0, costPrice: 380.0, markup: 35.0, lowStock: 5 },
  { id: "p0990", name: "MALMAN THAN", category: "General", unit: "piece", quantity: 0, costPrice: 360.5, markup: 40.0, lowStock: 5 },
  { id: "p0991", name: "MAN HOLE 15X15", category: "General", unit: "piece", quantity: 0, costPrice: 1180.0, markup: 30.0, lowStock: 5 },
  { id: "p0992", name: "MANHOLE 12X12", category: "General", unit: "piece", quantity: 0, costPrice: 540.0, markup: 30.0, lowStock: 5 },
  { id: "p0993", name: "MANHOLE COVER  12X12", category: "General", unit: "piece", quantity: 0, costPrice: 245.0, markup: 40.0, lowStock: 5 },
  { id: "p0994", name: "MANHOLE COVER  15X15", category: "General", unit: "piece", quantity: 0, costPrice: 575.0, markup: 40.0, lowStock: 5 },
  { id: "p0995", name: "MANHOLE COVER  18X18", category: "General", unit: "piece", quantity: 0, costPrice: 675.0, markup: 40.0, lowStock: 5 },
  { id: "p0996", name: "MANHOLE COVER  24X24", category: "General", unit: "piece", quantity: 0, costPrice: 1480.0, markup: 40.0, lowStock: 5 },
  { id: "p0997", name: "MANHOLE COVER  6X6", category: "General", unit: "piece", quantity: 0, costPrice: 125.0, markup: 40.0, lowStock: 5 },
  { id: "p0998", name: "MANHOLE COVER BLACK 09X09", category: "General", unit: "piece", quantity: 0, costPrice: 700.0, markup: 40.0, lowStock: 5 },
  { id: "p0999", name: "MANHOLE COVER BLACK 12X12", category: "General", unit: "piece", quantity: 0, costPrice: 850.0, markup: 40.0, lowStock: 5 },
  { id: "p1000", name: "MANHOLE COVER BLACK 15X15", category: "General", unit: "piece", quantity: 0, costPrice: 1100.0, markup: 40.0, lowStock: 5 },
  { id: "p1001", name: "MANHOLE COVER BLACK 18X18", category: "General", unit: "piece", quantity: 0, costPrice: 2350.0, markup: 40.0, lowStock: 5 },
  { id: "p1002", name: "MANHOLE COVER BLACK 24X24", category: "General", unit: "piece", quantity: 0, costPrice: 4700.0, markup: 22.0, lowStock: 5 },
  { id: "p1003", name: "MANHOLE COVER12X12", category: "General", unit: "piece", quantity: 0, costPrice: 360.0, markup: 40.0, lowStock: 5 },
  { id: "p1004", name: "MANHOLE CPVER  12X12", category: "General", unit: "piece", quantity: 0, costPrice: 350.0, markup: 40.0, lowStock: 5 },
  { id: "p1005", name: "MANHOLE CPVER  15X15", category: "General", unit: "piece", quantity: 0, costPrice: 675.0, markup: 40.0, lowStock: 5 },
  { id: "p1006", name: "MANHOLE CPVER  18X18", category: "General", unit: "piece", quantity: 0, costPrice: 850.0, markup: 40.0, lowStock: 5 },
  { id: "p1007", name: "MANHOLE CPVER  24X24", category: "General", unit: "piece", quantity: 0, costPrice: 1900.0, markup: 40.0, lowStock: 5 },
  { id: "p1008", name: "MANHOLE MASTER FIT 15X15", category: "General", unit: "piece", quantity: 0, costPrice: 1950.0, markup: 20.0, lowStock: 5 },
  { id: "p1009", name: "MANHOLE MASTER PLUS 15X15", category: "General", unit: "piece", quantity: 0, costPrice: 1750.0, markup: 20.0, lowStock: 5 },
  { id: "p1010", name: "MANHOLE MASTER PLUS 24X24", category: "General", unit: "piece", quantity: 0, costPrice: 3800.0, markup: 20.0, lowStock: 5 },
  { id: "p1011", name: "MANHOLE SONEX 12X12", category: "General", unit: "piece", quantity: 0, costPrice: 370.0, markup: 40.0, lowStock: 5 },
  { id: "p1012", name: "MANHOLE SONEX 18X18", category: "General", unit: "piece", quantity: 0, costPrice: 680.0, markup: 40.0, lowStock: 5 },
  { id: "p1013", name: "MARBAL JELLY", category: "General", unit: "piece", quantity: 0, costPrice: 115.0, markup: 40.0, lowStock: 5 },
  { id: "p1014", name: "MARBLE JELLY", category: "General", unit: "piece", quantity: 0, costPrice: 150.0, markup: 30.0, lowStock: 5 },
  { id: "p1015", name: "MARKER BIG", category: "General", unit: "piece", quantity: 0, costPrice: 55.0, markup: 30.0, lowStock: 5 },
  { id: "p1016", name: "MARKER SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 10.0, markup: 30.0, lowStock: 5 },
  { id: "p1017", name: "MASKING TAPE 1 BIG", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p1018", name: "MASKING TAPE 1\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 35.0, markup: 40.0, lowStock: 5 },
  { id: "p1019", name: "MASKING TAPE 1\" DOUBLE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p1020", name: "MASKING TAPE 1\" THIN", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 35.0, markup: 40.0, lowStock: 5 },
  { id: "p1021", name: "MASKING TAPE 2", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 135.0, markup: 40.0, lowStock: 5 },
  { id: "p1022", name: "MASKING TAPE 2\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p1023", name: "MASKING TAPE 2\" DOUBLE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 145.0, markup: 40.0, lowStock: 5 },
  { id: "p1024", name: "MASKING TAPE 2\" THIN", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 75.0, markup: 40.0, lowStock: 5 },
  { id: "p1025", name: "MASKING TAPE DOUBLE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 160.0, markup: 30.0, lowStock: 5 },
  { id: "p1026", name: "MASKING TAPE DOUBLE 1\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 85.0, markup: 35.0, lowStock: 5 },
  { id: "p1027", name: "MASKING TAPE DOUBLE 2", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 140.0, markup: 40.0, lowStock: 5 },
  { id: "p1028", name: "MAZAIL BOX", category: "General", unit: "piece", quantity: 0, costPrice: 1400.0, markup: 30.0, lowStock: 5 },
  { id: "p1029", name: "MAZAIL BREAKER  A20-A8-A16-A12", category: "Electrical", unit: "piece", quantity: 0, costPrice: 75.0, markup: 30.0, lowStock: 5 },
  { id: "p1030", name: "MAZAIL MOTOR SWITCH", category: "Electrical", unit: "piece", quantity: 0, costPrice: 180.0, markup: 30.0, lowStock: 5 },
  { id: "p1031", name: "MEASURING APE 16 FT LIKE SMT", category: "General", unit: "piece", quantity: 0, costPrice: 220.83, markup: 35.0, lowStock: 5 },
  { id: "p1032", name: "MEASURING TAP 16FT NRK", category: "General", unit: "piece", quantity: 0, costPrice: 140.0, markup: 40.0, lowStock: 5 },
  { id: "p1033", name: "MEASURING TAPE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 0.0, markup: 40.0, lowStock: 5 },
  { id: "p1034", name: "MEASURING TAPE 00FT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 350.0, markup: 30.0, lowStock: 5 },
  { id: "p1035", name: "MEASURING TAPE 100 FT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 360.0, markup: 40.0, lowStock: 5 },
  { id: "p1036", name: "MEASURING TAPE 100FT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 340.0, markup: 35.0, lowStock: 5 },
  { id: "p1037", name: "MEASURING TAPE 100FT CLOTH", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 450.0, markup: 40.0, lowStock: 5 },
  { id: "p1038", name: "MEASURING TAPE 100FT FIBER", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 350.0, markup: 40.0, lowStock: 5 },
  { id: "p1039", name: "MEASURING TAPE 16 FT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p1040", name: "MEASURING TAPE 16 FT CRYSTEL", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 166.67, markup: 35.0, lowStock: 5 },
  { id: "p1041", name: "MEASURING TAPE 16 FT GOOD", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 148.33, markup: 40.0, lowStock: 5 },
  { id: "p1042", name: "MEASURING TAPE 16 FT ORANGE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p1043", name: "MEASURING TAPE 16 FT TIGER", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 115.0, markup: 40.0, lowStock: 5 },
  { id: "p1044", name: "MEASURING TAPE 16FT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 104.17, markup: 30.0, lowStock: 5 },
  { id: "p1045", name: "MEASURING TAPE 1G FT ATI", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 130.0, markup: 40.0, lowStock: 5 },
  { id: "p1046", name: "MEASURING TAPE 5 MTR", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 123.33, markup: 40.0, lowStock: 5 },
  { id: "p1047", name: "MEASURING TAPE 50 FT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 240.0, markup: 40.0, lowStock: 5 },
  { id: "p1048", name: "MEASURING TAPE 50FT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 230.0, markup: 30.0, lowStock: 5 },
  { id: "p1049", name: "MEASURING TAPE 5MTR GOOD", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 290.0, markup: 40.0, lowStock: 5 },
  { id: "p1050", name: "MEASURING TAPE 7.5", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 358.0, markup: 30.0, lowStock: 5 },
  { id: "p1051", name: "MEASURING TAPE CLOTH 100FT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 480.0, markup: 40.0, lowStock: 5 },
  { id: "p1052", name: "MEASURING TAPE HEAVY", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p1053", name: "MEASURING TAPE JHDFG 16 FT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 145.0, markup: 40.0, lowStock: 5 },
  { id: "p1054", name: "MEASURING TAPE LHS BLACK 16FT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 210.0, markup: 30.0, lowStock: 5 },
  { id: "p1055", name: "MEASURING TAPE NRK 16FT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 183.33, markup: 40.0, lowStock: 5 },
  { id: "p1056", name: "MEASURING TAPE RABIT 100FT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 295.0, markup: 40.0, lowStock: 5 },
  { id: "p1057", name: "MEASURING TAPE RABIT 50FT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 175.0, markup: 40.0, lowStock: 5 },
  { id: "p1058", name: "MEASURING TAPE TIGER 16F", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 121.67, markup: 40.0, lowStock: 5 },
  { id: "p1059", name: "MEASURING TAPE TIGER 16FT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 104.17, markup: 35.0, lowStock: 5 },
  { id: "p1060", name: "MEASURING TAPE W TOOL YELLOW 16FT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 150.0, markup: 30.0, lowStock: 5 },
  { id: "p1061", name: "MEDIUM SOLUTION", category: "Paint", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p1062", name: "MESAURING TAPE", category: "General", unit: "piece", quantity: 0, costPrice: 128.33, markup: 40.0, lowStock: 5 },
  { id: "p1063", name: "METAL BOX 3X3", category: "General", unit: "piece", quantity: 0, costPrice: 115.0, markup: 40.0, lowStock: 5 },
  { id: "p1064", name: "METAL BOX 3X6", category: "General", unit: "piece", quantity: 0, costPrice: 165.0, markup: 40.0, lowStock: 5 },
  { id: "p1065", name: "METAL BOX 6X3", category: "General", unit: "piece", quantity: 0, costPrice: 155.0, markup: 40.0, lowStock: 5 },
  { id: "p1066", name: "METAL BOX 8X10", category: "General", unit: "piece", quantity: 0, costPrice: 466.67, markup: 40.0, lowStock: 5 },
  { id: "p1067", name: "METAL BOX ORANGE 3X3", category: "General", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p1068", name: "METAL BOX ORANGE 6X3", category: "General", unit: "piece", quantity: 0, costPrice: 150.0, markup: 40.0, lowStock: 5 },
  { id: "p1069", name: "METER BOX  DOUBLE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 2700.0, markup: 35.0, lowStock: 5 },
  { id: "p1070", name: "METER BOX DOUBLE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 2500.0, markup: 30.0, lowStock: 5 },
  { id: "p1071", name: "METER BOX SINGLE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 1650.0, markup: 35.0, lowStock: 5 },
  { id: "p1072", name: "MEZAIL BOX", category: "General", unit: "piece", quantity: 0, costPrice: 1550.0, markup: 40.0, lowStock: 5 },
  { id: "p1073", name: "MEZAIL BOX 0.75", category: "General", unit: "piece", quantity: 0, costPrice: 1350.0, markup: 35.0, lowStock: 5 },
  { id: "p1074", name: "MEZAIL BUTTON", category: "General", unit: "piece", quantity: 0, costPrice: 150.0, markup: 40.0, lowStock: 5 },
  { id: "p1075", name: "MIRROR", category: "General", unit: "piece", quantity: 0, costPrice: 460.0, markup: 40.0, lowStock: 5 },
  { id: "p1076", name: "MIRROR CLIP", category: "General", unit: "piece", quantity: 0, costPrice: 19.3, markup: 30.0, lowStock: 5 },
  { id: "p1077", name: "MIRROR QALAM SHEESHA QALAM", category: "General", unit: "piece", quantity: 0, costPrice: 220.0, markup: 35.0, lowStock: 5 },
  { id: "p1078", name: "MIXER CONNECTION PIPE", category: "Plumbing & Sanitary", unit: "pair", quantity: 0, costPrice: 160.0, markup: 30.0, lowStock: 5 },
  { id: "p1079", name: "MIXER CONNECTION PIPE CHROME", category: "Plumbing & Sanitary", unit: "jori", quantity: 0, costPrice: 220.0, markup: 30.0, lowStock: 5 },
  { id: "p1080", name: "MIXER GODAY", category: "General", unit: "piece", quantity: 0, costPrice: 264.0, markup: 40.0, lowStock: 5 },
  { id: "p1081", name: "MIXER JORI", category: "General", unit: "no", quantity: 0, costPrice: 137.5, markup: 40.0, lowStock: 5 },
  { id: "p1082", name: "MIXER JORI 18\"", category: "General", unit: "nos", quantity: 0, costPrice: 87.5, markup: 40.0, lowStock: 5 },
  { id: "p1083", name: "MIXER JORI 24\"", category: "General", unit: "nos", quantity: 0, costPrice: 92.5, markup: 40.0, lowStock: 5 },
  { id: "p1084", name: "MIXER NUT BRASS", category: "General", unit: "piece", quantity: 0, costPrice: 220.0, markup: 30.0, lowStock: 5 },
  { id: "p1085", name: "MIXER PIPE", category: "Plumbing & Sanitary", unit: "pair", quantity: 0, costPrice: 280.0, markup: 35.0, lowStock: 5 },
  { id: "p1086", name: "MOB", category: "General", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p1087", name: "MOBILE MULTI PLUG", category: "General", unit: "piece", quantity: 0, costPrice: 31.67, markup: 40.0, lowStock: 5 },
  { id: "p1088", name: "MOCHI KEEL", category: "Hardware & Tools", unit: "1", quantity: 0, costPrice: 450.0, markup: 40.0, lowStock: 5 },
  { id: "p1089", name: "MOOCHI NAIL", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 280.0, markup: 30.0, lowStock: 5 },
  { id: "p1090", name: "MOVILETH GLUE", category: "Paint", unit: "piece", quantity: 0, costPrice: 635.0, markup: 40.0, lowStock: 5 },
  { id: "p1091", name: "MOWCHI KEEL 3/4", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 340.0, markup: 35.0, lowStock: 5 },
  { id: "p1092", name: "MS MASTER CLASS B 3\"", category: "General", unit: "piece", quantity: 0, costPrice: 1000.0, markup: 40.0, lowStock: 5 },
  { id: "p1093", name: "MS MASTER CLASS B 4\"", category: "General", unit: "piece", quantity: 0, costPrice: 1500.0, markup: 40.0, lowStock: 5 },
  { id: "p1094", name: "MS MASTER DELUX 3", category: "General", unit: "piece", quantity: 0, costPrice: 875.0, markup: 40.0, lowStock: 5 },
  { id: "p1095", name: "MS MASTER DELUX 4", category: "General", unit: "piece", quantity: 0, costPrice: 1125.0, markup: 40.0, lowStock: 5 },
  { id: "p1096", name: "MULIM SHOWER BODY", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 175.0, markup: 30.0, lowStock: 5 },
  { id: "p1097", name: "MULTI  POWER PLUG 3 PIN", category: "General", unit: "piece", quantity: 0, costPrice: 108.33, markup: 30.0, lowStock: 5 },
  { id: "p1098", name: "MULTI PLUG", category: "General", unit: "piece", quantity: 0, costPrice: 55.0, markup: 30.0, lowStock: 5 },
  { id: "p1099", name: "MUSLIM SHOER BODY", category: "General", unit: "piece", quantity: 0, costPrice: 180.0, markup: 30.0, lowStock: 5 },
  { id: "p1100", name: "MUSLIM SHOER BODY HEAD", category: "General", unit: "piece", quantity: 0, costPrice: 180.0, markup: 35.0, lowStock: 5 },
  { id: "p1101", name: "MUSLIM SHOER COMPLETE", category: "General", unit: "piece", quantity: 0, costPrice: 750.0, markup: 40.0, lowStock: 5 },
  { id: "p1102", name: "MUSLIM SHOWER BODY", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 150.0, markup: 40.0, lowStock: 5 },
  { id: "p1103", name: "MUSLIM SHOWER BODY H DUTY", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 260.0, markup: 40.0, lowStock: 5 },
  { id: "p1104", name: "MUSLIM SHOWER COMPLETE SET", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 640.0, markup: 40.0, lowStock: 5 },
  { id: "p1105", name: "MUSLIM SHOWER GROHE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 750.0, markup: 40.0, lowStock: 5 },
  { id: "p1106", name: "MUSLIM SHOWER HEAD", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 170.0, markup: 35.0, lowStock: 5 },
  { id: "p1107", name: "MUSLIM SHOWER KRISS", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 350.0, markup: 35.0, lowStock: 5 },
  { id: "p1108", name: "MUSLIM SHOWER MASTER SHATTAF", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 625.0, markup: 40.0, lowStock: 5 },
  { id: "p1109", name: "MUSLIM SHOWER SONEX", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1750.0, markup: 40.0, lowStock: 5 },
  { id: "p1110", name: "MUSLIM SHOWER SS", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 480.0, markup: 35.0, lowStock: 5 },
  { id: "p1111", name: "NAIL 5\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 360.0, markup: 35.0, lowStock: 5 },
  { id: "p1112", name: "NARIYAL BRUSH", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p1113", name: "NECK ONLY", category: "General", unit: "piece", quantity: 0, costPrice: 195.0, markup: 40.0, lowStock: 5 },
  { id: "p1114", name: "NEEL", category: "General", unit: "piece", quantity: 0, costPrice: 38.33, markup: 40.0, lowStock: 5 },
  { id: "p1115", name: "NEEL 40G", category: "General", unit: "piece", quantity: 0, costPrice: 40.0, markup: 30.0, lowStock: 5 },
  { id: "p1116", name: "NEEL BLUE BIRD", category: "General", unit: "piece", quantity: 0, costPrice: 40.0, markup: 40.0, lowStock: 5 },
  { id: "p1117", name: "NEEL TOTA", category: "General", unit: "piece", quantity: 0, costPrice: 46.67, markup: 40.0, lowStock: 5 },
  { id: "p1118", name: "NON RETURN VALVE 1/2\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 380.0, markup: 40.0, lowStock: 5 },
  { id: "p1119", name: "NON RETURN VALVE 3/4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 630.0, markup: 40.0, lowStock: 5 },
  { id: "p1120", name: "NOSE PLAS", category: "General", unit: "piece", quantity: 0, costPrice: 510.0, markup: 40.0, lowStock: 5 },
  { id: "p1121", name: "NRV 25MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 880.0, markup: 35.0, lowStock: 5 },
  { id: "p1122", name: "NRV 32MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1180.0, markup: 35.0, lowStock: 5 },
  { id: "p1123", name: "NRV NON RETURN VALVE 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1190.0, markup: 35.0, lowStock: 5 },
  { id: "p1124", name: "NRV NON RETURN VALVE 3/4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 880.0, markup: 35.0, lowStock: 5 },
  { id: "p1125", name: "NUT", category: "General", unit: "kg", quantity: 0, costPrice: 600.0, markup: 30.0, lowStock: 5 },
  { id: "p1126", name: "NUT 16MM", category: "General", unit: "piece", quantity: 0, costPrice: 15.0, markup: 40.0, lowStock: 5 },
  { id: "p1127", name: "OIL PAINT", category: "Paint", unit: "piece", quantity: 0, costPrice: 130.0, markup: 40.0, lowStock: 5 },
  { id: "p1128", name: "OIL PAINT   0.91 ML", category: "Paint", unit: "piece", quantity: 0, costPrice: 417.5, markup: 30.0, lowStock: 5 },
  { id: "p1129", name: "OIL PAINT BLACK", category: "Paint", unit: "piece", quantity: 0, costPrice: 1550.0, markup: 40.0, lowStock: 5 },
  { id: "p1130", name: "ON", category: "General", unit: "piece", quantity: 0, costPrice: 156.67, markup: 40.0, lowStock: 5 },
  { id: "p1131", name: "ONLY JALI 3\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 15.0, markup: 40.0, lowStock: 5 },
  { id: "p1132", name: "ONLY JALI 4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 19.0, markup: 40.0, lowStock: 5 },
  { id: "p1133", name: "OPEN BOX", category: "General", unit: "piece", quantity: 0, costPrice: 45.0, markup: 40.0, lowStock: 5 },
  { id: "p1134", name: "OPEN BOX 3X3", category: "General", unit: "piece", quantity: 0, costPrice: 35.0, markup: 40.0, lowStock: 5 },
  { id: "p1135", name: "OPEN BOX 4X4", category: "General", unit: "piece", quantity: 0, costPrice: 31.67, markup: 40.0, lowStock: 5 },
  { id: "p1136", name: "OPEN BOX 4X7", category: "General", unit: "piece", quantity: 0, costPrice: 60.0, markup: 40.0, lowStock: 5 },
  { id: "p1137", name: "OPEN BOX 8X10", category: "General", unit: "piece", quantity: 0, costPrice: 125.0, markup: 40.0, lowStock: 5 },
  { id: "p1138", name: "OSAKA TAPE", category: "General", unit: "piece", quantity: 0, costPrice: 42.0, markup: 35.0, lowStock: 5 },
  { id: "p1139", name: "OSCAR DISC 1MMX4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 50.0, markup: 30.0, lowStock: 5 },
  { id: "p1140", name: "OSKAR DICK 7\"", category: "General", unit: "piece", quantity: 0, costPrice: 1280.0, markup: 30.0, lowStock: 5 },
  { id: "p1141", name: "P TRAP", category: "General", unit: "piece", quantity: 0, costPrice: 300.0, markup: 40.0, lowStock: 5 },
  { id: "p1142", name: "P TRAP 3\" BURJ", category: "General", unit: "piece", quantity: 0, costPrice: 296.0, markup: 40.0, lowStock: 5 },
  { id: "p1143", name: "P TRAP 4", category: "General", unit: "piece", quantity: 0, costPrice: 455.0, markup: 40.0, lowStock: 5 },
  { id: "p1144", name: "P TRAP CLAY", category: "General", unit: "piece", quantity: 0, costPrice: 280.0, markup: 40.0, lowStock: 5 },
  { id: "p1145", name: "PACKING TAPE", category: "General", unit: "piece", quantity: 0, costPrice: 75.0, markup: 40.0, lowStock: 5 },
  { id: "p1146", name: "PACKING TAPE 2\"", category: "General", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p1147", name: "PACKING TAPE BIG", category: "General", unit: "piece", quantity: 0, costPrice: 110.0, markup: 40.0, lowStock: 5 },
  { id: "p1148", name: "PACKING TAPE DOUBLE", category: "General", unit: "piece", quantity: 0, costPrice: 95.0, markup: 35.0, lowStock: 5 },
  { id: "p1149", name: "PACKING TAPE ORDINARY", category: "General", unit: "piece", quantity: 0, costPrice: 70.0, markup: 35.0, lowStock: 5 },
  { id: "p1150", name: "PAD LOCK CHINA 32MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 93.33, markup: 40.0, lowStock: 5 },
  { id: "p1151", name: "PAD LOCK CHINA 38MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 108.33, markup: 40.0, lowStock: 5 },
  { id: "p1152", name: "PAD LOCK CHINA 50MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 168.0, markup: 40.0, lowStock: 5 },
  { id: "p1153", name: "PAD LOCK GOLDEN 32MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 130.0, markup: 35.0, lowStock: 5 },
  { id: "p1154", name: "PAD LOCK GOLDEN 38MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 146.67, markup: 35.0, lowStock: 5 },
  { id: "p1155", name: "PAD LOCK GOLDEN 50MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 226.67, markup: 35.0, lowStock: 5 },
  { id: "p1156", name: "PADLO  2KG", category: "General", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p1157", name: "PADLO 1 KG", category: "General", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p1158", name: "PADLO 1KG", category: "General", unit: "piece", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p1159", name: "PADLO 2KG", category: "General", unit: "piece", quantity: 0, costPrice: 200.0, markup: 35.0, lowStock: 5 },
  { id: "p1160", name: "PAIN ROLLER SMALL REFILL MOTI 4\"", category: "Paint", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p1161", name: "PAINT", category: "Paint", unit: "piece", quantity: 0, costPrice: 140.0, markup: 40.0, lowStock: 5 },
  { id: "p1162", name: "PAINT BLACK", category: "Paint", unit: "piece", quantity: 0, costPrice: 140.0, markup: 30.0, lowStock: 5 },
  { id: "p1163", name: "PAINT BRUSH 1-1/2\" SINGLE", category: "Paint", unit: "piece", quantity: 0, costPrice: 37.0, markup: 30.0, lowStock: 5 },
  { id: "p1164", name: "PAINT BRUSH 2\" DOUBLE", category: "Paint", unit: "piece", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p1165", name: "PAINT BRUSH 2\" SINGLE", category: "Paint", unit: "piece", quantity: 0, costPrice: 50.0, markup: 30.0, lowStock: 5 },
  { id: "p1166", name: "PAINT BRUSH 3\"", category: "Paint", unit: "piece", quantity: 0, costPrice: 65.0, markup: 40.0, lowStock: 5 },
  { id: "p1167", name: "PAINT BRUSH 3\" DOUBLE", category: "Paint", unit: "piece", quantity: 0, costPrice: 135.0, markup: 40.0, lowStock: 5 },
  { id: "p1168", name: "PAINT BRUSH 3\" SINGLE", category: "Paint", unit: "piece", quantity: 0, costPrice: 75.0, markup: 30.0, lowStock: 5 },
  { id: "p1169", name: "PAINT BRUSH 4\"", category: "Paint", unit: "piece", quantity: 0, costPrice: 100.0, markup: 40.0, lowStock: 5 },
  { id: "p1170", name: "PAINT BRUSH 4\" DOUBLE", category: "Paint", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p1171", name: "PAINT BRUSH 5\"", category: "Paint", unit: "piece", quantity: 0, costPrice: 150.0, markup: 40.0, lowStock: 5 },
  { id: "p1172", name: "PAINT BRUSH DOUBLE 2\"", category: "Paint", unit: "piece", quantity: 0, costPrice: 90.0, markup: 30.0, lowStock: 5 },
  { id: "p1173", name: "PAINT BRUSH DOUBLE 3\"", category: "Paint", unit: "piece", quantity: 0, costPrice: 135.0, markup: 30.0, lowStock: 5 },
  { id: "p1174", name: "PAINT BRUSH PLASTIC HANDLE 3\"", category: "Paint", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p1175", name: "PAINT BRUSH PLASTIC HANDLE 4\"", category: "Paint", unit: "piece", quantity: 0, costPrice: 150.0, markup: 40.0, lowStock: 5 },
  { id: "p1176", name: "PAINT BRUSH PVC H/DUTYHANDLE 5\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 260.0, markup: 40.0, lowStock: 5 },
  { id: "p1177", name: "PAINT BRUSH TRIDENT 2\"", category: "Paint", unit: "piece", quantity: 0, costPrice: 140.0, markup: 40.0, lowStock: 5 },
  { id: "p1178", name: "PAINT BRUSH TRIDENT 3\"", category: "Paint", unit: "piece", quantity: 0, costPrice: 210.0, markup: 40.0, lowStock: 5 },
  { id: "p1179", name: "PAINT BRUSH WOOD HANDLE 2\" APPLE", category: "Paint", unit: "piece", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p1180", name: "PAINT BRUSH WOOD HANDLE 3\" APPLE", category: "Paint", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p1181", name: "PAINT BRUSH WOOD HANDLE 4\" APPLE", category: "Paint", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p1182", name: "PAINT BRUSH WOOD HANDLE 5\"", category: "Paint", unit: "piece", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p1183", name: "PAINT BRUSH WOOD HANDLE 5\" APPLE", category: "Paint", unit: "piece", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p1184", name: "PAINT DABBI", category: "Paint", unit: "piece", quantity: 0, costPrice: 130.0, markup: 40.0, lowStock: 5 },
  { id: "p1185", name: "PAINT QALAM RED", category: "Paint", unit: "packet", quantity: 0, costPrice: 190.0, markup: 35.0, lowStock: 5 },
  { id: "p1186", name: "PAINT QALAM WHITE", category: "Paint", unit: "packet", quantity: 0, costPrice: 260.0, markup: 35.0, lowStock: 5 },
  { id: "p1187", name: "PAINT ROLLAR", category: "Paint", unit: "piece", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p1188", name: "PAINT ROLLAR STIC", category: "Paint", unit: "piece", quantity: 0, costPrice: 290.0, markup: 40.0, lowStock: 5 },
  { id: "p1189", name: "PAINT ROLLER", category: "Paint", unit: "piece", quantity: 0, costPrice: 190.0, markup: 30.0, lowStock: 5 },
  { id: "p1190", name: "PAINT ROLLER GOOD", category: "Paint", unit: "piece", quantity: 0, costPrice: 360.0, markup: 40.0, lowStock: 5 },
  { id: "p1191", name: "PAINT ROLLER SMALL", category: "Paint", unit: "piece", quantity: 0, costPrice: 33.0, markup: 40.0, lowStock: 5 },
  { id: "p1192", name: "PAINT ROLLER SMALL FILLER", category: "Paint", unit: "piece", quantity: 0, costPrice: 35.0, markup: 40.0, lowStock: 5 },
  { id: "p1193", name: "PAINT ROLLER STIC SMALL", category: "Paint", unit: "piece", quantity: 0, costPrice: 70.0, markup: 30.0, lowStock: 5 },
  { id: "p1194", name: "PAINT ROLLRER", category: "Paint", unit: "piece", quantity: 0, costPrice: 210.0, markup: 40.0, lowStock: 5 },
  { id: "p1195", name: "PAINT RULLI SMALL", category: "Paint", unit: "piece", quantity: 0, costPrice: 45.0, markup: 30.0, lowStock: 5 },
  { id: "p1196", name: "PAINT SPRAY", category: "Paint", unit: "piece", quantity: 0, costPrice: 330.0, markup: 40.0, lowStock: 5 },
  { id: "p1197", name: "PAINT SPRRAY", category: "Paint", unit: "piece", quantity: 0, costPrice: 330.0, markup: 30.0, lowStock: 5 },
  { id: "p1198", name: "PAK BILAL CABLE 3/29", category: "Electrical", unit: "piece", quantity: 0, costPrice: 3600.0, markup: 40.0, lowStock: 5 },
  { id: "p1199", name: "PALO 1KG", category: "General", unit: "piece", quantity: 0, costPrice: 100.0, markup: 35.0, lowStock: 5 },
  { id: "p1200", name: "PAPER CUTTER", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 30.0, markup: 35.0, lowStock: 5 },
  { id: "p1201", name: "PAPER CUTTER BLADE", category: "Hardware & Tools", unit: "pkts", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p1202", name: "PAPER CUTTER NORMAL", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 28.33, markup: 40.0, lowStock: 5 },
  { id: "p1203", name: "PARDA BRACKET GOLDEN", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 36.0, markup: 30.0, lowStock: 5 },
  { id: "p1204", name: "PARDA BREACKET", category: "General", unit: "piece", quantity: 0, costPrice: 40.0, markup: 40.0, lowStock: 5 },
  { id: "p1205", name: "PARDA HOOCK SINGLE", category: "General", unit: "piece", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p1206", name: "PARDA PHOOL", category: "General", unit: "piece", quantity: 0, costPrice: 65.0, markup: 40.0, lowStock: 5 },
  { id: "p1207", name: "PARDAN WOOD PHATTI", category: "General", unit: "piece", quantity: 0, costPrice: 14.0, markup: 40.0, lowStock: 5 },
  { id: "p1208", name: "PATTRI", category: "General", unit: "piece", quantity: 0, costPrice: 300.0, markup: 40.0, lowStock: 5 },
  { id: "p1209", name: "PEDISTAL FAN", category: "General", unit: "piece", quantity: 0, costPrice: 7800.0, markup: 40.0, lowStock: 5 },
  { id: "p1210", name: "PEELA AMBER", category: "General", unit: "piece", quantity: 0, costPrice: 93.33, markup: 40.0, lowStock: 5 },
  { id: "p1211", name: "PIANO BOARD 6 HOLE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 70.0, markup: 30.0, lowStock: 5 },
  { id: "p1212", name: "PIANO BUTTON", category: "Electrical", unit: "piece", quantity: 0, costPrice: 27.5, markup: 40.0, lowStock: 5 },
  { id: "p1213", name: "PIANO LIGHT PLUG", category: "Electrical", unit: "piece", quantity: 0, costPrice: 100.0, markup: 30.0, lowStock: 5 },
  { id: "p1214", name: "PIANO SOCKET", category: "Electrical", unit: "piece", quantity: 0, costPrice: 28.33, markup: 30.0, lowStock: 5 },
  { id: "p1215", name: "PIANO SOCKET 1IN7", category: "Electrical", unit: "piece", quantity: 0, costPrice: 105.0, markup: 35.0, lowStock: 5 },
  { id: "p1216", name: "PIANO SOCKET PATHAR", category: "Electrical", unit: "piece", quantity: 0, costPrice: 40.83, markup: 30.0, lowStock: 5 },
  { id: "p1217", name: "PIANO SOCKET PATHER", category: "Electrical", unit: "piece", quantity: 0, costPrice: 31.67, markup: 40.0, lowStock: 5 },
  { id: "p1218", name: "PIANO SOCKET PVC", category: "Electrical", unit: "piece", quantity: 0, costPrice: 30.0, markup: 30.0, lowStock: 5 },
  { id: "p1219", name: "PIANO SOCKET SWITCH", category: "Electrical", unit: "piece", quantity: 0, costPrice: 25.83, markup: 40.0, lowStock: 5 },
  { id: "p1220", name: "PIANO SWITCH SOCKET", category: "Electrical", unit: "piece", quantity: 0, costPrice: 360.0, markup: 35.0, lowStock: 5 },
  { id: "p1221", name: "PIC AX GANTI", category: "General", unit: "piece", quantity: 0, costPrice: 716.0, markup: 40.0, lowStock: 5 },
  { id: "p1222", name: "PICK AX HANDLE", category: "General", unit: "piece", quantity: 0, costPrice: 155.0, markup: 35.0, lowStock: 5 },
  { id: "p1223", name: "PICKAX GAINTI", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 733.0, markup: 40.0, lowStock: 5 },
  { id: "p1224", name: "PIPE 1\" GM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 285.0, markup: 40.0, lowStock: 5 },
  { id: "p1225", name: "PIPE 1.8KG   2\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 423.0, markup: 40.0, lowStock: 5 },
  { id: "p1226", name: "PIPE 3/4\" GM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 195.0, markup: 40.0, lowStock: 5 },
  { id: "p1227", name: "PIPE 4 KG 3\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 940.0, markup: 40.0, lowStock: 5 },
  { id: "p1228", name: "PIPE 6KG 4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1410.0, markup: 40.0, lowStock: 5 },
  { id: "p1229", name: "PIPE NIPPLE  3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 48.33, markup: 40.0, lowStock: 5 },
  { id: "p1230", name: "PIPE NIPPLE  3/4X3\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 36.67, markup: 40.0, lowStock: 5 },
  { id: "p1231", name: "PIPE NIPPLE  3/4X4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 45.0, markup: 40.0, lowStock: 5 },
  { id: "p1232", name: "PIPE NIPPLE 1\"X12\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 235.0, markup: 40.0, lowStock: 5 },
  { id: "p1233", name: "PIPE NIPPLE 1-1/4X12", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 315.0, markup: 30.0, lowStock: 5 },
  { id: "p1234", name: "PIPE NIPPLE 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 41.92, markup: 40.0, lowStock: 5 },
  { id: "p1235", name: "PIPE NIPPLE 1/2X12\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 115.0, markup: 40.0, lowStock: 5 },
  { id: "p1236", name: "PIPE NIPPLE 1/2X4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 50.0, markup: 35.0, lowStock: 5 },
  { id: "p1237", name: "PIPE NIPPLE 1/2X6\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 60.0, markup: 40.0, lowStock: 5 },
  { id: "p1238", name: "PIPE NIPPLE 1X12", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 330.0, markup: 30.0, lowStock: 5 },
  { id: "p1239", name: "PIPE NIPPLE 3/4\"X12\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 165.0, markup: 60.0, lowStock: 5 },
  { id: "p1240", name: "PIPE NIPPLE 3/4X3", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 35.0, markup: 40.0, lowStock: 5 },
  { id: "p1241", name: "PIPE NIPPLE 3/4X4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 47.5, markup: 30.0, lowStock: 5 },
  { id: "p1242", name: "PIPE NIPPLE 3/4X5", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 55.0, markup: 40.0, lowStock: 5 },
  { id: "p1243", name: "PIPE NIPPLE 3/4X6", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 65.0, markup: 35.0, lowStock: 5 },
  { id: "p1244", name: "PIPE WRENCH 12\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 800.0, markup: 40.0, lowStock: 5 },
  { id: "p1245", name: "PIPE WRENCH 14\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 920.0, markup: 35.0, lowStock: 5 },
  { id: "p1246", name: "PIPIE WRENCH 12", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 835.0, markup: 40.0, lowStock: 5 },
  { id: "p1247", name: "PLASTIC GURMALA BIG", category: "General", unit: "piece", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p1248", name: "PLASTIC GURMALA MEDIUM", category: "General", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p1249", name: "PLASTIC GURMALA SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 50.0, markup: 40.0, lowStock: 5 },
  { id: "p1250", name: "PLAT DOORI", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 166.0, markup: 40.0, lowStock: 5 },
  { id: "p1251", name: "PLIER", category: "General", unit: "piece", quantity: 0, costPrice: 380.0, markup: 35.0, lowStock: 5 },
  { id: "p1252", name: "PLIER BLACK 8\"", category: "General", unit: "piece", quantity: 0, costPrice: 500.0, markup: 40.0, lowStock: 5 },
  { id: "p1253", name: "PLIER CHINA", category: "General", unit: "piece", quantity: 0, costPrice: 210.0, markup: 40.0, lowStock: 5 },
  { id: "p1254", name: "PLIER H DUTY", category: "General", unit: "piece", quantity: 0, costPrice: 550.0, markup: 40.0, lowStock: 5 },
  { id: "p1255", name: "PLIER OK FORT HIGH QLTY", category: "General", unit: "piece", quantity: 0, costPrice: 220.0, markup: 30.0, lowStock: 5 },
  { id: "p1256", name: "PLIER ORANGE", category: "General", unit: "piece", quantity: 0, costPrice: 300.0, markup: 40.0, lowStock: 5 },
  { id: "p1257", name: "PLIER PRIDE BLACK", category: "General", unit: "piece", quantity: 0, costPrice: 650.0, markup: 40.0, lowStock: 5 },
  { id: "p1258", name: "PLIER RED", category: "General", unit: "piece", quantity: 0, costPrice: 310.0, markup: 40.0, lowStock: 5 },
  { id: "p1259", name: "PLOT DORI", category: "General", unit: "piece", quantity: 0, costPrice: 177.0, markup: 30.0, lowStock: 5 },
  { id: "p1260", name: "PLOT DORRI", category: "General", unit: "1310", quantity: 0, costPrice: 163.75, markup: 30.0, lowStock: 5 },
  { id: "p1261", name: "PLOT RASSI", category: "General", unit: "piece", quantity: 0, costPrice: 153.0, markup: 40.0, lowStock: 5 },
  { id: "p1262", name: "PLUG", category: "General", unit: "piece", quantity: 0, costPrice: 2.0, markup: 40.0, lowStock: 5 },
  { id: "p1263", name: "PLUG 1/2", category: "General", unit: "piece", quantity: 0, costPrice: 2.8, markup: 40.0, lowStock: 5 },
  { id: "p1264", name: "PLUMBOB SAAL", category: "General", unit: "piece", quantity: 0, costPrice: 191.67, markup: 35.0, lowStock: 5 },
  { id: "p1265", name: "POLY PIPE CONDUIT 1-1/4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 300.0, markup: 40.0, lowStock: 5 },
  { id: "p1266", name: "POLYPIPE  1\"(31.1KGX@255=7930/50) 79 PER PCS DEEVAR", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 79.0, markup: 30.0, lowStock: 5 },
  { id: "p1267", name: "POLYPIPE  1\"(80KGX@255=20400/50)102 PER PCS ROOF", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 102.0, markup: 30.0, lowStock: 5 },
  { id: "p1268", name: "POLYPIPE  3/4\"(43.9KGX@255=11195/50)56PER PCS DEEVAR", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 56.0, markup: 30.0, lowStock: 5 },
  { id: "p1269", name: "POLYPIPE3/4(90.7KGX@255=23128/50) 77 PER PCS ROOF", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 77.0, markup: 30.0, lowStock: 5 },
  { id: "p1270", name: "POLYTHEEN SHEET 7FT 21.5KG/440PER KG", category: "General", unit: "piece", quantity: 0, costPrice: 440.0, markup: 40.0, lowStock: 5 },
  { id: "p1271", name: "POLYTHEEN SHEET SUPER JEELANI 270FT=82MTR", category: "General", unit: "piece", quantity: 0, costPrice: 11500.0, markup: 40.0, lowStock: 5 },
  { id: "p1272", name: "POWER PLUG 9 IN ONE", category: "General", unit: "piece", quantity: 0, costPrice: 270.0, markup: 30.0, lowStock: 5 },
  { id: "p1273", name: "POWER PLUG EXT", category: "General", unit: "piece", quantity: 0, costPrice: 150.0, markup: 40.0, lowStock: 5 },
  { id: "p1274", name: "PPR CUTTER", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 590.0, markup: 35.0, lowStock: 5 },
  { id: "p1275", name: "PPR GATE VALVE 32MM LOCAL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 360.0, markup: 30.0, lowStock: 5 },
  { id: "p1276", name: "PPR HEATER MOULD GUTKA   32MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 210.0, markup: 40.0, lowStock: 5 },
  { id: "p1277", name: "PPR HEATER MOULD GUTKA   40MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 240.0, markup: 40.0, lowStock: 5 },
  { id: "p1278", name: "PPR SADDLE 25MM", category: "Electrical", unit: "packet", quantity: 0, costPrice: 520.0, markup: 35.0, lowStock: 5 },
  { id: "p1279", name: "PPR SADDLE 32MM", category: "Electrical", unit: "packet", quantity: 0, costPrice: 620.0, markup: 35.0, lowStock: 5 },
  { id: "p1280", name: "PPR TEE 40X1-1/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1.0, markup: 40.0, lowStock: 5 },
  { id: "p1281", name: "PPR UNION 32MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 90.0, markup: 35.0, lowStock: 5 },
  { id: "p1282", name: "PRESSURE SWITCH", category: "Electrical", unit: "piece", quantity: 0, costPrice: 4200.0, markup: 40.0, lowStock: 5 },
  { id: "p1283", name: "PRIDE PLIER BLACK", category: "General", unit: "piece", quantity: 0, costPrice: 630.0, markup: 40.0, lowStock: 5 },
  { id: "p1284", name: "PU FOAM SPRAY BIG", category: "General", unit: "piece", quantity: 0, costPrice: 900.0, markup: 30.0, lowStock: 5 },
  { id: "p1285", name: "PU FOAM SPRAY SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 650.0, markup: 30.0, lowStock: 5 },
  { id: "p1286", name: "PUMP NOZAL", category: "General", unit: "piece", quantity: 0, costPrice: 20.76, markup: 40.0, lowStock: 5 },
  { id: "p1287", name: "PUMP NOZAL 1\"", category: "General", unit: "piece", quantity: 0, costPrice: 25.0, markup: 30.0, lowStock: 5 },
  { id: "p1288", name: "PUMP NOZAL 1X3/4", category: "General", unit: "piece", quantity: 0, costPrice: 31.67, markup: 40.0, lowStock: 5 },
  { id: "p1289", name: "PUMP NOZAL 3/4", category: "General", unit: "piece", quantity: 0, costPrice: 23.33, markup: 30.0, lowStock: 5 },
  { id: "p1290", name: "PUMP NOZAL 3/4\"", category: "General", unit: "piece", quantity: 0, costPrice: 24.17, markup: 40.0, lowStock: 5 },
  { id: "p1291", name: "PUMP NOZZAL 1\"", category: "General", unit: "piece", quantity: 0, costPrice: 30.0, markup: 40.0, lowStock: 5 },
  { id: "p1292", name: "PUMP NOZZAL 3/4\"", category: "General", unit: "piece", quantity: 0, costPrice: 22.08, markup: 40.0, lowStock: 5 },
  { id: "p1293", name: "PUSH BUTTON SINGLE BIG", category: "General", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p1294", name: "PUSH BUTTON SINGLE SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 110.0, markup: 40.0, lowStock: 5 },
  { id: "p1295", name: "PVC  RAWAL PLUG", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 35.0, markup: 35.0, lowStock: 5 },
  { id: "p1296", name: "PVC BIB COCK", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 80.0, markup: 30.0, lowStock: 5 },
  { id: "p1297", name: "PVC BIB COCK DANY DAR", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p1298", name: "PVC BIB COCK DANYDAAR", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 85.0, markup: 35.0, lowStock: 5 },
  { id: "p1299", name: "PVC BIB COCK TOOTI DANYDAAR", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 75.0, markup: 40.0, lowStock: 5 },
  { id: "p1300", name: "PVC BIB COCK WHITE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 87.5, markup: 35.0, lowStock: 5 },
  { id: "p1301", name: "PVC BIB COK NOZZALWALA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 65.0, markup: 40.0, lowStock: 5 },
  { id: "p1302", name: "PVC BOX 3X3", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 38.0, markup: 40.0, lowStock: 5 },
  { id: "p1303", name: "PVC CHECK VALVE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p1304", name: "PVC CONNECTER  STRIP 60AMP", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 215.0, markup: 40.0, lowStock: 5 },
  { id: "p1305", name: "PVC CONNECTER  STRIP30AMP", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 130.0, markup: 40.0, lowStock: 5 },
  { id: "p1306", name: "PVC CONNECTER STRIP 100AMP", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 355.0, markup: 40.0, lowStock: 5 },
  { id: "p1307", name: "PVC CONNECTER STRIP 15AMP", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p1308", name: "PVC CONNECTION PIPE 24\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 140.0, markup: 40.0, lowStock: 5 },
  { id: "p1309", name: "PVC CONNECTION PIPE 36\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 150.0, markup: 40.0, lowStock: 5 },
  { id: "p1310", name: "PVC CONNECTOR 100AMP", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 2500.0, markup: 35.0, lowStock: 5 },
  { id: "p1311", name: "PVC CONNECTOR 100AMP (05STRIP)", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 390.0, markup: 30.0, lowStock: 5 },
  { id: "p1312", name: "PVC CONNECTOR 10AMP", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 3000.0, markup: 35.0, lowStock: 5 },
  { id: "p1313", name: "PVC CONNECTOR 10AMP   (50 STRIP)", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 50.0, markup: 30.0, lowStock: 5 },
  { id: "p1314", name: "PVC CONNECTOR 30AMP   (10 STRIP)", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 155.0, markup: 30.0, lowStock: 5 },
  { id: "p1315", name: "PVC CONNECTOR 60AMP   (10 STRIP)", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 185.0, markup: 30.0, lowStock: 5 },
  { id: "p1316", name: "PVC CONNECTOR 6AMP", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 3200.0, markup: 35.0, lowStock: 5 },
  { id: "p1317", name: "PVC CONNETOR STRIP 100AMP", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 430.0, markup: 40.0, lowStock: 5 },
  { id: "p1318", name: "PVC CONNETOR STRIP 20AMP", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 75.0, markup: 40.0, lowStock: 5 },
  { id: "p1319", name: "PVC CONNETOR STRIP 30AMP", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 190.0, markup: 40.0, lowStock: 5 },
  { id: "p1320", name: "PVC CONNETOR STRIP 60AMP", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p1321", name: "PVC DABRA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 300.0, markup: 40.0, lowStock: 5 },
  { id: "p1322", name: "PVC ELBOW 2\" BURJ", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 56.0, markup: 40.0, lowStock: 5 },
  { id: "p1323", name: "PVC ELBOW 4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 201.0, markup: 35.0, lowStock: 5 },
  { id: "p1324", name: "PVC ELBOW 4\" BURJ", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 177.0, markup: 40.0, lowStock: 5 },
  { id: "p1325", name: "PVC ELBOW 4X45", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 182.0, markup: 40.0, lowStock: 5 },
  { id: "p1326", name: "PVC GLOVES FULL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 85.0, markup: 40.0, lowStock: 5 },
  { id: "p1327", name: "PVC GURMALA  BIG", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 90.0, markup: 30.0, lowStock: 5 },
  { id: "p1328", name: "PVC GURMALA BIG", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p1329", name: "PVC GURMALA SMALL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 70.0, markup: 30.0, lowStock: 5 },
  { id: "p1330", name: "PVC HAMMER  BIG", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 490.0, markup: 40.0, lowStock: 5 },
  { id: "p1331", name: "PVC HAMMER  MEDIUM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 390.0, markup: 40.0, lowStock: 5 },
  { id: "p1332", name: "PVC HAMMER  SMALL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 290.0, markup: 40.0, lowStock: 5 },
  { id: "p1333", name: "PVC JALI", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 88.0, markup: 40.0, lowStock: 5 },
  { id: "p1334", name: "PVC JALI 6X6", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 110.0, markup: 30.0, lowStock: 5 },
  { id: "p1335", name: "PVC MANHOLE 24X24 PLUS", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 3450.0, markup: 35.0, lowStock: 5 },
  { id: "p1336", name: "PVC MANHOLE COVER SONEX 12X12", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 370.0, markup: 40.0, lowStock: 5 },
  { id: "p1337", name: "PVC MANHOLE COVER SONEX 9X9", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 275.0, markup: 40.0, lowStock: 5 },
  { id: "p1338", name: "PVC PIPE    4KG  3\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 940.0, markup: 40.0, lowStock: 5 },
  { id: "p1339", name: "PVC PIPE    6KG  4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1410.0, markup: 40.0, lowStock: 5 },
  { id: "p1340", name: "PVC PIPE    6KG  5\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1410.0, markup: 40.0, lowStock: 5 },
  { id: "p1341", name: "PVC PIPE    8KG  6\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1880.0, markup: 40.0, lowStock: 5 },
  { id: "p1342", name: "PVC PIPE 2.5KG  2\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 588.0, markup: 40.0, lowStock: 5 },
  { id: "p1343", name: "PVC PIPE B CLASS 3", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 840.0, markup: 40.0, lowStock: 5 },
  { id: "p1344", name: "PVC PIPE B CLASS 4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1120.0, markup: 40.0, lowStock: 5 },
  { id: "p1345", name: "PVC PIPE D CLASS 2\"        280/KG", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 560.0, markup: 40.0, lowStock: 5 },
  { id: "p1346", name: "PVC PIPE GOLD 2\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 760.0, markup: 30.0, lowStock: 5 },
  { id: "p1347", name: "PVC PIPE GOLD 3\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1140.0, markup: 30.0, lowStock: 5 },
  { id: "p1348", name: "PVC PIPE GOLD 4 \"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1520.0, markup: 30.0, lowStock: 5 },
  { id: "p1349", name: "PVC PIPE MASTER B 2\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 650.0, markup: 30.0, lowStock: 5 },
  { id: "p1350", name: "PVC PIPE MASTER B 3\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1040.0, markup: 30.0, lowStock: 5 },
  { id: "p1351", name: "PVC PIPE MASTER B 4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1560.0, markup: 30.0, lowStock: 5 },
  { id: "p1352", name: "PVC PIPE MASTER B 5\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 2080.0, markup: 30.0, lowStock: 5 },
  { id: "p1353", name: "PVC PIPE MASTER B 6\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 2600.0, markup: 30.0, lowStock: 5 },
  { id: "p1354", name: "PVC PIPE MASTER GOLD 3\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1040.0, markup: 40.0, lowStock: 5 },
  { id: "p1355", name: "PVC PIPE MASTER GOLD 4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1275.0, markup: 40.0, lowStock: 5 },
  { id: "p1356", name: "PVC PIPE MASTER GOLD 5\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1685.0, markup: 40.0, lowStock: 5 },
  { id: "p1357", name: "PVC RAWAL PLUG", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 27.08, markup: 30.0, lowStock: 5 },
  { id: "p1358", name: "PVC RAWAL PLUG  GRAY", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 11.67, markup: 40.0, lowStock: 5 },
  { id: "p1359", name: "PVC RAWAL PLUG 12#", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 14.58, markup: 40.0, lowStock: 5 },
  { id: "p1360", name: "PVC RAWAL PLUG 14#", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 380.0, markup: 40.0, lowStock: 5 },
  { id: "p1361", name: "PVC RAWAL PLUG 18NO", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1.94, markup: 40.0, lowStock: 5 },
  { id: "p1362", name: "PVC RAWAL PLUG BLUE PACKET", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 60.0, markup: 40.0, lowStock: 5 },
  { id: "p1363", name: "PVC RAWAL PLUG WHITE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 14.58, markup: 40.0, lowStock: 5 },
  { id: "p1364", name: "PVC SOCKET 3", category: "Electrical", unit: "piece", quantity: 0, costPrice: 110.0, markup: 30.0, lowStock: 5 },
  { id: "p1365", name: "PVC SOCKET 4", category: "Electrical", unit: "piece", quantity: 0, costPrice: 165.0, markup: 30.0, lowStock: 5 },
  { id: "p1366", name: "PVC SOCKET 4\"", category: "Electrical", unit: "piece", quantity: 0, costPrice: 161.0, markup: 35.0, lowStock: 5 },
  { id: "p1367", name: "PVC SOCKET FOR GARDEN PIPE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 380.0, markup: 40.0, lowStock: 5 },
  { id: "p1368", name: "PVC TOOTI WHITE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 150.0, markup: 35.0, lowStock: 5 },
  { id: "p1369", name: "QABZA DOOR HINGES 3\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 48.0, markup: 40.0, lowStock: 5 },
  { id: "p1370", name: "QABZA DOOR HINGES 4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 68.0, markup: 40.0, lowStock: 5 },
  { id: "p1371", name: "QABZA HINGES 3\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 33.33, markup: 40.0, lowStock: 5 },
  { id: "p1372", name: "QALAM", category: "General", unit: "piece", quantity: 0, costPrice: 0.0, markup: 40.0, lowStock: 5 },
  { id: "p1373", name: "QALAM BRUSH 7CS WHITE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 230.0, markup: 30.0, lowStock: 5 },
  { id: "p1374", name: "QALAM BRUSH RED 12PCS", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 210.0, markup: 30.0, lowStock: 5 },
  { id: "p1375", name: "QALAM SET", category: "General", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p1376", name: "RAIGMAR 12", category: "General", unit: "piece", quantity: 0, costPrice: 35.0, markup: 40.0, lowStock: 5 },
  { id: "p1377", name: "RAMBBA", category: "General", unit: "piece", quantity: 0, costPrice: 140.0, markup: 40.0, lowStock: 5 },
  { id: "p1378", name: "RAT BOOK", category: "General", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p1379", name: "RAW CENA", category: "General", unit: "piece", quantity: 0, costPrice: 60.0, markup: 40.0, lowStock: 5 },
  { id: "p1380", name: "RAWA; PLUG", category: "General", unit: "piece", quantity: 0, costPrice: 20.0, markup: 30.0, lowStock: 5 },
  { id: "p1381", name: "RAWAL BOLT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 17.0, markup: 40.0, lowStock: 5 },
  { id: "p1382", name: "RAWAL BOLT 10MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 27.0, markup: 30.0, lowStock: 5 },
  { id: "p1383", name: "RAWAL BOLT 1OMM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 34.5, markup: 40.0, lowStock: 5 },
  { id: "p1384", name: "RAWAL BOLT 3 SOOTAR", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 16.67, markup: 40.0, lowStock: 5 },
  { id: "p1385", name: "RAWAL BOLT 5/16", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 19.17, markup: 40.0, lowStock: 5 },
  { id: "p1386", name: "RAWAL BOLT 8MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 23.0, markup: 30.0, lowStock: 5 },
  { id: "p1387", name: "RAWAL BOLT KIT", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 75.0, markup: 35.0, lowStock: 5 },
  { id: "p1388", name: "RAWAL PLUG", category: "General", unit: "piece", quantity: 0, costPrice: 13.33, markup: 40.0, lowStock: 5 },
  { id: "p1389", name: "RED OXIDE BIG", category: "General", unit: "piece", quantity: 0, costPrice: 193.0, markup: 30.0, lowStock: 5 },
  { id: "p1390", name: "REDOXIDE PAINT", category: "Paint", unit: "qrt", quantity: 0, costPrice: 240.0, markup: 40.0, lowStock: 5 },
  { id: "p1391", name: "REINFORCE HOSE PIPE GARDEN PIPE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 630.0, markup: 40.0, lowStock: 5 },
  { id: "p1392", name: "RIGHT ANGLE GUNYA 12\"", category: "General", unit: "piece", quantity: 0, costPrice: 260.0, markup: 40.0, lowStock: 5 },
  { id: "p1393", name: "RIGHT ANGLE GUNYA 24\"", category: "General", unit: "piece", quantity: 0, costPrice: 415.0, markup: 40.0, lowStock: 5 },
  { id: "p1394", name: "RIVOT GUN", category: "General", unit: "piece", quantity: 0, costPrice: 480.0, markup: 40.0, lowStock: 5 },
  { id: "p1395", name: "ROBICON SCRAPER", category: "General", unit: "piece", quantity: 0, costPrice: 52.0, markup: 40.0, lowStock: 5 },
  { id: "p1396", name: "ROPE LIGHT", category: "General", unit: "piece", quantity: 0, costPrice: 8500.0, markup: 40.0, lowStock: 5 },
  { id: "p1397", name: "ROPE LIGHT 100MTR", category: "General", unit: "piece", quantity: 0, costPrice: 8750.0, markup: 40.0, lowStock: 5 },
  { id: "p1398", name: "ROPE LIGHT 70MTR", category: "General", unit: "roll", quantity: 0, costPrice: 6700.0, markup: 40.0, lowStock: 5 },
  { id: "p1399", name: "ROPE LIGHT ADAPTER", category: "General", unit: "piece", quantity: 0, costPrice: 110.0, markup: 35.0, lowStock: 5 },
  { id: "p1400", name: "ROPE LIGHT CONNECTOR STRIP", category: "General", unit: "piece", quantity: 0, costPrice: 150.0, markup: 40.0, lowStock: 5 },
  { id: "p1401", name: "ROUND EXHAUST FAN", category: "General", unit: "piece", quantity: 0, costPrice: 2400.0, markup: 40.0, lowStock: 5 },
  { id: "p1402", name: "ROUND SAND PAKER STICKY 4\"", category: "General", unit: "piece", quantity: 0, costPrice: 13.0, markup: 30.0, lowStock: 5 },
  { id: "p1403", name: "RUBBER GLOVES LONG", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p1404", name: "RUBBER HAMMER GREEN", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 380.0, markup: 30.0, lowStock: 5 },
  { id: "p1405", name: "RUBBER HANGING CLAM    1\"", category: "General", unit: "piece", quantity: 0, costPrice: 85.0, markup: 30.0, lowStock: 5 },
  { id: "p1406", name: "RUBBER HANGING CLAM 1-1/4\"", category: "General", unit: "piece", quantity: 0, costPrice: 95.0, markup: 30.0, lowStock: 5 },
  { id: "p1407", name: "RUBBER HANGING CLAM 3/4\"", category: "General", unit: "piece", quantity: 0, costPrice: 75.0, markup: 30.0, lowStock: 5 },
  { id: "p1408", name: "RUBBER HANGING CLAMP", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p1409", name: "RUBBER HANGING CLAMP 3\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 170.0, markup: 40.0, lowStock: 5 },
  { id: "p1410", name: "RUBBER HANGING CLAMP 4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 165.0, markup: 40.0, lowStock: 5 },
  { id: "p1411", name: "RUBBER HANGING CLAMP 6\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 225.0, markup: 40.0, lowStock: 5 },
  { id: "p1412", name: "RUBBER WASHER", category: "General", unit: "packet", quantity: 0, costPrice: 150.0, markup: 30.0, lowStock: 5 },
  { id: "p1413", name: "RUBBER WASHERS", category: "General", unit: "piece", quantity: 0, costPrice: 65.0, markup: 40.0, lowStock: 5 },
  { id: "p1414", name: "RUBICON SCRAPER", category: "General", unit: "piece", quantity: 0, costPrice: 45.0, markup: 40.0, lowStock: 5 },
  { id: "p1415", name: "RUBICON SCRAPER 6\"", category: "General", unit: "piece", quantity: 0, costPrice: 50.0, markup: 40.0, lowStock: 5 },
  { id: "p1416", name: "SAAL PLUMBOB BULBUL STEEL", category: "Electrical", unit: "piece", quantity: 0, costPrice: 186.0, markup: 40.0, lowStock: 5 },
  { id: "p1417", name: "SADDLE  10MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p1418", name: "SADDLE  6MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 110.0, markup: 40.0, lowStock: 5 },
  { id: "p1419", name: "SADDLE  8MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p1420", name: "SADDLE 3/4    100 PCS", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 3.9, markup: 40.0, lowStock: 5 },
  { id: "p1421", name: "SADDLE 3/4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 4.69, markup: 40.0, lowStock: 5 },
  { id: "p1422", name: "SADDLE 4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p1423", name: "SADDLE 5", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p1424", name: "SADDLE 5NO   96 PCS", category: "Plumbing & Sanitary", unit: "pkts", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p1425", name: "SADDLE 6 NO  96 PCS", category: "Plumbing & Sanitary", unit: "pkts", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p1426", name: "SADDLE CLAMP 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 5.4, markup: 40.0, lowStock: 5 },
  { id: "p1427", name: "SADDLE CLIP     4                         160/8= 20", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p1428", name: "SADDLE CLIP     5                         160/8= 20", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p1429", name: "SADDLE CLIP     6                         180/8= 22.5", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p1430", name: "SADDLE CLIP     7                         190/8= 23.75", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 190.0, markup: 40.0, lowStock: 5 },
  { id: "p1431", name: "SADDLE CLIP     8                         200/8= 25", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p1432", name: "SADDLE CLIP     9                         210/8= 26.25", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 210.0, markup: 40.0, lowStock: 5 },
  { id: "p1433", name: "SADDLE CLIP  05", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 130.0, markup: 40.0, lowStock: 5 },
  { id: "p1434", name: "SADDLE CLIP  06", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 130.0, markup: 40.0, lowStock: 5 },
  { id: "p1435", name: "SADDLE CLIP  07", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 140.0, markup: 40.0, lowStock: 5 },
  { id: "p1436", name: "SADDLE CLIP  1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 550.0, markup: 30.0, lowStock: 5 },
  { id: "p1437", name: "SADDLE CLIP  10", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 260.0, markup: 30.0, lowStock: 5 },
  { id: "p1438", name: "SADDLE CLIP  12", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 280.0, markup: 30.0, lowStock: 5 },
  { id: "p1439", name: "SADDLE CLIP  3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 450.0, markup: 30.0, lowStock: 5 },
  { id: "p1440", name: "SADDLE CLIP  5", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 130.0, markup: 30.0, lowStock: 5 },
  { id: "p1441", name: "SADDLE CLIP  6", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 130.0, markup: 30.0, lowStock: 5 },
  { id: "p1442", name: "SADDLE CLIP  7", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 140.0, markup: 30.0, lowStock: 5 },
  { id: "p1443", name: "SADDLE CLIP  8", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 140.0, markup: 30.0, lowStock: 5 },
  { id: "p1444", name: "SADDLE CLIP  9", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 140.0, markup: 30.0, lowStock: 5 },
  { id: "p1445", name: "SADDLE CLIP 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 5.42, markup: 40.0, lowStock: 5 },
  { id: "p1446", name: "SADDLE CLIP 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 3.13, markup: 40.0, lowStock: 5 },
  { id: "p1447", name: "SADDLE WIRE CLIP   10 NO", category: "Electrical", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p1448", name: "SADDLE WIRE CLIP   9 NO", category: "Electrical", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p1449", name: "SADDLE WIRE CLIP  1\"", category: "Electrical", unit: "piece", quantity: 0, costPrice: 480.0, markup: 40.0, lowStock: 5 },
  { id: "p1450", name: "SADDLE WIRE CLIP  12 NO", category: "Electrical", unit: "piece", quantity: 0, costPrice: 240.0, markup: 40.0, lowStock: 5 },
  { id: "p1451", name: "SADDLE WIRE CLIP  3 NO", category: "Electrical", unit: "piece", quantity: 0, costPrice: 110.0, markup: 40.0, lowStock: 5 },
  { id: "p1452", name: "SADDLE WIRE CLIP  3/4\"", category: "Electrical", unit: "piece", quantity: 0, costPrice: 360.0, markup: 40.0, lowStock: 5 },
  { id: "p1453", name: "SAFAIDA", category: "General", unit: "piece", quantity: 0, costPrice: 46.67, markup: 30.0, lowStock: 5 },
  { id: "p1454", name: "SAFETY VALVE 1", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 820.0, markup: 40.0, lowStock: 5 },
  { id: "p1455", name: "SAFETY VALVE 1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 640.0, markup: 40.0, lowStock: 5 },
  { id: "p1456", name: "SAMAD BACHAT BIG", category: "Paint", unit: "piece", quantity: 0, costPrice: 660.0, markup: 35.0, lowStock: 5 },
  { id: "p1457", name: "SAMAD BACHAT SMALL", category: "Paint", unit: "piece", quantity: 0, costPrice: 230.0, markup: 35.0, lowStock: 5 },
  { id: "p1458", name: "SAMAD BOND", category: "Paint", unit: "tin", quantity: 0, costPrice: 185.0, markup: 40.0, lowStock: 5 },
  { id: "p1459", name: "SAMAD BOND JIMSA", category: "Paint", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p1460", name: "SAMD BOND TIN", category: "Cement & Aggregates", unit: "piece", quantity: 0, costPrice: 370.0, markup: 40.0, lowStock: 5 },
  { id: "p1461", name: "SAMD BOND TUBE", category: "Cement & Aggregates", unit: "piece", quantity: 0, costPrice: 52.0, markup: 40.0, lowStock: 5 },
  { id: "p1462", name: "SAND PAER DISC 4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 75.0, markup: 40.0, lowStock: 5 },
  { id: "p1463", name: "SAND PAER DISC 5\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 100.0, markup: 40.0, lowStock: 5 },
  { id: "p1464", name: "SAND PAER DISC 7\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 150.0, markup: 40.0, lowStock: 5 },
  { id: "p1465", name: "SAND PAPER", category: "Cement & Aggregates", unit: "ft", quantity: 0, costPrice: 31.0, markup: 30.0, lowStock: 5 },
  { id: "p1466", name: "SAND PAPER #60", category: "Cement & Aggregates", unit: "piece", quantity: 0, costPrice: 15.0, markup: 40.0, lowStock: 5 },
  { id: "p1467", name: "SAND PAPER 120 AND 150", category: "Cement & Aggregates", unit: "piece", quantity: 0, costPrice: 33.0, markup: 40.0, lowStock: 5 },
  { id: "p1468", name: "SAND PAPER DISC", category: "Cement & Aggregates", unit: "piece", quantity: 0, costPrice: 150.0, markup: 40.0, lowStock: 5 },
  { id: "p1469", name: "SAND PAPER KOREAN", category: "Cement & Aggregates", unit: "piece", quantity: 0, costPrice: 85.0, markup: 40.0, lowStock: 5 },
  { id: "p1470", name: "SAND PAPER KORIAN", category: "Cement & Aggregates", unit: "piece", quantity: 0, costPrice: 90.0, markup: 30.0, lowStock: 5 },
  { id: "p1471", name: "SAND PAPER PAD  4\"", category: "Cement & Aggregates", unit: "piece", quantity: 0, costPrice: 97.0, markup: 30.0, lowStock: 5 },
  { id: "p1472", name: "SAND PAPER PAD 5\"", category: "Cement & Aggregates", unit: "piece", quantity: 0, costPrice: 137.0, markup: 30.0, lowStock: 5 },
  { id: "p1473", name: "SAND PAPER RAIG MAAR 120", category: "Cement & Aggregates", unit: "piece", quantity: 0, costPrice: 3800.0, markup: 35.0, lowStock: 5 },
  { id: "p1474", name: "SCRAPER ORDINORY", category: "General", unit: "piece", quantity: 0, costPrice: 35.83, markup: 40.0, lowStock: 5 },
  { id: "p1475", name: "SCRAPER RUBICONE", category: "General", unit: "piece", quantity: 0, costPrice: 45.0, markup: 35.0, lowStock: 5 },
  { id: "p1476", name: "SCREW", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 2.0, markup: 40.0, lowStock: 5 },
  { id: "p1477", name: "SCREW BIT SOUBLE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 75.0, markup: 30.0, lowStock: 5 },
  { id: "p1478", name: "SCREW DRIVER 4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 145.0, markup: 30.0, lowStock: 5 },
  { id: "p1479", name: "SCREW DRIVER 4\"BLACK ORANGE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 150.0, markup: 35.0, lowStock: 5 },
  { id: "p1480", name: "SCREW DRIVER 4X6\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 150.0, markup: 40.0, lowStock: 5 },
  { id: "p1481", name: "SCREW DRIVER 6X6\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p1482", name: "SCREW DRIVER AMERICAN", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 115.0, markup: 40.0, lowStock: 5 },
  { id: "p1483", name: "SCREW DRIVER BABY", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 60.0, markup: 40.0, lowStock: 5 },
  { id: "p1484", name: "SCREW DRIVER BLACK", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 140.0, markup: 30.0, lowStock: 5 },
  { id: "p1485", name: "SCREW DRIVER BLUE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 125.0, markup: 30.0, lowStock: 5 },
  { id: "p1486", name: "SCREW DRIVER RED 6\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 175.0, markup: 35.0, lowStock: 5 },
  { id: "p1487", name: "SCREW DRIVER ST SMALL", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 230.0, markup: 40.0, lowStock: 5 },
  { id: "p1488", name: "SCREW DRIVER USA", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p1489", name: "SCREW DRIVER WHITE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 55.0, markup: 40.0, lowStock: 5 },
  { id: "p1490", name: "SCREW WRENCH 10\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 530.0, markup: 40.0, lowStock: 5 },
  { id: "p1491", name: "SCREW WRENCH 6\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 320.0, markup: 40.0, lowStock: 5 },
  { id: "p1492", name: "SCREW WRENCH 8\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 420.0, markup: 40.0, lowStock: 5 },
  { id: "p1493", name: "SEEBA DOORI", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 45.38, markup: 40.0, lowStock: 5 },
  { id: "p1494", name: "SHAVOL 4", category: "General", unit: "piece", quantity: 0, costPrice: 450.0, markup: 40.0, lowStock: 5 },
  { id: "p1495", name: "SHAVOL 5", category: "General", unit: "piece", quantity: 0, costPrice: 467.0, markup: 40.0, lowStock: 5 },
  { id: "p1496", name: "SHEET 3+1", category: "General", unit: "piece", quantity: 0, costPrice: 350.0, markup: 40.0, lowStock: 5 },
  { id: "p1497", name: "SHELTER BOX 3X3", category: "General", unit: "piece", quantity: 0, costPrice: 80.0, markup: 35.0, lowStock: 5 },
  { id: "p1498", name: "SHOWER HEAD", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 188.0, markup: 40.0, lowStock: 5 },
  { id: "p1499", name: "SHOWER MIXER", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 4200.0, markup: 30.0, lowStock: 5 },
  { id: "p1500", name: "SHOWER SET", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 6800.0, markup: 40.0, lowStock: 5 },
  { id: "p1501", name: "SHULTER BOX", category: "General", unit: "piece", quantity: 0, costPrice: 40.0, markup: 40.0, lowStock: 5 },
  { id: "p1502", name: "SHUTTRING HAND SAW", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 185.0, markup: 35.0, lowStock: 5 },
  { id: "p1503", name: "SHUTTRING NAIL", category: "Hardware & Tools", unit: "kg", quantity: 0, costPrice: 320.0, markup: 40.0, lowStock: 5 },
  { id: "p1504", name: "SHUTTRING NAIL 2X12", category: "Hardware & Tools", unit: "kg", quantity: 0, costPrice: 332.0, markup: 30.0, lowStock: 5 },
  { id: "p1505", name: "SHUTTRING NAIL 3X12", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 316.67, markup: 30.0, lowStock: 5 },
  { id: "p1506", name: "SIDE CUTTER PRIDE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 450.0, markup: 40.0, lowStock: 5 },
  { id: "p1507", name: "SIDE PILLER COCK", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1050.0, markup: 40.0, lowStock: 5 },
  { id: "p1508", name: "SIDE PILLER COCK MIXER", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 980.0, markup: 40.0, lowStock: 5 },
  { id: "p1509", name: "SILICON", category: "General", unit: "piece", quantity: 0, costPrice: 380.0, markup: 30.0, lowStock: 5 },
  { id: "p1510", name: "SILICON 20GRM", category: "General", unit: "piece", quantity: 0, costPrice: 55.0, markup: 30.0, lowStock: 5 },
  { id: "p1511", name: "SILICON 50G", category: "General", unit: "piece", quantity: 0, costPrice: 115.0, markup: 30.0, lowStock: 5 },
  { id: "p1512", name: "SILICON BIG", category: "General", unit: "piece", quantity: 0, costPrice: 110.0, markup: 40.0, lowStock: 5 },
  { id: "p1513", name: "SILICON BIG TUBE", category: "General", unit: "piece", quantity: 0, costPrice: 115.0, markup: 35.0, lowStock: 5 },
  { id: "p1514", name: "SILICON BOTEL", category: "General", unit: "piece", quantity: 0, costPrice: 330.0, markup: 30.0, lowStock: 5 },
  { id: "p1515", name: "SILICON GUN", category: "General", unit: "piece", quantity: 0, costPrice: 480.0, markup: 30.0, lowStock: 5 },
  { id: "p1516", name: "SILICON SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 60.0, markup: 40.0, lowStock: 5 },
  { id: "p1517", name: "SILICON SMALL TUBE", category: "General", unit: "piece", quantity: 0, costPrice: 60.0, markup: 35.0, lowStock: 5 },
  { id: "p1518", name: "SILICON TUBE", category: "General", unit: "piece", quantity: 0, costPrice: 350.0, markup: 40.0, lowStock: 5 },
  { id: "p1519", name: "SILICON TUBE BIG", category: "General", unit: "piece", quantity: 0, costPrice: 110.0, markup: 40.0, lowStock: 5 },
  { id: "p1520", name: "SILICON TUBE SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 60.0, markup: 40.0, lowStock: 5 },
  { id: "p1521", name: "SILICONE BIG", category: "General", unit: "piece", quantity: 0, costPrice: 130.0, markup: 40.0, lowStock: 5 },
  { id: "p1522", name: "SILICONE SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 65.0, markup: 40.0, lowStock: 5 },
  { id: "p1523", name: "SILICONE TUBE", category: "General", unit: "piece", quantity: 0, costPrice: 495.0, markup: 40.0, lowStock: 5 },
  { id: "p1524", name: "SILICONE TUBE MEDIUM", category: "General", unit: "piece", quantity: 0, costPrice: 100.0, markup: 40.0, lowStock: 5 },
  { id: "p1525", name: "SILICONE TUBE SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 50.0, markup: 40.0, lowStock: 5 },
  { id: "p1526", name: "SILION BLACK", category: "General", unit: "piece", quantity: 0, costPrice: 440.0, markup: 40.0, lowStock: 5 },
  { id: "p1527", name: "SILION TUBE BOTEL", category: "General", unit: "piece", quantity: 0, costPrice: 370.0, markup: 40.0, lowStock: 5 },
  { id: "p1528", name: "SILION WHITE", category: "General", unit: "piece", quantity: 0, costPrice: 435.0, markup: 40.0, lowStock: 5 },
  { id: "p1529", name: "SILOCON BUN RED", category: "General", unit: "piece", quantity: 0, costPrice: 420.0, markup: 40.0, lowStock: 5 },
  { id: "p1530", name: "SINGLE BIB COCK SONEX", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 2200.0, markup: 40.0, lowStock: 5 },
  { id: "p1531", name: "SINGLE SINK MIXER", category: "General", unit: "piece", quantity: 0, costPrice: 1690.0, markup: 40.0, lowStock: 5 },
  { id: "p1532", name: "SINGLE WASH BASIN MIXER", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1690.0, markup: 40.0, lowStock: 5 },
  { id: "p1533", name: "SINK BOWL 14X17", category: "General", unit: "piece", quantity: 0, costPrice: 1180.0, markup: 40.0, lowStock: 5 },
  { id: "p1534", name: "SINK BOWL 14X17X8", category: "General", unit: "piece", quantity: 0, costPrice: 1900.0, markup: 40.0, lowStock: 5 },
  { id: "p1535", name: "SINK BOWL DOUBLE RASHID 7843", category: "General", unit: "piece", quantity: 0, costPrice: 17000.0, markup: 40.0, lowStock: 5 },
  { id: "p1536", name: "SINK MIXER CHINA", category: "General", unit: "piece", quantity: 0, costPrice: 3000.0, markup: 30.0, lowStock: 5 },
  { id: "p1537", name: "SINK MIXER SHORP", category: "General", unit: "piece", quantity: 0, costPrice: 5800.0, markup: 40.0, lowStock: 5 },
  { id: "p1538", name: "SINK MIXER SINGLE", category: "General", unit: "piece", quantity: 0, costPrice: 980.0, markup: 40.0, lowStock: 5 },
  { id: "p1539", name: "SINK MIXERDOUBLE", category: "General", unit: "piece", quantity: 0, costPrice: 1750.0, markup: 40.0, lowStock: 5 },
  { id: "p1540", name: "SJL DISC 4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 490.0, markup: 35.0, lowStock: 5 },
  { id: "p1541", name: "SJL DISC 5\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 590.0, markup: 35.0, lowStock: 5 },
  { id: "p1542", name: "SJL DISC 7\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 1160.0, markup: 35.0, lowStock: 5 },
  { id: "p1543", name: "SMALL SOLUTION", category: "Paint", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p1544", name: "SMD 12W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 210.0, markup: 40.0, lowStock: 5 },
  { id: "p1545", name: "SMD LIGHT", category: "Electrical", unit: "piece", quantity: 0, costPrice: 216.0, markup: 40.0, lowStock: 5 },
  { id: "p1546", name: "SMD LIGHT 12W TUFF", category: "Electrical", unit: "piece", quantity: 0, costPrice: 290.0, markup: 40.0, lowStock: 5 },
  { id: "p1547", name: "SMD LIGHT 7W TUFF", category: "Electrical", unit: "piece", quantity: 0, costPrice: 190.0, markup: 40.0, lowStock: 5 },
  { id: "p1548", name: "SMD LIGHT TUFF 7W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 190.0, markup: 40.0, lowStock: 5 },
  { id: "p1549", name: "SMT DISC 4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 108.15, markup: 40.0, lowStock: 5 },
  { id: "p1550", name: "SOAP HOLDER GOLDEN", category: "Electrical", unit: "piece", quantity: 0, costPrice: 145.0, markup: 40.0, lowStock: 5 },
  { id: "p1551", name: "SOCKET 1\"", category: "Electrical", unit: "piece", quantity: 0, costPrice: 6.6, markup: 40.0, lowStock: 5 },
  { id: "p1552", name: "SOCKET 1\" GM", category: "Electrical", unit: "piece", quantity: 0, costPrice: 11.11, markup: 40.0, lowStock: 5 },
  { id: "p1553", name: "SOCKET 3/4\"", category: "Electrical", unit: "piece", quantity: 0, costPrice: 5.9, markup: 40.0, lowStock: 5 },
  { id: "p1554", name: "SOCKET 32X3/4", category: "Electrical", unit: "piece", quantity: 0, costPrice: 225.4, markup: 40.0, lowStock: 5 },
  { id: "p1555", name: "SOCKET BLACK BIG", category: "Electrical", unit: "piece", quantity: 0, costPrice: 91.67, markup: 40.0, lowStock: 5 },
  { id: "p1556", name: "SOCKET BLACK SMALL", category: "Electrical", unit: "piece", quantity: 0, costPrice: 50.0, markup: 40.0, lowStock: 5 },
  { id: "p1557", name: "SOCKET GM 3/4", category: "Electrical", unit: "piece", quantity: 0, costPrice: 9.17, markup: 40.0, lowStock: 5 },
  { id: "p1558", name: "SOCKETT GASS", category: "Electrical", unit: "piece", quantity: 0, costPrice: 5.83, markup: 40.0, lowStock: 5 },
  { id: "p1559", name: "SOFT BRUSH ANGLE", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 240.0, markup: 40.0, lowStock: 5 },
  { id: "p1560", name: "SOFT BRUSH CLEANING", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p1561", name: "SOLEX LOCK 40MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 460.0, markup: 40.0, lowStock: 5 },
  { id: "p1562", name: "SOLEX LOCK 50MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 510.0, markup: 40.0, lowStock: 5 },
  { id: "p1563", name: "SOLEX LOCK 60MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 550.0, markup: 40.0, lowStock: 5 },
  { id: "p1564", name: "SOLUTION 125ML", category: "Paint", unit: "piece", quantity: 0, costPrice: 165.0, markup: 30.0, lowStock: 5 },
  { id: "p1565", name: "SOLUTION 12KG", category: "Paint", unit: "piece", quantity: 0, costPrice: 420.0, markup: 30.0, lowStock: 5 },
  { id: "p1566", name: "SOLUTION 250ML", category: "Paint", unit: "piece", quantity: 0, costPrice: 265.0, markup: 30.0, lowStock: 5 },
  { id: "p1567", name: "SOLUTION 4OZ", category: "Paint", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p1568", name: "SOLUTION 500ML", category: "Paint", unit: "piece", quantity: 0, costPrice: 420.0, markup: 30.0, lowStock: 5 },
  { id: "p1569", name: "SOLUTION 75G", category: "Paint", unit: "piece", quantity: 0, costPrice: 85.0, markup: 30.0, lowStock: 5 },
  { id: "p1570", name: "SOLUTION BABY", category: "Paint", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p1571", name: "SOLUTION GLUE", category: "Paint", unit: "piece", quantity: 0, costPrice: 420.0, markup: 30.0, lowStock: 5 },
  { id: "p1572", name: "SOLUTION GLUE 1/2KG", category: "Paint", unit: "piece", quantity: 0, costPrice: 350.0, markup: 40.0, lowStock: 5 },
  { id: "p1573", name: "SOLUTION GLUE AR MASTER  129ML", category: "Paint", unit: "piece", quantity: 0, costPrice: 165.0, markup: 30.0, lowStock: 5 },
  { id: "p1574", name: "SOLUTION GLUE AR MASTER  250ML", category: "Paint", unit: "piece", quantity: 0, costPrice: 280.0, markup: 30.0, lowStock: 5 },
  { id: "p1575", name: "SOLUTION GLUE AR MASTER  500ML", category: "Paint", unit: "piece", quantity: 0, costPrice: 420.0, markup: 30.0, lowStock: 5 },
  { id: "p1576", name: "SOLUTION GLUE AR MASTER  BABY", category: "Paint", unit: "piece", quantity: 0, costPrice: 85.0, markup: 30.0, lowStock: 5 },
  { id: "p1577", name: "SOLUTION GLUE BABY", category: "Paint", unit: "piece", quantity: 0, costPrice: 75.0, markup: 40.0, lowStock: 5 },
  { id: "p1578", name: "SOLUTION GLUE BIG", category: "Paint", unit: "piece", quantity: 0, costPrice: 350.0, markup: 40.0, lowStock: 5 },
  { id: "p1579", name: "SOLUTION GLUE MEDIUM", category: "Paint", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p1580", name: "SOLUTION GLUE SMALL", category: "Paint", unit: "piece", quantity: 0, costPrice: 120.0, markup: 40.0, lowStock: 5 },
  { id: "p1581", name: "SOOTAR GOLA", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 35.0, markup: 35.0, lowStock: 5 },
  { id: "p1582", name: "SOOTAR GOLA HEAVY DUTY", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 40.0, markup: 40.0, lowStock: 5 },
  { id: "p1583", name: "SPACER", category: "General", unit: "piece", quantity: 0, costPrice: 25.0, markup: 30.0, lowStock: 5 },
  { id: "p1584", name: "SPANDEL", category: "General", unit: "piece", quantity: 0, costPrice: 130.0, markup: 40.0, lowStock: 5 },
  { id: "p1585", name: "SPANDLE", category: "General", unit: "piece", quantity: 0, costPrice: 118.0, markup: 40.0, lowStock: 5 },
  { id: "p1586", name: "SPANDLE PART #4", category: "General", unit: "piece", quantity: 0, costPrice: 60.0, markup: 35.0, lowStock: 5 },
  { id: "p1587", name: "SPANDLE PART #5", category: "General", unit: "piece", quantity: 0, costPrice: 70.0, markup: 35.0, lowStock: 5 },
  { id: "p1588", name: "SPANDLE PART #6", category: "General", unit: "piece", quantity: 0, costPrice: 75.0, markup: 35.0, lowStock: 5 },
  { id: "p1589", name: "SPANLE", category: "General", unit: "piece", quantity: 0, costPrice: 160.0, markup: 35.0, lowStock: 5 },
  { id: "p1590", name: "SPIRIT", category: "General", unit: "ltr", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p1591", name: "SPIRIT 1LTR", category: "General", unit: "piece", quantity: 0, costPrice: 165.0, markup: 30.0, lowStock: 5 },
  { id: "p1592", name: "SPRANDLE", category: "General", unit: "piece", quantity: 0, costPrice: 120.0, markup: 30.0, lowStock: 5 },
  { id: "p1593", name: "SPRAY BOTTLE", category: "General", unit: "piece", quantity: 0, costPrice: 120.0, markup: 30.0, lowStock: 5 },
  { id: "p1594", name: "SPRAY HEAD", category: "General", unit: "piece", quantity: 0, costPrice: 65.0, markup: 30.0, lowStock: 5 },
  { id: "p1595", name: "SPRAY PAINT", category: "Paint", unit: "piece", quantity: 0, costPrice: 300.0, markup: 40.0, lowStock: 5 },
  { id: "p1596", name: "SPRIT LEVEL 1FT ANP", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 400.0, markup: 40.0, lowStock: 5 },
  { id: "p1597", name: "SQUARE EXHAUST FAN", category: "General", unit: "piece", quantity: 0, costPrice: 4000.0, markup: 40.0, lowStock: 5 },
  { id: "p1598", name: "STEEL AIL1\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 300.0, markup: 30.0, lowStock: 5 },
  { id: "p1599", name: "STEEL BIB COCK", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 575.0, markup: 40.0, lowStock: 5 },
  { id: "p1600", name: "STEEL CUTTING DISC 14\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 490.0, markup: 40.0, lowStock: 5 },
  { id: "p1601", name: "STEEL CUTTING DISC 4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 25.0, markup: 40.0, lowStock: 5 },
  { id: "p1602", name: "STEEL CUTTING DISC 4\" SMT", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 105.0, markup: 40.0, lowStock: 5 },
  { id: "p1603", name: "STEEL CUTTING DISC 4\" THIN", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 25.0, markup: 40.0, lowStock: 5 },
  { id: "p1604", name: "STEEL CUTTING DISC 4' THIN", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 25.0, markup: 40.0, lowStock: 5 },
  { id: "p1605", name: "STEEL CUTTING DISC 5\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 45.0, markup: 40.0, lowStock: 5 },
  { id: "p1606", name: "STEEL CUTTING DISC 5\"  1MM  OSKAR", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p1607", name: "STEEL CUTTING DISC 5\" SMT", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 155.0, markup: 40.0, lowStock: 5 },
  { id: "p1608", name: "STEEL CUTTING DISC 5\" THIN", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 50.0, markup: 40.0, lowStock: 5 },
  { id: "p1609", name: "STEEL CUTTING DISC 5\"X1MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 42.0, markup: 30.0, lowStock: 5 },
  { id: "p1610", name: "STEEL CUTTING DISC 7\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 215.0, markup: 40.0, lowStock: 5 },
  { id: "p1611", name: "STEEL CUTTING DISC 7\" AKI", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 210.0, markup: 40.0, lowStock: 5 },
  { id: "p1612", name: "STEEL CUTTING DISC 7\" SMT", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 250.0, markup: 40.0, lowStock: 5 },
  { id: "p1613", name: "STEEL CUTTING DISC 9 AKI", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 330.0, markup: 40.0, lowStock: 5 },
  { id: "p1614", name: "STEEL CUTTING DISC 9\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 300.0, markup: 35.0, lowStock: 5 },
  { id: "p1615", name: "STEEL CUTTING DISC 9\" AKI", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 340.0, markup: 30.0, lowStock: 5 },
  { id: "p1616", name: "STEEL CUTTING DISC 9\" AKT", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 270.0, markup: 30.0, lowStock: 5 },
  { id: "p1617", name: "STEEL CUTTING DISC GREEN 4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 25.0, markup: 40.0, lowStock: 5 },
  { id: "p1618", name: "STEEL CUTTING DISC OSKAR 1MMX4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 45.0, markup: 30.0, lowStock: 5 },
  { id: "p1619", name: "STEEL CUTTING DISC SS. 1MMMX5\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 60.0, markup: 30.0, lowStock: 5 },
  { id: "p1620", name: "STEEL GRINDING DISC 4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p1621", name: "STEEL GRINDING DISC 5\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 100.0, markup: 40.0, lowStock: 5 },
  { id: "p1622", name: "STEEL GURMALA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 200.0, markup: 40.0, lowStock: 5 },
  { id: "p1623", name: "STEEL GURMALA CUT WALA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 175.0, markup: 40.0, lowStock: 5 },
  { id: "p1624", name: "STEEL GURMALA EAGLE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 212.5, markup: 40.0, lowStock: 5 },
  { id: "p1625", name: "STEEL GURMALA GOLDEN", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 216.67, markup: 30.0, lowStock: 5 },
  { id: "p1626", name: "STEEL NAIL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 300.0, markup: 35.0, lowStock: 5 },
  { id: "p1627", name: "STEEL NAIL 1X1MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 350.0, markup: 30.0, lowStock: 5 },
  { id: "p1628", name: "STEEL NAIL 3/4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 380.0, markup: 40.0, lowStock: 5 },
  { id: "p1629", name: "STEEL NAIL 4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 300.0, markup: 30.0, lowStock: 5 },
  { id: "p1630", name: "STEEL NAIL 6\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 350.0, markup: 30.0, lowStock: 5 },
  { id: "p1631", name: "STEEL NAIL BAREEK", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 365.0, markup: 35.0, lowStock: 5 },
  { id: "p1632", name: "STEEL NAIL CHINA ORIGNAL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 330.0, markup: 40.0, lowStock: 5 },
  { id: "p1633", name: "STEEL NAIL THIN", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 360.0, markup: 40.0, lowStock: 5 },
  { id: "p1634", name: "STEEL NUT CONNECTION PIPE 24\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 100.0, markup: 40.0, lowStock: 5 },
  { id: "p1635", name: "STEEL NUT CONNECTION PIPE 36\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 110.0, markup: 40.0, lowStock: 5 },
  { id: "p1636", name: "STEEL PUTIN", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 200.0, markup: 30.0, lowStock: 5 },
  { id: "p1637", name: "STEEL UTTING DISC 9\" SMT", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 350.0, markup: 40.0, lowStock: 5 },
  { id: "p1638", name: "STEEL WAIST", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 165.0, markup: 40.0, lowStock: 5 },
  { id: "p1639", name: "STEEL WAIST FOR ABOVE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p1640", name: "STEEL WIRE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 580.0, markup: 40.0, lowStock: 5 },
  { id: "p1641", name: "STEEL WIRE 1/2 KG", category: "Electrical", unit: "piece", quantity: 0, costPrice: 500.0, markup: 40.0, lowStock: 5 },
  { id: "p1642", name: "STEEL WIRE 15MTR NEW", category: "Electrical", unit: "piece", quantity: 0, costPrice: 225.0, markup: 40.0, lowStock: 5 },
  { id: "p1643", name: "STEEL WIRE BRUSH", category: "Electrical", unit: "piece", quantity: 0, costPrice: 70.0, markup: 40.0, lowStock: 5 },
  { id: "p1644", name: "STEEL WIRE PAO", category: "Electrical", unit: "kg", quantity: 0, costPrice: 500.0, markup: 40.0, lowStock: 5 },
  { id: "p1645", name: "STEELL CUTTING DISC SS 1MMX7\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 120.0, markup: 30.0, lowStock: 5 },
  { id: "p1646", name: "STEELWIRE", category: "Electrical", unit: "kg", quantity: 0, costPrice: 495.0, markup: 40.0, lowStock: 5 },
  { id: "p1647", name: "STELL CUTTING DISC 7\" THIK", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 210.0, markup: 40.0, lowStock: 5 },
  { id: "p1648", name: "STELL NAIL", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 350.0, markup: 40.0, lowStock: 5 },
  { id: "p1649", name: "STICKEY  DISK 4\" ELLOW", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 70.0, markup: 30.0, lowStock: 5 },
  { id: "p1650", name: "SUCTION PIPE", category: "Plumbing & Sanitary", unit: "ft", quantity: 0, costPrice: 55.79, markup: 30.0, lowStock: 5 },
  { id: "p1651", name: "SUMBBI", category: "General", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p1652", name: "SUPER MASTER BIB COCK", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 625.0, markup: 40.0, lowStock: 5 },
  { id: "p1653", name: "SYPHAN BIG", category: "General", unit: "piece", quantity: 0, costPrice: 850.0, markup: 30.0, lowStock: 5 },
  { id: "p1654", name: "SYPHAN SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 350.0, markup: 30.0, lowStock: 5 },
  { id: "p1655", name: "SYPHON TANK FITTING ABS", category: "General", unit: "piece", quantity: 0, costPrice: 550.0, markup: 40.0, lowStock: 5 },
  { id: "p1656", name: "SYPHON TANK FITTING LEVER", category: "General", unit: "piece", quantity: 0, costPrice: 650.0, markup: 40.0, lowStock: 5 },
  { id: "p1657", name: "SYPHON TANK FITTING UNIVERSAL", category: "General", unit: "piece", quantity: 0, costPrice: 470.0, markup: 40.0, lowStock: 5 },
  { id: "p1658", name: "T COCK", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 700.0, markup: 30.0, lowStock: 5 },
  { id: "p1659", name: "T COCK AONE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 690.0, markup: 35.0, lowStock: 5 },
  { id: "p1660", name: "T COCK ASIA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 420.0, markup: 40.0, lowStock: 5 },
  { id: "p1661", name: "T COCK CHINA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 185.0, markup: 30.0, lowStock: 5 },
  { id: "p1662", name: "T COCK MASTER NICE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 570.0, markup: 40.0, lowStock: 5 },
  { id: "p1663", name: "T COCK PLASTIC", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 155.0, markup: 40.0, lowStock: 5 },
  { id: "p1664", name: "T COCK PVC", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 280.0, markup: 40.0, lowStock: 5 },
  { id: "p1665", name: "T COCK RBS", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 810.0, markup: 40.0, lowStock: 5 },
  { id: "p1666", name: "T COCK RT", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 720.0, markup: 35.0, lowStock: 5 },
  { id: "p1667", name: "T COCK S ASIA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 630.0, markup: 40.0, lowStock: 5 },
  { id: "p1668", name: "T COCK S.MASTER", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 645.0, markup: 40.0, lowStock: 5 },
  { id: "p1669", name: "T COCK SMALL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 540.0, markup: 30.0, lowStock: 5 },
  { id: "p1670", name: "T COCK SONEX", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1380.0, markup: 40.0, lowStock: 5 },
  { id: "p1671", name: "T COCK TPS", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 630.0, markup: 40.0, lowStock: 5 },
  { id: "p1672", name: "T PANA 13 14", category: "General", unit: "piece", quantity: 0, costPrice: 195.0, markup: 40.0, lowStock: 5 },
  { id: "p1673", name: "TAISI", category: "General", unit: "piece", quantity: 0, costPrice: 191.67, markup: 35.0, lowStock: 5 },
  { id: "p1674", name: "TAISI CHIPPING", category: "General", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p1675", name: "TAISI DASTA", category: "General", unit: "piece", quantity: 0, costPrice: 30.0, markup: 40.0, lowStock: 5 },
  { id: "p1676", name: "TAKORA", category: "General", unit: "piece", quantity: 0, costPrice: 75.0, markup: 30.0, lowStock: 5 },
  { id: "p1677", name: "TANKI BUSH", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 140.0, markup: 40.0, lowStock: 5 },
  { id: "p1678", name: "TANKI BUSH 1", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p1679", name: "TANKI BUSH 1-1/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 275.0, markup: 30.0, lowStock: 5 },
  { id: "p1680", name: "TANKI BUSH 1-1/4X1", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 365.0, markup: 40.0, lowStock: 5 },
  { id: "p1681", name: "TANKI BUSH 1/2", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 140.0, markup: 40.0, lowStock: 5 },
  { id: "p1682", name: "TANKI BUSH 1X3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 185.0, markup: 40.0, lowStock: 5 },
  { id: "p1683", name: "TANKI BUSH 1X3/4 BHARI", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 190.0, markup: 40.0, lowStock: 5 },
  { id: "p1684", name: "TANKI BUSH 1X3/4 SADA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 140.0, markup: 40.0, lowStock: 5 },
  { id: "p1685", name: "TANKI BUSH 3/4", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p1686", name: "TANKI BUSH 3/4X1/2 BHARI", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 140.0, markup: 40.0, lowStock: 5 },
  { id: "p1687", name: "TANKI BUSH 3/4X1/2 SADA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 110.0, markup: 40.0, lowStock: 5 },
  { id: "p1688", name: "TANTED COLOUR", category: "Paint", unit: "piece", quantity: 0, costPrice: 24.17, markup: 40.0, lowStock: 5 },
  { id: "p1689", name: "TAPE", category: "General", unit: "piece", quantity: 0, costPrice: 45.0, markup: 30.0, lowStock: 5 },
  { id: "p1690", name: "TASTER", category: "General", unit: "piece", quantity: 0, costPrice: 39.0, markup: 40.0, lowStock: 5 },
  { id: "p1691", name: "TEE ARM FOR SHOWER HEAD", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 315.0, markup: 40.0, lowStock: 5 },
  { id: "p1692", name: "TEFLON TAPE", category: "General", unit: "piece", quantity: 0, costPrice: 15.0, markup: 35.0, lowStock: 5 },
  { id: "p1693", name: "TEFLON TAPE BID", category: "General", unit: "piece", quantity: 0, costPrice: 24.0, markup: 30.0, lowStock: 5 },
  { id: "p1694", name: "TEFLON TAPE BIG", category: "General", unit: "piece", quantity: 0, costPrice: 15.5, markup: 40.0, lowStock: 5 },
  { id: "p1695", name: "TEFLON TAPE SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 15.0, markup: 40.0, lowStock: 5 },
  { id: "p1696", name: "TELEPHONE 2 COR", category: "General", unit: "piece", quantity: 0, costPrice: 2050.0, markup: 40.0, lowStock: 5 },
  { id: "p1697", name: "TESTER", category: "Electrical", unit: "piece", quantity: 0, costPrice: 38.0, markup: 30.0, lowStock: 5 },
  { id: "p1698", name: "TESTER BLUE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 28.0, markup: 40.0, lowStock: 5 },
  { id: "p1699", name: "TESTER GOLDEN", category: "Electrical", unit: "piece", quantity: 0, costPrice: 28.0, markup: 40.0, lowStock: 5 },
  { id: "p1700", name: "TESTER HEAVY DUTY RED", category: "Electrical", unit: "piece", quantity: 0, costPrice: 55.0, markup: 35.0, lowStock: 5 },
  { id: "p1701", name: "TESTER RED", category: "Electrical", unit: "piece", quantity: 0, costPrice: 54.0, markup: 40.0, lowStock: 5 },
  { id: "p1702", name: "TESTER RED BIG", category: "Electrical", unit: "piece", quantity: 0, costPrice: 53.0, markup: 40.0, lowStock: 5 },
  { id: "p1703", name: "TESTER WHITE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 45.0, markup: 35.0, lowStock: 5 },
  { id: "p1704", name: "TESTER YELLOW", category: "Electrical", unit: "piece", quantity: 0, costPrice: 38.0, markup: 30.0, lowStock: 5 },
  { id: "p1705", name: "THIMMAL", category: "Electrical", unit: "piece", quantity: 0, costPrice: 220.0, markup: 40.0, lowStock: 5 },
  { id: "p1706", name: "THIMMAL 10MM", category: "Electrical", unit: "piece", quantity: 0, costPrice: 30.0, markup: 35.0, lowStock: 5 },
  { id: "p1707", name: "THIMMAL 16MM", category: "Electrical", unit: "piece", quantity: 0, costPrice: 32.0, markup: 35.0, lowStock: 5 },
  { id: "p1708", name: "THIMMAL 6MM", category: "Electrical", unit: "piece", quantity: 0, costPrice: 20.0, markup: 35.0, lowStock: 5 },
  { id: "p1709", name: "THINNER", category: "General", unit: "piece", quantity: 0, costPrice: 240.0, markup: 30.0, lowStock: 5 },
  { id: "p1710", name: "THREADED ROD 10MM", category: "General", unit: "piece", quantity: 0, costPrice: 410.0, markup: 35.0, lowStock: 5 },
  { id: "p1711", name: "THREE PANA", category: "General", unit: "piece", quantity: 0, costPrice: 240.0, markup: 40.0, lowStock: 5 },
  { id: "p1712", name: "TIFLON TAPE", category: "General", unit: "piece", quantity: 0, costPrice: 15.5, markup: 30.0, lowStock: 5 },
  { id: "p1713", name: "TIGER ACID SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 20.83, markup: 35.0, lowStock: 5 },
  { id: "p1714", name: "TIKKI 3X3", category: "General", unit: "piece", quantity: 0, costPrice: 20.0, markup: 40.0, lowStock: 5 },
  { id: "p1715", name: "TILE SPACER", category: "General", unit: "piece", quantity: 0, costPrice: 85.0, markup: 30.0, lowStock: 5 },
  { id: "p1716", name: "TIN SNIP 300MM 12\"", category: "General", unit: "piece", quantity: 0, costPrice: 650.0, markup: 40.0, lowStock: 5 },
  { id: "p1717", name: "TMALMAL THAN", category: "General", unit: "piece", quantity: 0, costPrice: 375.0, markup: 40.0, lowStock: 5 },
  { id: "p1718", name: "TORUS SCREW 8X1/2", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 145.0, markup: 40.0, lowStock: 5 },
  { id: "p1719", name: "TP 100AMP", category: "General", unit: "piece", quantity: 0, costPrice: 2000.0, markup: 40.0, lowStock: 5 },
  { id: "p1720", name: "TREE CUTTER", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 330.0, markup: 40.0, lowStock: 5 },
  { id: "p1721", name: "TRUSS 1/2X8", category: "General", unit: "piece", quantity: 0, costPrice: 178.0, markup: 40.0, lowStock: 5 },
  { id: "p1722", name: "TUBE LIGHT 1FT", category: "General", unit: "piece", quantity: 0, costPrice: 430.0, markup: 40.0, lowStock: 5 },
  { id: "p1723", name: "TUBE LIGHT 2FT", category: "General", unit: "piece", quantity: 0, costPrice: 580.0, markup: 40.0, lowStock: 5 },
  { id: "p1724", name: "TUFF BULB 12W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 112.0, markup: 30.0, lowStock: 5 },
  { id: "p1725", name: "TUFF BULB 18W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 270.0, markup: 30.0, lowStock: 5 },
  { id: "p1726", name: "TUFF BULB 50W", category: "Electrical", unit: "piece", quantity: 0, costPrice: 1050.0, markup: 30.0, lowStock: 5 },
  { id: "p1727", name: "TURKEY UMBER", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 94.42, markup: 40.0, lowStock: 5 },
  { id: "p1728", name: "TURKY AMBER", category: "General", unit: "piece", quantity: 0, costPrice: 95.0, markup: 40.0, lowStock: 5 },
  { id: "p1729", name: "TURKY TOOTI", category: "General", unit: "piece", quantity: 0, costPrice: 110.0, markup: 30.0, lowStock: 5 },
  { id: "p1730", name: "TV CABLE RJ7   1X55 MTR", category: "Electrical", unit: "piece", quantity: 0, costPrice: 1150.0, markup: 40.0, lowStock: 5 },
  { id: "p1731", name: "U CLAMP 3\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 17.5, markup: 40.0, lowStock: 5 },
  { id: "p1732", name: "U CLAMP 4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 20.0, markup: 40.0, lowStock: 5 },
  { id: "p1733", name: "UNIFIX", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 12.4, markup: 40.0, lowStock: 5 },
  { id: "p1734", name: "UNIFIX 10MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 18.0, markup: 40.0, lowStock: 5 },
  { id: "p1735", name: "UNIFIX FICHER", category: "Hardware & Tools", unit: "packet", quantity: 0, costPrice: 900.0, markup: 35.0, lowStock: 5 },
  { id: "p1736", name: "UNION 25MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 98.0, markup: 40.0, lowStock: 5 },
  { id: "p1737", name: "UNION 32MM", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 129.5, markup: 40.0, lowStock: 5 },
  { id: "p1738", name: "VANITY MIXER sonex", category: "General", unit: "piece", quantity: 0, costPrice: 7000.0, markup: 40.0, lowStock: 5 },
  { id: "p1739", name: "VANITY WAIST", category: "General", unit: "piece", quantity: 0, costPrice: 250.0, markup: 40.0, lowStock: 5 },
  { id: "p1740", name: "VANITY WAIST PIPE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p1741", name: "VIPER GREEN LATOO", category: "General", unit: "piece", quantity: 0, costPrice: 250.0, markup: 40.0, lowStock: 5 },
  { id: "p1742", name: "VIPER RUBBER ONLY", category: "General", unit: "piece", quantity: 0, costPrice: 40.0, markup: 40.0, lowStock: 5 },
  { id: "p1743", name: "VIPER STEEL H DUTY", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 430.0, markup: 40.0, lowStock: 5 },
  { id: "p1744", name: "VIPER STEEL ORDINORY", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 280.0, markup: 40.0, lowStock: 5 },
  { id: "p1745", name: "WAHOO LOCK 40MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 270.0, markup: 40.0, lowStock: 5 },
  { id: "p1746", name: "WAHOO LOCK 50MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 330.0, markup: 40.0, lowStock: 5 },
  { id: "p1747", name: "WAHOO LOCK 60MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 390.0, markup: 40.0, lowStock: 5 },
  { id: "p1748", name: "WAHOO LOCK 70MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 465.0, markup: 40.0, lowStock: 5 },
  { id: "p1749", name: "WAHU LOCK  30MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 195.0, markup: 40.0, lowStock: 5 },
  { id: "p1750", name: "WAHU LOCK  40MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 285.0, markup: 40.0, lowStock: 5 },
  { id: "p1751", name: "WAHU LOCK  50MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 385.0, markup: 40.0, lowStock: 5 },
  { id: "p1752", name: "WAHU LOCK  60MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 420.0, markup: 40.0, lowStock: 5 },
  { id: "p1753", name: "WAHU LOCK 40MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 320.0, markup: 40.0, lowStock: 5 },
  { id: "p1754", name: "WAHU LOCK 50MM", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 380.0, markup: 40.0, lowStock: 5 },
  { id: "p1755", name: "WAIST NORMAL", category: "General", unit: "piece", quantity: 0, costPrice: 27.5, markup: 40.0, lowStock: 5 },
  { id: "p1756", name: "WAIST PIPE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 160.0, markup: 40.0, lowStock: 5 },
  { id: "p1757", name: "WAIST PIPE  COMPLETE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 75.0, markup: 40.0, lowStock: 5 },
  { id: "p1758", name: "WAIST PIPE CHROME", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 250.0, markup: 40.0, lowStock: 5 },
  { id: "p1759", name: "WAIST PIPE COMPLETE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 70.0, markup: 30.0, lowStock: 5 },
  { id: "p1760", name: "WAIST PIPE GRAY", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 45.0, markup: 35.0, lowStock: 5 },
  { id: "p1761", name: "WAIST PIPE GRAY COMPLETE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 85.0, markup: 35.0, lowStock: 5 },
  { id: "p1762", name: "WAIST PIPE NORMAL", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 40.0, markup: 30.0, lowStock: 5 },
  { id: "p1763", name: "WAIST PIPE SADA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 40.0, markup: 40.0, lowStock: 5 },
  { id: "p1764", name: "WAIST PIPE SONIC", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 150.0, markup: 40.0, lowStock: 5 },
  { id: "p1765", name: "WAIST PIPE WHITE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 150.0, markup: 40.0, lowStock: 5 },
  { id: "p1766", name: "WAIST PIPIE", category: "General", unit: "piece", quantity: 0, costPrice: 38.0, markup: 40.0, lowStock: 5 },
  { id: "p1767", name: "WALL BOX 3X3", category: "General", unit: "piece", quantity: 0, costPrice: 48.0, markup: 40.0, lowStock: 5 },
  { id: "p1768", name: "WALL BOX 6X3", category: "General", unit: "piece", quantity: 0, costPrice: 3900.0, markup: 40.0, lowStock: 5 },
  { id: "p1769", name: "WALL GRINDING DISC 5\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 580.0, markup: 40.0, lowStock: 5 },
  { id: "p1770", name: "WALL HANGING BASIN", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 11500.0, markup: 40.0, lowStock: 5 },
  { id: "p1771", name: "WALL HOLDER", category: "Electrical", unit: "piece", quantity: 0, costPrice: 46.67, markup: 40.0, lowStock: 5 },
  { id: "p1772", name: "WAPDA CABLE 7/52", category: "Electrical", unit: "bndl", quantity: 0, costPrice: 7200.0, markup: 40.0, lowStock: 5 },
  { id: "p1773", name: "WAPDA CABLE 7/64", category: "Electrical", unit: "roll", quantity: 0, costPrice: 9500.0, markup: 40.0, lowStock: 5 },
  { id: "p1774", name: "WAPDA FULL GAUGE CABLE 7/64 SERVICE CABLE", category: "Electrical", unit: "piece", quantity: 0, costPrice: 13800.0, markup: 30.0, lowStock: 5 },
  { id: "p1775", name: "WARNISH", category: "General", unit: "piece", quantity: 0, costPrice: 235.0, markup: 30.0, lowStock: 5 },
  { id: "p1776", name: "WARNISH BIG", category: "General", unit: "piece", quantity: 0, costPrice: 520.0, markup: 40.0, lowStock: 5 },
  { id: "p1777", name: "WARNISH SMALL", category: "General", unit: "piece", quantity: 0, costPrice: 190.0, markup: 40.0, lowStock: 5 },
  { id: "p1778", name: "WASH BASIN", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 2800.0, markup: 40.0, lowStock: 5 },
  { id: "p1779", name: "WASH BASON 3STAR ALFA", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 8000.0, markup: 40.0, lowStock: 5 },
  { id: "p1780", name: "WASH BASON MIXER PVC", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 1050.0, markup: 40.0, lowStock: 5 },
  { id: "p1781", name: "WASH BASON NECK", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 148.0, markup: 40.0, lowStock: 5 },
  { id: "p1782", name: "WASH BASON PIPE 18\"", category: "Plumbing & Sanitary", unit: "pcs", quantity: 0, costPrice: 85.0, markup: 40.0, lowStock: 5 },
  { id: "p1783", name: "WASHER", category: "General", unit: "kg", quantity: 0, costPrice: 500.0, markup: 30.0, lowStock: 5 },
  { id: "p1784", name: "WASHER SMALL", category: "General", unit: "packet", quantity: 0, costPrice: 80.0, markup: 40.0, lowStock: 5 },
  { id: "p1785", name: "WASHING MACHINE PIPE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 45.0, markup: 40.0, lowStock: 5 },
  { id: "p1786", name: "WATER PROOF TAPE", category: "General", unit: "piece", quantity: 0, costPrice: 90.0, markup: 40.0, lowStock: 5 },
  { id: "p1787", name: "WATER PROOF TAPE QJC", category: "General", unit: "piece", quantity: 0, costPrice: 380.0, markup: 30.0, lowStock: 5 },
  { id: "p1788", name: "WATER TANK KING MASTER   1000LTR", category: "General", unit: "44900", quantity: 0, costPrice: 11220.0, markup: 30.0, lowStock: 5 },
  { id: "p1789", name: "WATER TANK KING MASTER   1200LTR", category: "General", unit: "51000", quantity: 0, costPrice: 12988.0, markup: 30.0, lowStock: 5 },
  { id: "p1790", name: "WATER TANK KING MASTER   2000LTR", category: "General", unit: "73000", quantity: 0, costPrice: 20128.0, markup: 30.0, lowStock: 5 },
  { id: "p1791", name: "WATER TANK KING MASTER   250LTR", category: "General", unit: "16150", quantity: 0, costPrice: 3944.0, markup: 30.0, lowStock: 5 },
  { id: "p1792", name: "WATER TANK KING MASTER   500LTR", category: "General", unit: "24800", quantity: 0, costPrice: 6018.0, markup: 30.0, lowStock: 5 },
  { id: "p1793", name: "WATTI", category: "General", unit: "piece", quantity: 0, costPrice: 65.0, markup: 30.0, lowStock: 5 },
  { id: "p1794", name: "WC", category: "General", unit: "piece", quantity: 0, costPrice: 1450.0, markup: 40.0, lowStock: 5 },
  { id: "p1795", name: "WC CONNECTION PIPE", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 125.0, markup: 40.0, lowStock: 5 },
  { id: "p1796", name: "WC ENGLISH", category: "General", unit: "piece", quantity: 0, costPrice: 7500.0, markup: 40.0, lowStock: 5 },
  { id: "p1797", name: "WC PIPE BIG", category: "Plumbing & Sanitary", unit: "pair", quantity: 0, costPrice: 175.0, markup: 40.0, lowStock: 5 },
  { id: "p1798", name: "WC PIPE SMALL", category: "Plumbing & Sanitary", unit: "pair", quantity: 0, costPrice: 165.0, markup: 40.0, lowStock: 5 },
  { id: "p1799", name: "WC WASHER", category: "General", unit: "piece", quantity: 0, costPrice: 35.0, markup: 40.0, lowStock: 5 },
  { id: "p1800", name: "WD 40", category: "General", unit: "piece", quantity: 0, costPrice: 170.0, markup: 40.0, lowStock: 5 },
  { id: "p1801", name: "WELDIN ROD 10#   168 PCS", category: "General", unit: "ctn", quantity: 0, costPrice: 1500.0, markup: 40.0, lowStock: 5 },
  { id: "p1802", name: "WELDIN ROD 12#   150 PCS", category: "General", unit: "ctn", quantity: 0, costPrice: 800.0, markup: 40.0, lowStock: 5 },
  { id: "p1803", name: "WELDING GLASSES", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 55.0, markup: 40.0, lowStock: 5 },
  { id: "p1804", name: "WELDING GOOGLE BLACK", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 180.0, markup: 40.0, lowStock: 5 },
  { id: "p1805", name: "WELDING ROD 12#", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 920.0, markup: 35.0, lowStock: 5 },
  { id: "p1806", name: "WELDING SHELD", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 420.0, markup: 40.0, lowStock: 5 },
  { id: "p1807", name: "WELLBON PLIER BLACK", category: "General", unit: "piece", quantity: 0, costPrice: 620.0, markup: 40.0, lowStock: 5 },
  { id: "p1808", name: "WHEEL RED", category: "General", unit: "piece", quantity: 0, costPrice: 340.0, markup: 30.0, lowStock: 5 },
  { id: "p1809", name: "WHITE EMULSION", category: "General", unit: "gal", quantity: 0, costPrice: 1650.0, markup: 40.0, lowStock: 5 },
  { id: "p1810", name: "WHITE ROPE RASSI 214.80 KG TOTAL", category: "General", unit: "piece", quantity: 0, costPrice: 93.0, markup: 40.0, lowStock: 5 },
  { id: "p1811", name: "WHITE TOOTI", category: "General", unit: "piece", quantity: 0, costPrice: 140.0, markup: 30.0, lowStock: 5 },
  { id: "p1812", name: "WIPER H DUTY", category: "General", unit: "piece", quantity: 0, costPrice: 575.0, markup: 40.0, lowStock: 5 },
  { id: "p1813", name: "WIPER RUBBER ONLY", category: "General", unit: "piece", quantity: 0, costPrice: 43.33, markup: 30.0, lowStock: 5 },
  { id: "p1814", name: "WIRE 23/76", category: "Electrical", unit: "roll", quantity: 0, costPrice: 1250.0, markup: 40.0, lowStock: 5 },
  { id: "p1815", name: "WIRE BRUSH", category: "Electrical", unit: "piece", quantity: 0, costPrice: 75.0, markup: 35.0, lowStock: 5 },
  { id: "p1816", name: "WIRE CONNECTOR", category: "Electrical", unit: "piece", quantity: 0, costPrice: 60.0, markup: 40.0, lowStock: 5 },
  { id: "p1817", name: "WIRE CUP BRUSH BRASS", category: "Electrical", unit: "piece", quantity: 0, costPrice: 230.0, markup: 30.0, lowStock: 5 },
  { id: "p1818", name: "WIRE CUP BRUSH HARD", category: "Electrical", unit: "piece", quantity: 0, costPrice: 250.0, markup: 30.0, lowStock: 5 },
  { id: "p1819", name: "WOOD BLADE 7\"", category: "General", unit: "piece", quantity: 0, costPrice: 520.0, markup: 40.0, lowStock: 5 },
  { id: "p1820", name: "WOOD CLEANING BRUSH", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 330.0, markup: 40.0, lowStock: 5 },
  { id: "p1821", name: "WOOD CUTTING DISC 4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 430.0, markup: 40.0, lowStock: 5 },
  { id: "p1822", name: "WOOD GLUE", category: "Paint", unit: "tin", quantity: 0, costPrice: 100.0, markup: 40.0, lowStock: 5 },
  { id: "p1823", name: "WOOD GLUE 1 KG MT", category: "Paint", unit: "piece", quantity: 0, costPrice: 280.0, markup: 40.0, lowStock: 5 },
  { id: "p1824", name: "WOOD GLUE 1/2 KG", category: "Paint", unit: "piece", quantity: 0, costPrice: 145.0, markup: 40.0, lowStock: 5 },
  { id: "p1825", name: "WOOD GLUE 1KG", category: "Paint", unit: "piece", quantity: 0, costPrice: 270.0, markup: 40.0, lowStock: 5 },
  { id: "p1826", name: "WOOD GLUE GERMAN", category: "Paint", unit: "kg", quantity: 0, costPrice: 290.0, markup: 40.0, lowStock: 5 },
  { id: "p1827", name: "WOOD GLUE MOVILETH 1/2KG", category: "Paint", unit: "piece", quantity: 0, costPrice: 290.0, markup: 40.0, lowStock: 5 },
  { id: "p1828", name: "WOOD GLUE MOVILETH 1KG", category: "Paint", unit: "piece", quantity: 0, costPrice: 560.0, markup: 40.0, lowStock: 5 },
  { id: "p1829", name: "WOOD NAIL", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 390.0, markup: 30.0, lowStock: 5 },
  { id: "p1830", name: "WOOD NAIL 17X1", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 400.0, markup: 40.0, lowStock: 5 },
  { id: "p1831", name: "WOOD NAIL 1X12 S", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 430.0, markup: 35.0, lowStock: 5 },
  { id: "p1832", name: "WOOD NAIL 1X17", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 380.0, markup: 40.0, lowStock: 5 },
  { id: "p1833", name: "WOOD NAIL 1X20 T", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 470.0, markup: 35.0, lowStock: 5 },
  { id: "p1834", name: "WOOD NAIL 2/17", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 450.0, markup: 40.0, lowStock: 5 },
  { id: "p1835", name: "WOOD NAIL 2X15", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 420.0, markup: 40.0, lowStock: 5 },
  { id: "p1836", name: "WOOD NAIL 2X17", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 420.0, markup: 40.0, lowStock: 5 },
  { id: "p1837", name: "WOOD PHATTI", category: "General", unit: "piece", quantity: 0, costPrice: 18.46, markup: 40.0, lowStock: 5 },
  { id: "p1838", name: "WOOR HAND SAW AARI", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 175.0, markup: 30.0, lowStock: 5 },
  { id: "p1839", name: "WULB 50W", category: "General", unit: "piece", quantity: 0, costPrice: 750.0, markup: 40.0, lowStock: 5 },
  { id: "p1840", name: "Y TEE 4\"", category: "Plumbing & Sanitary", unit: "piece", quantity: 0, costPrice: 448.33, markup: 40.0, lowStock: 5 },
  { id: "p1841", name: "YEE 4\"", category: "General", unit: "piece", quantity: 0, costPrice: 372.0, markup: 40.0, lowStock: 5 },
  { id: "p1842", name: "YELLOW AMBER", category: "General", unit: "piece", quantity: 0, costPrice: 93.33, markup: 40.0, lowStock: 5 },
  { id: "p1843", name: "YELLOW DISC 5\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 95.0, markup: 40.0, lowStock: 5 },
  { id: "p1844", name: "YELLOW GLOVES", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 155.0, markup: 30.0, lowStock: 5 },
  { id: "p1845", name: "YELLOW OXIDE POWDER", category: "General", unit: "piece", quantity: 0, costPrice: 90.0, markup: 30.0, lowStock: 5 },
  { id: "p1846", name: "ZEROX DISC 4\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 520.0, markup: 40.0, lowStock: 5 },
  { id: "p1847", name: "ZEROX DISC 5\"", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 580.0, markup: 40.0, lowStock: 5 },
  { id: "p1848", name: "ZEROX DISC 5'", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 490.0, markup: 40.0, lowStock: 5 },
  { id: "p1849", name: "y", category: "General", unit: "piece", quantity: 0, costPrice: 520.0, markup: 40.0, lowStock: 5 },
];

/* ---------------------------------------------------------
   ROOT APP
--------------------------------------------------------- */
export default function HardwareInventoryApp({ user, onLogout }) {
  const [view, setView] = useState("dashboard");
  const [products, setProducts] = useState([]);
  const [invoices, setInvoices] = useState([]);
  const [expenses, setExpenses] = useState([]);
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [loading, setLoading] = useState(true);
  const [toast, setToast] = useState(null);
  const [cart, setCart] = useState([]); // draft invoice items
  const [customerName, setCustomerName] = useState("");
  const [customerEmail, setCustomerEmail] = useState("");
  const [customerPhone, setCustomerPhone] = useState("");
  const [discount, setDiscount] = useState(0);
  const [lastInvoice, setLastInvoice] = useState(null);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [editingInvoice, setEditingInvoice] = useState(null);
  const [editingQuotation, setEditingQuotation] = useState(null);

  const [quotationCart, setQuotationCart] = useState([]);
  const [quotationCustName, setQuotationCustName] = useState("");
  const [quotationCustEmail, setQuotationCustEmail] = useState("");
  const [quotationCustPhone, setQuotationCustPhone] = useState("");
  const [quotationDiscount, setQuotationDiscount] = useState(0);
  const [quotations, setQuotations] = useState(() => {
    try {
      const saved = localStorage.getItem("hw_quotations");
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  useEffect(() => {
    try {
      localStorage.setItem("hw_quotations", JSON.stringify(quotations));
    } catch (err) {
      console.error("Save quotations to local storage failed:", err);
    }
  }, [quotations]);

  const showToast = useCallback((msg, kind = "ok") => {
    setToast({ msg, kind });
    setTimeout(() => setToast(null), 2600);
  }, []);

  useEffect(() => {
    if (!user) return;
    
    const loadUserData = async () => {
      setLoading(true);
      try {
        const data = await window.db.fetchData();
        
        let finalSettings = data.settings;
        if (!finalSettings) {
          finalSettings = {
            ...DEFAULT_SETTINGS,
            shopName: user.organizationName || DEFAULT_SETTINGS.shopName,
            email: user.email || '',
            phone: user.phone || ''
          };
          await window.db.saveSettings(finalSettings);
        } else {
          finalSettings = {
            ...DEFAULT_SETTINGS,
            ...finalSettings,
            email: finalSettings.email || user.email || '',
            phone: finalSettings.phone || user.phone || '',
            shopName: finalSettings.shopName || user.organizationName || DEFAULT_SETTINGS.shopName
          };
        }

        let finalProducts = data.products || [];

        setProducts(finalProducts);
        setInvoices(data.invoices || []);
        setExpenses(data.expenses || []);
        setSettings(finalSettings);
      } catch (err) {
        console.error("Error loading data from cloud:", err);
        showToast("Error loading shop data from database", "err");
      } finally {
        setLoading(false);
      }
    };

    loadUserData();
  }, [user, showToast]);

  const persistProducts = useCallback(async (next) => {
    if (!user) return;
    setProducts(next);
    try {
      await window.db.saveProducts(next);
    } catch (err) {
      console.error("Save products error:", err);
      showToast("Could not save products to cloud", "err");
    }
  }, [user, showToast]);

  const persistInvoices = useCallback(async (next) => {
    if (!user) return;
    setInvoices(next);
    try {
      await window.db.saveInvoices(next);
    } catch (err) {
      console.error("Save invoices error:", err);
      showToast("Could not save invoice to cloud", "err");
    }
  }, [user, showToast]);

  const persistExpenses = useCallback(async (next) => {
    if (!user) return;
    setExpenses(next);
    try {
      await window.db.saveExpenses(next);
    } catch (err) {
      console.error("Save expenses error:", err);
      showToast("Could not save expense to cloud", "err");
    }
  }, [user, showToast]);

  const persistSettings = useCallback(async (next) => {
    if (!user) return;
    setSettings(next);
    try {
      await window.db.saveSettings(next);
    } catch (err) {
      console.error("Save settings error:", err);
      showToast("Could not save settings to cloud", "err");
    }
  }, [user, showToast]);

  const handleLogout = async () => {
    try {
      await window.db.logout();
      if (onLogout) onLogout();
    } catch (err) {
      console.error("Signout error:", err);
      showToast("Could not sign out. Try again.", "err");
    }
  };

  const lowStockItems = useMemo(() => products.filter(p => p.quantity <= (p.lowStock ?? 0)), [products]);
  const totalStockValue = useMemo(() => products.reduce((s, p) => s + p.quantity * sellPrice(p), 0), [products]);
  const todayStr = new Date().toDateString();
  const invoicesToday = useMemo(() => invoices.filter(i => new Date(i.date).toDateString() === todayStr), [invoices]);
  const revenueToday = useMemo(() => invoicesToday.reduce((s, i) => s + i.total, 0), [invoicesToday]);

  if (loading) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100vh", width: "100vw", background: "var(--bg)", color: "var(--ink-soft)", fontFamily: "var(--font-body)" }}>
        <LoadingOverlay message="Loading your shop data from cloud..." />
      </div>
    );
  }

  return (
    <div className="hw-root">
      <Style />
      <aside className="hw-sidebar">
        <div className="hw-brand">
          <div className="hw-brand-mark">⛏</div>
          <div>
            <div className="hw-brand-name">{settings.shopName}</div>
            <div className="hw-brand-sub">Inventory &amp; Billing</div>
          </div>
        </div>
        <nav className="hw-nav">
          <NavBtn icon={<LayoutDashboard size={17} />} label="Dashboard" active={view === "dashboard"} onClick={() => setView("dashboard")} />
          <NavBtn icon={<Package size={17} />} label="Inventory" active={view === "inventory"} onClick={() => setView("inventory")} badge={lowStockItems.length || null} />
          <NavBtn icon={<ShoppingCart size={17} />} label="New Invoice" active={view === "newInvoice"} onClick={() => setView("newInvoice")} />
          <NavBtn icon={<FileClock size={17} />} label="Invoice History" active={view === "invoices"} onClick={() => setView("invoices")} />
          <NavBtn icon={<Receipt size={17} />} label="Quotation" active={view === "quotation"} onClick={() => setView("quotation")} />
          <NavBtn icon={<Wallet size={17} />} label="Daily Ledger" active={view === "ledger"} onClick={() => setView("ledger")} />
        </nav>
        <div style={{ display: "flex", flexDirection: "column", gap: 6, marginTop: 10 }}>
          <button className="hw-settings-btn" onClick={() => setSettingsOpen(true)}>
            <SettingsIcon size={16} /> Settings
          </button>
          <button className="hw-settings-btn" onClick={handleLogout} style={{ borderColor: "rgba(179, 58, 58, 0.45)", color: "#EFEAE0" }}>
            <LogOut size={16} /> Log Out
          </button>
        </div>
      </aside>

      <main className="hw-main">
        {view === "dashboard" && (
          <Dashboard
            products={products} invoices={invoices} settings={settings}
            lowStockItems={lowStockItems} totalStockValue={totalStockValue}
            invoicesToday={invoicesToday} revenueToday={revenueToday}
            goInventory={() => setView("inventory")} goInvoice={() => setView("newInvoice")}
          />
        )}
        {view === "inventory" && (
          <Inventory products={products} settings={settings} persistProducts={persistProducts} showToast={showToast} />
        )}
        {view === "newInvoice" && (
          <NewInvoice
            products={products} settings={settings} cart={cart} setCart={setCart}
            customerName={customerName} setCustomerName={setCustomerName}
            customerEmail={customerEmail} setCustomerEmail={setCustomerEmail}
            customerPhone={customerPhone} setCustomerPhone={setCustomerPhone}
            discount={discount} setDiscount={setDiscount}
            persistProducts={persistProducts} persistInvoices={persistInvoices}
            persistSettings={persistSettings}
            invoices={invoices} showToast={showToast} setLastInvoice={setLastInvoice}
            setView={setView} user={user}
            editingInvoice={editingInvoice} setEditingInvoice={setEditingInvoice}
          />
        )}
        {view === "invoices" && (
          <InvoiceHistory invoices={invoices} settings={settings} persistInvoices={persistInvoices} persistProducts={persistProducts} products={products} showToast={showToast}
            onEditInvoice={(inv) => {
              setEditingInvoice(inv);
              setCart(inv.items.map(i => {
                const prod = products.find(p => p.id === i.productId);
                const originalStock = prod ? prod.quantity + i.qty : i.qty;
                return { productId: i.productId, name: i.name, unit: i.unit, qty: i.qty, costPrice: i.costPrice, markup: i.markup, maxStock: originalStock };
              }));
              setCustomerName(inv.customerName === "Walk-in Customer" ? "" : (inv.customerName || ""));
              setCustomerEmail(inv.customerEmail || "");
              setCustomerPhone(inv.customerPhone || "");
              setDiscount(inv.discount || 0);
              setView("newInvoice");
              showToast(`Loaded ${inv.invoiceNumber} for editing`);
            }}
          />
        )}
        {view === "ledger" && (
          <DailyLedger invoices={invoices} expenses={expenses} settings={settings} persistExpenses={persistExpenses} showToast={showToast} products={products} />
        )}
        {view === "quotation" && (
          <QuotationView
            products={products} settings={settings}
            cart={quotationCart} setCart={setQuotationCart}
            customerName={quotationCustName} setCustomerName={setQuotationCustName}
            customerEmail={quotationCustEmail} setCustomerEmail={setQuotationCustEmail}
            customerPhone={quotationCustPhone} setCustomerPhone={setQuotationCustPhone}
            discount={quotationDiscount} setDiscount={setQuotationDiscount}
            quotations={quotations} setQuotations={setQuotations}
            showToast={showToast} setView={setView}
            editingQuotation={editingQuotation} setEditingQuotation={setEditingQuotation}
            onConvertToInvoice={(items, custName, custEmail, custPhone, disc) => {
              setEditingInvoice(null);
              setCart(items.map(i => ({ productId: i.productId, name: i.name, unit: i.unit, qty: i.qty, costPrice: i.costPrice, markup: i.markup, maxStock: products.find(p => p.id === i.productId)?.quantity || 9999 })));
              setCustomerName(custName);
              setCustomerEmail(custEmail);
              setCustomerPhone(custPhone || "");
              setDiscount(disc);
              setView("newInvoice");
              showToast("Quotation loaded into Invoice Cart!");
            }}
          />
        )}
      </main>

      {settingsOpen && (
        <SettingsModal settings={settings} onClose={() => setSettingsOpen(false)} onSave={(s) => { persistSettings(s); setSettingsOpen(false); showToast("Settings saved"); }} />
      )}

      {toast && (
        <div className={`hw-toast ${toast.kind === "err" ? "hw-toast-err" : ""}`}>
          {toast.kind === "err" ? <AlertTriangle size={15} /> : <Check size={15} />} {toast.msg}
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   NAV BUTTON
--------------------------------------------------------- */
function NavBtn({ icon, label, active, onClick, badge }) {
  return (
    <button className={`hw-navbtn ${active ? "active" : ""}`} onClick={onClick}>
      {icon}
      <span>{label}</span>
      {badge ? <span className="hw-navbadge">{badge}</span> : null}
      {active && <ChevronRight size={14} className="hw-navchevron" />}
    </button>
  );
}

/* ---------------------------------------------------------
   DASHBOARD
--------------------------------------------------------- */
function Dashboard({ products, invoices, settings, lowStockItems, totalStockValue, invoicesToday, revenueToday, goInventory, goInvoice }) {
  const recent = [...invoices].sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 6);
  const cs = settings.currencySymbol;
  return (
    <div className="hw-view">
      <ViewHeader eyebrow="Overview" title="Dashboard" />
      <div className="hw-stat-grid">
        <StatCard label="Products tracked" value={products.length} icon={<Package size={18} />} />
        <StatCard label="Stock value" value={`${cs}${fmtNum(totalStockValue)}`} icon={<ArrowUpCircle size={18} />} accent="steel" />
        <StatCard label="Low stock items" value={lowStockItems.length} icon={<AlertTriangle size={18} />} accent={lowStockItems.length ? "warn" : "ok"} onClick={goInventory} />
        <StatCard label="Invoiced today" value={`${cs}${fmtNum(revenueToday)}`} sub={`${invoicesToday.length} invoice${invoicesToday.length === 1 ? "" : "s"}`} icon={<Receipt size={18} />} accent="accent" />
      </div>

      <div className="hw-dash-cols">
        <div className="hw-card">
          <div className="hw-card-head">
            <h3>Low stock</h3>
            <button className="hw-link" onClick={goInventory}>Manage inventory <ChevronRight size={13} /></button>
          </div>
          {lowStockItems.length === 0 ? (
            <EmptyRow text="Nothing running low. Stock levels look healthy." />
          ) : (
            <ul className="hw-simple-list">
              {lowStockItems.slice(0, 8).map(p => (
                <li key={p.id} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 0' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <span className="hw-hazard-dot" />
                    <span className="hw-il-name" style={{ fontWeight: 500 }}>{p.name}</span>
                  </div>
                  <span style={{ fontSize: '12px', fontWeight: 600, color: 'var(--danger)' }}>
                    {p.quantity} {p.unit} left
                  </span>
                </li>
              ))}
            </ul>
          )}
        </div>

        <div className="hw-card">
          <div className="hw-card-head">
            <h3>Recent invoices</h3>
            <button className="hw-link" onClick={goInvoice}>New invoice <ChevronRight size={13} /></button>
          </div>
          {recent.length === 0 ? (
            <EmptyRow text="No invoices yet. Create your first one." />
          ) : (
            <ul className="hw-simple-list">
              {recent.map(inv => (
                <li key={inv.id} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 0' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <span className="hw-inv-num">{inv.invoiceNumber}</span>
                    <span className="hw-il-name" style={{ fontWeight: 600 }}>{inv.customerName || "Walk-in Customer"}</span>
                  </div>
                  <span className="hw-il-meta" style={{ fontWeight: 700, color: 'var(--ink)' }}>
                    {cs}{fmtNum(inv.total)}
                  </span>
                </li>
              ))}
            </ul>
          )}
        </div>
      </div>
    </div>
  );
}

function StatCard({ label, value, sub, icon, accent, onClick }) {
  return (
    <div className={`hw-stat hw-stat-${accent || "default"}`} onClick={onClick} style={onClick ? { cursor: "pointer" } : undefined}>
      <div className="hw-stat-icon">{icon}</div>
      <div className="hw-stat-value">{value}</div>
      <div className="hw-stat-label">{label}</div>
      {sub && <div className="hw-stat-sub">{sub}</div>}
    </div>
  );
}

function EmptyRow({ text }) {
  return <div className="hw-empty">{text}</div>;
}

function ViewHeader({ eyebrow, title, right }) {
  return (
    <div className="hw-view-head">
      <div>
        <div className="hw-eyebrow">{eyebrow}</div>
        <h1 className="hw-title">{title}</h1>
      </div>
      {right}
    </div>
  );
}

/* ---------------------------------------------------------
   INVENTORY
--------------------------------------------------------- */
function Inventory({ products, settings, persistProducts, showToast }) {
  const [query, setQuery] = useState("");
  const [editing, setEditing] = useState(null); // product or "new"
  const [adjusting, setAdjusting] = useState(null); // {product, dir}
  const [selectedIds, setSelectedIds] = useState([]);
  const [bulkEditType, setBulkEditType] = useState(null); // "category" | "markup" | "stock" | null

  const [excelData, setExcelData] = useState([]);
  const [excelHeaders, setExcelHeaders] = useState([]);
  const [excelMappings, setExcelMappings] = useState({
    name: -1,
    quantity: -1,
    unit: -1,
    costPrice: -1,
    salePrice: -1,
    category: -1,
    lowStock: -1
  });
  const [showExcelMapper, setShowExcelMapper] = useState(false);
  const [hasHeaderRow, setHasHeaderRow] = useState(true);

  const displayHeaders = useMemo(() => {
    if (!excelData || excelData.length === 0) return [];
    const columnsCount = excelHeaders.length;
    const getColumnLetter = (colIdx) => {
      let letter = "";
      let temp = colIdx;
      while (temp >= 0) {
        letter = String.fromCharCode((temp % 26) + 65) + letter;
        temp = Math.floor(temp / 26) - 1;
      }
      return letter;
    };
    return hasHeaderRow ? excelHeaders : Array.from({ length: columnsCount }, (_, i) => `Column ${getColumnLetter(i)}`);
  }, [hasHeaderRow, excelHeaders, excelData]);

  const handleExcelUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1 });

        if (rawRows.length < 1) {
          showToast("Excel file is empty", "err");
          return;
        }

        const cleanedRows = rawRows.filter(row => row && row.some(cell => cell !== null && cell !== ''));
        if (cleanedRows.length === 0) {
          showToast("Excel file contains no data", "err");
          return;
        }

        setExcelData(cleanedRows);

        const firstRow = cleanedRows[0] || [];
        const columnsCount = Math.max(...cleanedRows.slice(0, 5).map(r => r.length), firstRow.length);
        const getColumnLetter = (colIdx) => {
          let letter = "";
          let temp = colIdx;
          while (temp >= 0) {
            letter = String.fromCharCode((temp % 26) + 65) + letter;
            temp = Math.floor(temp / 26) - 1;
          }
          return letter;
        };

        const detectedHeaders = firstRow.map((cell, idx) => {
          if (cell !== null && cell !== undefined && String(cell).trim() !== '') {
            return String(cell).trim();
          }
          return `Column ${getColumnLetter(idx)}`;
        });

        while (detectedHeaders.length < columnsCount) {
          detectedHeaders.push(`Column ${getColumnLetter(detectedHeaders.length)}`);
        }

        setExcelHeaders(detectedHeaders);

        const lowerHeaders = detectedHeaders.map(h => h.toLowerCase());
        const nameIdx = lowerHeaders.findIndex(h => h.includes("name") || h.includes("title") || h.includes("product") || h.includes("item") || h.includes("desc"));
        const qtyIdx = lowerHeaders.findIndex(h => h.includes("qty") || h.includes("quantity") || h.includes("stock") || h.includes("count") || h.includes("avail"));
        const unitIdx = lowerHeaders.findIndex(h => h.includes("unit") || h.includes("pack") || h.includes("measure"));
        
        let costIdx = lowerHeaders.findIndex(h => h.includes("cost") || h.includes("buying") || h.includes("purchase") || h.includes("buy") || h.includes("cp"));
        let saleIdx = lowerHeaders.findIndex(h => h.includes("sale") || h.includes("selling") || h.includes("retail") || h.includes("sp"));

        if (costIdx === -1 && saleIdx === -1) {
          const priceIdx = lowerHeaders.findIndex(h => h === "price" || h.includes("price"));
          costIdx = priceIdx;
        } else if (costIdx === -1) {
          costIdx = lowerHeaders.findIndex(h => h.includes("price") && h !== (detectedHeaders[saleIdx] ? detectedHeaders[saleIdx].toLowerCase() : ""));
        } else if (saleIdx === -1) {
          saleIdx = lowerHeaders.findIndex(h => h.includes("price") && h !== (detectedHeaders[costIdx] ? detectedHeaders[costIdx].toLowerCase() : ""));
        }

        const categoryIdx = lowerHeaders.findIndex(h => h.includes("category") || h.includes("type") || h.includes("dept") || h.includes("group"));
        const lowStockIdx = lowerHeaders.findIndex(h => h.includes("low") || h.includes("alert") || h.includes("min") || h.includes("warn"));

        setExcelMappings({
          name: nameIdx,
          quantity: qtyIdx,
          unit: unitIdx,
          costPrice: costIdx,
          salePrice: saleIdx,
          category: categoryIdx,
          lowStock: lowStockIdx
        });

        setHasHeaderRow(true);
        setShowExcelMapper(true);
      } catch (err) {
        console.error("Error reading excel:", err);
        showToast("Failed to parse Excel file", "err");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const handleFinalExcelImport = async () => {
    try {
      const startIdx = hasHeaderRow ? 1 : 0;
      const importedProducts = [];

      if (excelMappings.name === -1) {
        showToast("Product Name column mapping is required", "err");
        return;
      }

      for (let i = startIdx; i < excelData.length; i++) {
        const row = excelData[i];
        if (!row || row.length === 0) continue;

        const name = excelMappings.name !== -1 && row[excelMappings.name] 
          ? String(row[excelMappings.name]).trim() 
          : null;

        if (!name || name === "0" || name.startsWith("==") || name === "") {
          continue;
        }

        let costPrice = 0;
        if (excelMappings.costPrice !== -1 && row[excelMappings.costPrice] !== undefined && row[excelMappings.costPrice] !== null) {
          costPrice = Number(row[excelMappings.costPrice]) || 0;
        }

        let salePrice = 0;
        if (excelMappings.salePrice !== -1 && row[excelMappings.salePrice] !== undefined && row[excelMappings.salePrice] !== null) {
          salePrice = Number(row[excelMappings.salePrice]) || 0;
        }

        let markup = 40;
        if (costPrice > 0 && salePrice > 0) {
          markup = roundNum(((salePrice - costPrice) / costPrice) * 100, 1);
          if (markup < 0 || markup > 1000) markup = 40;
        }

        let quantity = 0;
        if (excelMappings.quantity !== -1 && row[excelMappings.quantity] !== undefined && row[excelMappings.quantity] !== null) {
          quantity = Number(row[excelMappings.quantity]) || 0;
        }

        let category = "General";
        if (excelMappings.category !== -1 && row[excelMappings.category]) {
          category = String(row[excelMappings.category]).trim();
        } else {
          category = categorizeProductName(name);
        }

        let unit = "piece";
        if (excelMappings.unit !== -1 && row[excelMappings.unit]) {
          unit = String(row[excelMappings.unit]).trim().toLowerCase();
          if (unit === "" || unit === "none") unit = "piece";
        }

        let lowStock = 5;
        if (excelMappings.lowStock !== -1 && row[excelMappings.lowStock] !== undefined && row[excelMappings.lowStock] !== null) {
          lowStock = parseInt(row[excelMappings.lowStock]) || 5;
        }

        importedProducts.push({
          id: `p_imp_${Math.random().toString(36).slice(2, 10)}`,
          name,
          category,
          unit,
          quantity,
          costPrice,
          markup,
          lowStock,
        });
      }

      if (importedProducts.length === 0) {
        showToast("No valid products found to import", "err");
        return;
      }

      const next = [...products];
      let idCounter = next.length + 1;
      
      const finalImported = importedProducts.map(p => {
        const existing = next.find(x => x.name.toLowerCase() === p.name.toLowerCase());
        if (existing) {
          return {
            ...existing,
            quantity: existing.quantity + p.quantity,
            costPrice: p.costPrice > 0 ? p.costPrice : existing.costPrice,
            markup: p.costPrice > 0 && p.salePrice > 0 ? p.markup : existing.markup,
          };
        } else {
          const newId = `p${String(idCounter++).padStart(4, "0")}`;
          return {
            ...p,
            id: newId,
          };
        }
      });

      const mergedProducts = [...products];
      finalImported.forEach(imp => {
        const idx = mergedProducts.findIndex(x => x.id === imp.id);
        if (idx !== -1) {
          mergedProducts[idx] = imp;
        } else {
          mergedProducts.push(imp);
        }
      });

      await persistProducts(mergedProducts);
      showToast(`Imported/updated ${finalImported.length} products successfully!`);
      setShowExcelMapper(false);
    } catch (err) {
      console.error("Error finalizing Excel import:", err);
      showToast("Failed to finalize import", "err");
    }
  };

  const filtered = products.filter(p =>
    p.name.toLowerCase().includes(query.toLowerCase()) ||
    p.category.toLowerCase().includes(query.toLowerCase())
  ).sort((a, b) => a.name.localeCompare(b.name));

  // Reset selection if query changes and selected items are no longer visible
  useEffect(() => {
    setSelectedIds(prev => prev.filter(id => filtered.some(p => p.id === id)));
  }, [query]);

  const upsertProduct = async (prod) => {
    let next;
    if (products.some(p => p.id === prod.id)) {
      next = products.map(p => p.id === prod.id ? prod : p);
    } else {
      next = [...products, prod];
    }
    await persistProducts(next);
    setEditing(null);
    showToast("Product saved");
  };

  const deleteProduct = async (id) => {
    await persistProducts(products.filter(p => p.id !== id));
    setEditing(null);
    showToast("Product removed");
  };

  const quickAdjust = async (product, delta) => {
    const newQty = Math.max(0, product.quantity + delta);
    const next = products.map(p => p.id === product.id ? { ...p, quantity: newQty } : p);
    await persistProducts(next);
  };

  const toggleSelect = (id) => {
    setSelectedIds(prev => prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id]);
  };

  const toggleSelectAll = () => {
    if (filtered.length > 0 && selectedIds.length === filtered.length) {
      setSelectedIds([]);
    } else {
      setSelectedIds(filtered.map(p => p.id));
    }
  };

  const handleBulkSave = async (value) => {
    const next = products.map(p => {
      if (selectedIds.includes(p.id)) {
        if (bulkEditType === "category") {
          return { ...p, category: value.trim() };
        } else if (bulkEditType === "markup") {
          return { ...p, markup: Number(value) || 0 };
        } else if (bulkEditType === "stock") {
          const valStr = String(value).trim();
          let newQty = p.quantity;
          if (valStr.startsWith("+")) {
            newQty = Math.max(0, p.quantity + (Number(valStr.slice(1)) || 0));
          } else if (valStr.startsWith("-")) {
            newQty = Math.max(0, p.quantity - (Number(valStr.slice(1)) || 0));
          } else {
            newQty = Math.max(0, Number(valStr) || 0);
          }
          return { ...p, quantity: newQty };
        }
      }
      return p;
    });
    await persistProducts(next);
    setSelectedIds([]);
    setBulkEditType(null);
    showToast(`Bulk updated ${selectedIds.length} items`);
  };

  const handleBulkDelete = async () => {
    if (window.confirm(`Are you sure you want to delete ${selectedIds.length} selected items?`)) {
      const next = products.filter(p => !selectedIds.includes(p.id));
      await persistProducts(next);
      setSelectedIds([]);
      showToast(`Deleted ${selectedIds.length} items`);
    }
  };

  return (
    <div className="hw-view">
      <ViewHeader
        eyebrow={`${products.length} product${products.length === 1 ? "" : "s"}`}
        title="Inventory"
        right={
          <div style={{ display: "flex", gap: 8 }}>
            <label className="hw-btn-ghost" style={{ cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 6, fontSize: "13px", padding: "8px 12px", border: "1px solid var(--border)", borderRadius: "6px" }}>
              <ArrowUpCircle size={16} />
              Import Excel/CSV
              <input 
                type="file" 
                accept=".xlsx,.xls,.csv" 
                style={{ display: "none" }} 
                onChange={handleExcelUpload} 
              />
            </label>
            <button className="hw-btn-accent" onClick={() => setEditing("new")}><Plus size={16} /> Add product</button>
          </div>
        }
      />

      <div className="hw-search-row">
        <Search size={16} />
        <input placeholder="Search by name or category…" value={query} onChange={e => setQuery(e.target.value)} />
      </div>

      {selectedIds.length > 0 && (
        <div className="hw-bulk-actions-bar">
          <span>{selectedIds.length} item{selectedIds.length === 1 ? "" : "s"} selected</span>
          <button className="hw-btn-accent" onClick={() => setBulkEditType("category")}>Update Category</button>
          <button className="hw-btn-accent" onClick={() => setBulkEditType("markup")}>Update Markup %</button>
          <button className="hw-btn-accent" onClick={() => setBulkEditType("stock")}>Update Stock</button>
          <button className="hw-btn-danger" onClick={handleBulkDelete}><Trash2 size={14} /> Delete Selected</button>
          <button className="hw-btn-ghost" onClick={() => setSelectedIds([])}>Clear</button>
        </div>
      )}

      <div className="hw-table-wrap">
        <table className="hw-table">
          <thead>
            <tr>
              <th style={{ width: 40, textAlign: "center" }}>
                <input 
                  type="checkbox" 
                  checked={filtered.length > 0 && selectedIds.length === filtered.length} 
                  onChange={toggleSelectAll} 
                  style={{ cursor: "pointer" }}
                />
              </th>
              <th>Product</th>
              <th>Category</th>
              <th>Unit</th>
              <th>Stock</th>
              <th>Cost</th>
              <th>Markup</th>
              <th>Sell price</th>
              <th>Value</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map(p => {
              const low = p.quantity <= (p.lowStock ?? 0);
              const sp = sellPrice(p);
              return (
                <tr key={p.id} className={low ? "hw-row-low" : ""}>
                  <td style={{ textAlign: "center" }}>
                    <input 
                      type="checkbox" 
                      checked={selectedIds.includes(p.id)} 
                      onChange={() => toggleSelect(p.id)} 
                      style={{ cursor: "pointer" }}
                    />
                  </td>
                  <td>
                    <div className="hw-prod-name">{p.name}</div>
                    {low && <div className="hw-low-tag"><AlertTriangle size={11} /> Low stock</div>}
                  </td>
                  <td className="hw-muted">{p.category || "—"}</td>
                  <td className="hw-muted">{p.unit}</td>
                  <td>
                    <div className="hw-qty-cell">
                      <button onClick={() => quickAdjust(p, -1)} disabled={p.quantity <= 0}><Minus size={13} /></button>
                      <span className="hw-mono">{p.quantity}</span>
                      <button onClick={() => quickAdjust(p, 1)}><Plus size={13} /></button>
                    </div>
                  </td>
                  <td className="hw-mono hw-muted">{settings.currencySymbol}{fmtNum(p.costPrice)}</td>
                  <td className="hw-mono hw-muted">{fmtNum(p.markup)}%</td>
                  <td className="hw-mono">{settings.currencySymbol}{fmtNum(sp)}</td>
                  <td className="hw-mono hw-muted">{settings.currencySymbol}{fmtNum(p.quantity * sp)}</td>
                  <td>
                    <button className="hw-icon-btn" onClick={() => setEditing(p)}><Pencil size={14} /></button>
                  </td>
                </tr>
              );
            })}
            {filtered.length === 0 && (
              <tr><td colSpan={10}><EmptyRow text={query ? "No products match your search." : "No products yet — add your first one."} /></td></tr>
            )}
          </tbody>
        </table>
      </div>

      {editing && (
        <ProductModal
          product={editing === "new" ? null : editing}
          defaultLowStock={settings.lowStockDefault}
          onClose={() => setEditing(null)}
          onSave={upsertProduct}
          onDelete={editing !== "new" ? () => deleteProduct(editing.id) : null}
        />
      )}

      {bulkEditType && (
        <BulkEditModal
          type={bulkEditType}
          selectedCount={selectedIds.length}
          onClose={() => setBulkEditType(null)}
          onSave={handleBulkSave}
        />
      )}

      {/* Excel Column Mapper Modal */}
      {showExcelMapper && (
        <div style={modalStyles.overlay}>
          <div style={modalStyles.modalCard}>
            <div style={modalStyles.header}>
              <h2 style={modalStyles.title}>⚙️ Map Excel Columns</h2>
              <p style={modalStyles.subtitle}>Align your spreadsheet columns with the store inventory details.</p>
            </div>

            {/* Has Header Checkbox */}
            <div style={modalStyles.checkboxWrapper}>
              <label style={modalStyles.checkboxLabel}>
                <input
                  type="checkbox"
                  checked={hasHeaderRow}
                  onChange={(e) => setHasHeaderRow(e.target.checked)}
                  style={modalStyles.checkbox}
                />
                First row contains column headers (e.g. "Name", "Cost", "Stock")
              </label>
            </div>

            {/* Data Preview */}
            <div style={modalStyles.previewSection}>
              <div style={modalStyles.sectionTitle}>Data Preview (First few rows)</div>
              <div style={modalStyles.previewTableScroll}>
                <table style={modalStyles.table}>
                  <thead>
                    <tr>
                      {displayHeaders.map((h, idx) => (
                        <th key={idx} style={modalStyles.th}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {excelData.slice(hasHeaderRow ? 1 : 0, hasHeaderRow ? 4 : 3).map((row, rIdx) => (
                      <tr key={rIdx} style={modalStyles.tr}>
                        {displayHeaders.map((_, cIdx) => (
                          <td key={cIdx} style={modalStyles.td}>
                            {row[cIdx] !== undefined ? String(row[cIdx]).slice(0, 30) : ""}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Mappings Form */}
            <div style={modalStyles.formSection}>
              <div style={modalStyles.sectionTitle}>Select Matching Columns</div>
              <div style={modalStyles.grid}>
                
                {/* Required: Name */}
                <div style={modalStyles.field}>
                  <label style={modalStyles.label}>Product Name <span style={{color: '#ea580c'}}>*</span></label>
                  <select
                    value={excelMappings.name}
                    onChange={(e) => setExcelMappings(prev => ({ ...prev, name: parseInt(e.target.value) }))}
                    style={modalStyles.select}
                  >
                    <option value={-1}>-- Select Column --</option>
                    {displayHeaders.map((h, idx) => (
                      <option key={idx} value={idx}>{h}</option>
                    ))}
                  </select>
                </div>

                {/* Required: Cost Price */}
                <div style={modalStyles.field}>
                  <label style={modalStyles.label}>Cost Price <span style={{color: '#ea580c'}}>*</span></label>
                  <select
                    value={excelMappings.costPrice}
                    onChange={(e) => setExcelMappings(prev => ({ ...prev, costPrice: parseInt(e.target.value) }))}
                    style={modalStyles.select}
                  >
                    <option value={-1}>-- Select Column --</option>
                    {displayHeaders.map((h, idx) => (
                      <option key={idx} value={idx}>{h}</option>
                    ))}
                  </select>
                </div>

                {/* Optional: Sale Price */}
                <div style={modalStyles.field}>
                  <label style={modalStyles.label}>Sale Price (Optional - calculates markup)</label>
                  <select
                    value={excelMappings.salePrice}
                    onChange={(e) => setExcelMappings(prev => ({ ...prev, salePrice: parseInt(e.target.value) }))}
                    style={modalStyles.select}
                  >
                    <option value={-1}>-- None (Use 40% markup) --</option>
                    {displayHeaders.map((h, idx) => (
                      <option key={idx} value={idx}>{h}</option>
                    ))}
                  </select>
                </div>

                {/* Optional: Quantity */}
                <div style={modalStyles.field}>
                  <label style={modalStyles.label}>Quantity / Stock (Optional)</label>
                  <select
                    value={excelMappings.quantity}
                    onChange={(e) => setExcelMappings(prev => ({ ...prev, quantity: parseInt(e.target.value) }))}
                    style={modalStyles.select}
                  >
                    <option value={-1}>-- None (Defaults to 0) --</option>
                    {displayHeaders.map((h, idx) => (
                      <option key={idx} value={idx}>{h}</option>
                    ))}
                  </select>
                </div>

                {/* Optional: Category */}
                <div style={modalStyles.field}>
                  <label style={modalStyles.label}>Category (Optional)</label>
                  <select
                    value={excelMappings.category}
                    onChange={(e) => setExcelMappings(prev => ({ ...prev, category: parseInt(e.target.value) }))}
                    style={modalStyles.select}
                  >
                    <option value={-1}>-- Auto-Categorize by Name --</option>
                    {displayHeaders.map((h, idx) => (
                      <option key={idx} value={idx}>{h}</option>
                    ))}
                  </select>
                </div>

                {/* Optional: Unit */}
                <div style={modalStyles.field}>
                  <label style={modalStyles.label}>Unit (Optional)</label>
                  <select
                    value={excelMappings.unit}
                    onChange={(e) => setExcelMappings(prev => ({ ...prev, unit: parseInt(e.target.value) }))}
                    style={modalStyles.select}
                  >
                    <option value={-1}>-- None (Defaults to "piece") --</option>
                    {displayHeaders.map((h, idx) => (
                      <option key={idx} value={idx}>{h}</option>
                    ))}
                  </select>
                </div>

                {/* Optional: Low Stock Limit */}
                <div style={modalStyles.field}>
                  <label style={modalStyles.label}>Low Stock Limit (Optional)</label>
                  <select
                    value={excelMappings.lowStock}
                    onChange={(e) => setExcelMappings(prev => ({ ...prev, lowStock: parseInt(e.target.value) }))}
                    style={modalStyles.select}
                  >
                    <option value={-1}>-- None (Defaults to 5) --</option>
                    {displayHeaders.map((h, idx) => (
                      <option key={idx} value={idx}>{h}</option>
                    ))}
                  </select>
                </div>

              </div>
            </div>

            {/* Actions */}
            <div style={modalStyles.actions}>
              <button
                type="button"
                onClick={() => setShowExcelMapper(false)}
                style={modalStyles.btnCancel}
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleFinalExcelImport}
                disabled={excelMappings.name === -1 || excelMappings.costPrice === -1}
                style={excelMappings.name === -1 || excelMappings.costPrice === -1 ? modalStyles.btnImportDisabled : modalStyles.btnImport}
              >
                Import Products
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

const modalStyles = {
  overlay: {
    position: 'fixed',
    top: 0, left: 0, right: 0, bottom: 0,
    background: 'rgba(38, 36, 32, 0.75)',
    backdropFilter: 'blur(4px)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    zIndex: 9999,
    padding: '20px',
    boxSizing: 'border-box'
  },
  modalCard: {
    background: '#FFFFFF',
    border: '1px solid #E0D9C9',
    borderRadius: '16px',
    width: '100%',
    maxWidth: '800px',
    padding: '28px',
    boxShadow: '0 10px 40px rgba(38, 36, 32, 0.15)',
    fontFamily: "'Inter', sans-serif",
    color: '#262420',
    maxHeight: '90vh',
    display: 'flex',
    flexDirection: 'column',
    overflowY: 'auto',
    boxSizing: 'border-box'
  },
  header: {
    marginBottom: '20px',
    borderBottom: '1px solid #E5DFD3',
    paddingBottom: '12px'
  },
  title: {
    fontSize: '22px',
    fontWeight: '700',
    margin: 0,
    color: '#262420',
  },
  subtitle: {
    fontSize: '13px',
    color: '#746C5E',
    marginTop: '4px',
    marginRight: 0, marginBottom: 0, marginLeft: 0
  },
  checkboxWrapper: {
    marginBottom: '16px',
  },
  checkboxLabel: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    fontSize: '13.5px',
    fontWeight: '500',
    color: '#262420',
    cursor: 'pointer'
  },
  checkbox: {
    accentColor: '#D9720B',
    width: '16px',
    height: '16px',
    cursor: 'pointer'
  },
  previewSection: {
    marginBottom: '24px',
    background: '#FAF9F6',
    border: '1px solid #E5DFD3',
    borderRadius: '8px',
    padding: '14px'
  },
  sectionTitle: {
    fontSize: '13px',
    fontWeight: '700',
    color: '#746C5E',
    letterSpacing: '0.5px',
    textTransform: 'uppercase',
    marginBottom: '10px'
  },
  previewTableScroll: {
    overflowX: 'auto',
    maxHeight: '130px',
    overflowY: 'auto',
    border: '1px solid #EAE5D9',
    borderRadius: '6px'
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
    fontSize: '12px',
    textAlign: 'left'
  },
  th: {
    background: '#F0EBE0',
    color: '#524B40',
    padding: '8px 12px',
    fontWeight: '600',
    borderBottom: '1px solid #EAE5D9',
    whiteSpace: 'nowrap'
  },
  tr: {
    borderBottom: '1px solid #F3EFE6'
  },
  td: {
    padding: '8px 12px',
    color: '#524B40',
    whiteSpace: 'nowrap',
    background: '#FFFFFF'
  },
  formSection: {
    marginBottom: '24px'
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
    gap: '16px'
  },
  field: {
    display: 'flex',
    flexDirection: 'column',
    gap: '6px'
  },
  label: {
    fontSize: '12px',
    fontWeight: '600',
    color: '#746C5E'
  },
  select: {
    padding: '8px 10px',
    border: '1px solid #E0D9C9',
    borderRadius: '6px',
    fontSize: '13px',
    background: '#F6F3EC',
    outline: 'none',
    color: '#262420',
    cursor: 'pointer',
    width: '100%'
  },
  actions: {
    display: 'flex',
    justifyContent: 'flex-end',
    gap: '12px',
    borderTop: '1px solid #E5DFD3',
    paddingTop: '20px',
    marginTop: 'auto'
  },
  btnCancel: {
    padding: '9px 18px',
    border: '1px solid #E0D9C9',
    borderRadius: '8px',
    background: 'transparent',
    color: '#746C5E',
    fontWeight: '600',
    fontSize: '13.5px',
    cursor: 'pointer',
    fontFamily: "'Inter', sans-serif",
    transition: 'background-color 0.2s'
  },
  btnImport: {
    padding: '9px 22px',
    border: 'none',
    borderRadius: '8px',
    background: '#D9720B',
    color: '#FFFFFF',
    fontWeight: '600',
    fontSize: '13.5px',
    cursor: 'pointer',
    fontFamily: "'Inter', sans-serif",
    transition: 'background-color 0.2s'
  },
  btnImportDisabled: {
    padding: '9px 22px',
    border: 'none',
    borderRadius: '8px',
    background: '#EAE5D9',
    color: '#A89F90',
    fontWeight: '600',
    fontSize: '13.5px',
    cursor: 'not-allowed',
    fontFamily: "'Inter', sans-serif"
  }
};

function ProductModal({ product, defaultLowStock, onClose, onSave, onDelete }) {
  const [form, setForm] = useState(product || { id: uid(), name: "", category: "", unit: "piece", quantity: 0, costPrice: 0, markup: 0, lowStock: defaultLowStock });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const valid = form.name.trim().length > 0 && form.unit.trim().length > 0;

  return (
    <ModalShell onClose={onClose} title={product ? "Edit product" : "Add product"}>
      <div className="hw-form-grid">
        <Field label="Product name" span={2}>
          <input autoFocus value={form.name} onChange={e => set("name", e.target.value)} placeholder="e.g. Portland Cement 50kg" />
        </Field>
        <Field label="Category">
          <input value={form.category} onChange={e => set("category", e.target.value)} placeholder="e.g. Cement" />
        </Field>
        <Field label="Unit">
          <input value={form.unit} onChange={e => set("unit", e.target.value)} placeholder="bag / piece / ton" />
        </Field>
        <Field label="Quantity in stock">
          <NumInput value={form.quantity} onChange={n => set("quantity", Math.max(0, n))} />
        </Field>
        <Field label="Cost price (what you paid)">
          <NumInput step="0.01" value={form.costPrice} onChange={n => {
            const newCost = Math.max(0, n);
            const currentSell = (Number(form.costPrice) || 0) * (1 + (Number(form.markup) || 0) / 100);
            const newMarkup = newCost > 0 ? ((currentSell - newCost) / newCost) * 100 : 0;
            setForm(f => ({ ...f, costPrice: newCost, markup: newMarkup }));
          }} />
        </Field>
        <Field label="Markup %">
          <NumInput step="0.1" value={fmtNum(form.markup)} onChange={n => set("markup", Math.max(0, n))} />
        </Field>
        <Field label="Sell price">
          <NumInput step="0.01" value={fmtNum((Number(form.costPrice) || 0) * (1 + (Number(form.markup) || 0) / 100))} onChange={n => {
            const newSell = Math.max(0, n);
            const cost = Number(form.costPrice) || 0;
            const newMarkup = cost > 0 ? ((newSell - cost) / cost) * 100 : 0;
            set("markup", newMarkup);
          }} />
          <div className="hw-hint">Type either the markup % or the sell price directly — the other updates itself. Handy for items priced case-by-case.</div>
        </Field>
        <Field label="Reorder alert level" span={2}>
          <NumInput value={form.lowStock} onChange={n => set("lowStock", Math.max(0, n))} />
          <div className="hw-hint">You'll see a low-stock warning once quantity drops to this number or below. Sell price = cost price + markup %, and updates automatically as either changes.</div>
        </Field>
      </div>
      <div className="hw-modal-actions">
        {onDelete && <button className="hw-btn-danger" onClick={onDelete}><Trash2 size={14} /> Delete</button>}
        <div style={{ flex: 1 }} />
        <button className="hw-btn-ghost" onClick={onClose}>Cancel</button>
        <button className="hw-btn-accent" disabled={!valid} onClick={() => onSave(form)}><Save size={14} /> Save product</button>
      </div>
    </ModalShell>
  );
}

function BulkEditModal({ type, selectedCount, onClose, onSave }) {
  const [val, setVal] = useState("");
  
  const title = {
    category: "Bulk Update Category",
    markup: "Bulk Update Markup %",
    stock: "Bulk Update Stock Quantity"
  }[type];

  const label = {
    category: "New Category Name",
    markup: "New Markup Percentage (%)",
    stock: "Stock Adjustment (e.g. 50 to set, +10 to add, -5 to subtract)"
  }[type];

  const placeholder = {
    category: "e.g. Electrical, Plumbing",
    markup: "e.g. 40",
    stock: "e.g. +10, -5, 50"
  }[type];

  return (
    <ModalShell onClose={onClose} title={title}>
      <div style={{ padding: "8px 0" }}>
        <p style={{ margin: "0 0 16px 0", color: "var(--ink-soft)" }}>
          You are updating <strong>{selectedCount}</strong> selected items.
        </p>
        <Field label={label}>
          <input 
            autoFocus 
            type="text" 
            value={val} 
            onChange={e => setVal(e.target.value)} 
            placeholder={placeholder} 
          />
        </Field>
      </div>
      <div className="hw-modal-actions" style={{ marginTop: 24 }}>
        <button className="hw-btn-ghost" onClick={onClose}>Cancel</button>
        <button className="hw-btn-accent" disabled={!val.trim()} onClick={() => onSave(val)}><Save size={14} /> Update Items</button>
      </div>
    </ModalShell>
  );
}

/* ---------------------------------------------------------
   NEW INVOICE
--------------------------------------------------------- */
function NewInvoice({ products, settings, cart, setCart, customerName, setCustomerName, customerEmail, setCustomerEmail, customerPhone, setCustomerPhone, discount, setDiscount, persistProducts, persistInvoices, persistSettings, invoices, showToast, setLastInvoice, setView, user, editingInvoice, setEditingInvoice }) {
  const [search, setSearch] = useState("");
  const [preview, setPreview] = useState(null);
  const [isSaving, setIsSaving] = useState(false);
  const cs = settings.currencySymbol;

  const matches = search.trim() ? products.filter(p => p.name.toLowerCase().includes(search.toLowerCase())).slice(0, 12) : [];

  const addToCart = (product) => {
    setSearch("");
    setCart(c => {
      const existing = c.find(i => i.productId === product.id);
      if (existing) {
        return c.map(i => i.productId === product.id ? { ...i, qty: i.qty + 1 } : i);
      }
      return [...c, { productId: product.id, name: product.name, unit: product.unit, qty: 1, costPrice: product.costPrice, markup: product.markup, maxStock: product.quantity }];
    });
  };

  const updateItem = (productId, patch) => {
    setCart(c => c.map(i => i.productId === productId ? { ...i, ...patch } : i));
  };
  const removeItem = (productId) => setCart(c => c.filter(i => i.productId !== productId));

  const cartWithPrice = cart.map(i => ({ ...i, sellPrice: sellPrice({ costPrice: i.costPrice, markup: i.markup }), lineTotal: i.qty * sellPrice({ costPrice: i.costPrice, markup: i.markup }) }));
  const subtotal = cartWithPrice.reduce((s, i) => s + i.lineTotal, 0);
  const totalCost = cartWithPrice.reduce((s, i) => s + i.qty * i.costPrice, 0);
  const discountAmt = discount > 0 ? (discount <= 100 ? subtotal * (discount / 100) : discount) : 0;
  const total = Math.max(0, subtotal - discountAmt);
  const profit = total - totalCost;

  const overStock = cart.filter(i => i.qty > i.maxStock);

  const cancelEdit = () => {
    setEditingInvoice(null);
    setCart([]);
    setCustomerName("");
    setCustomerEmail("");
    setCustomerPhone("");
    setDiscount(0);
    showToast("Invoice edit cancelled");
  };

  const saveInvoice = async () => {
    if (cart.length === 0) return;
    setIsSaving(true);
    const isEdit = !!editingInvoice;
    const invoiceNumber = isEdit ? editingInvoice.invoiceNumber : `INV-${String(settings.invoiceCounter).padStart(4, "0")}`;
    const invoiceId = isEdit ? editingInvoice.id : uid();
    const invoiceDate = isEdit ? editingInvoice.date : new Date().toISOString();

    try {
      const invoice = {
        id: invoiceId,
        invoiceNumber,
        customerName: customerName.trim(),
        customerEmail: customerEmail.trim(),
        customerPhone: customerPhone.trim(),
        date: invoiceDate,
        items: cartWithPrice.map(i => ({ productId: i.productId, name: i.name, unit: i.unit, qty: i.qty, costPrice: i.costPrice, markup: i.markup, price: i.sellPrice, lineTotal: i.lineTotal })),
        subtotal, discount: discountAmt, total, totalCost, profit,
      };

      let nextProducts;
      if (isEdit) {
        const restoredMap = {};
        for (const item of editingInvoice.items) {
          restoredMap[item.productId] = (restoredMap[item.productId] || 0) + item.qty;
        }
        const cartMap = {};
        for (const item of cart) {
          cartMap[item.productId] = (cartMap[item.productId] || 0) + item.qty;
        }

        nextProducts = products.map(p => {
          const restored = restoredMap[p.id] || 0;
          const deducted = cartMap[p.id] || 0;
          const netChange = restored - deducted;
          return { ...p, quantity: Math.max(0, p.quantity + netChange) };
        });
      } else {
        nextProducts = products.map(p => {
          const item = cart.find(i => i.productId === p.id);
          return item ? { ...p, quantity: Math.max(0, p.quantity - item.qty) } : p;
        });
      }

      const nextInvoices = isEdit
        ? invoices.map(i => i.id === invoiceId ? invoice : i)
        : [...invoices, invoice];

      await persistInvoices(nextInvoices);
      await persistProducts(nextProducts);
      if (!isEdit) {
        await persistSettings({ ...settings, invoiceCounter: settings.invoiceCounter + 1 });
      }
      
      // Automatically email invoice via Google SMTP if customer email is provided
      if (invoice.customerEmail) {
        try {
          const res = await window.db.sendInvoiceEmail(invoice, user);
          if (res && !res.success) {
            console.error("Email send error:", res.error);
            showToast("Invoice saved, but email dispatch failed", "err");
          } else {
            showToast(isEdit ? "Invoice updated & emailed!" : "Invoice saved & emailed to customer!");
          }
        } catch (err) {
          console.error("Email dispatch exception:", err);
          showToast("Invoice saved, but email dispatch failed", "err");
        }
      } else {
        showToast(isEdit ? `${invoiceNumber} updated!` : `${invoiceNumber} saved`);
      }

      setCart([]);
      setCustomerName("");
      setCustomerEmail("");
      setCustomerPhone("");
      setDiscount(0);
      setEditingInvoice(null);
      setLastInvoice(invoice);
      setPreview(invoice);
    } catch (err) {
      console.error("Invoice save error:", err);
      showToast("Failed to save transaction", "err");
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="hw-view">
      {isSaving && <LoadingOverlay message={editingInvoice ? "Updating transaction..." : "Saving transaction & dispatching email..."} />}
      <ViewHeader 
        eyebrow={editingInvoice ? "Editing Transaction" : "Point of sale"} 
        title={editingInvoice ? `Edit Invoice (${editingInvoice.invoiceNumber})` : "New Invoice"}
        right={
          editingInvoice ? (
            <button className="hw-btn-ghost" onClick={cancelEdit} style={{ color: "var(--danger)" }}>
              Cancel Edit
            </button>
          ) : null
        }
      />

      <div className="hw-invoice-layout">
        <div className="hw-card">
          <div style={{ display: 'flex', gap: '12px' }}>
            <div style={{ flex: 1.2 }}>
              <Field label="Customer name (optional)">
                <input value={customerName} onChange={e => setCustomerName(e.target.value)} placeholder="Walk-in customer" />
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label="Customer email (optional)">
                <input type="email" value={customerEmail} onChange={e => setCustomerEmail(e.target.value)} placeholder="customer@email.com" />
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label="Customer phone (optional)">
                <input type="text" value={customerPhone} onChange={e => setCustomerPhone(e.target.value)} placeholder="e.g. 03328898666" />
              </Field>
            </div>
          </div>

          <div className="hw-search-row" style={{ marginTop: 14, position: "relative" }}>
            <Search size={16} />
            <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search products to add…" />
            {matches.length > 0 && (
              <div className="hw-autocomplete">
                {matches.map(p => (
                  <button key={p.id} onClick={() => addToCart(p)} disabled={p.quantity <= 0}>
                    <span>{p.name}</span>
                    <span className="hw-mono hw-muted">{cs}{fmtNum(sellPrice(p))} · {p.quantity} {p.unit} in stock</span>
                  </button>
                ))}
              </div>
            )}
          </div>

          {cart.length > 0 && (
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 14, marginBottom: -6, fontSize: '12.5px', color: 'var(--ink-soft)', fontWeight: 600 }}>
              <span>Items in Invoice ({cart.length})</span>
              <span>Total Units: {cart.reduce((s, i) => s + i.qty, 0)}</span>
            </div>
          )}

          <div className="hw-table-wrap" style={{ marginTop: 14 }}>
            <table className="hw-table hw-table-tight">
              <thead>
                <tr><th>Item</th><th>Qty</th><th>Cost</th><th>Markup</th><th>Sell price</th><th>Total</th><th></th></tr>
              </thead>
              <tbody>
                {cartWithPrice.map(i => (
                  <tr key={i.productId} className={i.qty > i.maxStock ? "hw-row-low" : ""}>
                    <td>{i.name}<div className="hw-il-meta">{i.unit}</div></td>
                    <td>
                      <NumInput className="hw-mini-input" min={0} value={i.qty} onChange={n => updateItem(i.productId, { qty: Math.max(0, n) })} />
                    </td>
                    <td>
                      <NumInput className="hw-mini-input" step="0.01" min={0} value={i.costPrice} onChange={n => updateItem(i.productId, { costPrice: Math.max(0, n) })} />
                    </td>
                    <td>
                      <NumInput className="hw-mini-input" step="0.1" min={0} value={i.markup} onChange={n => updateItem(i.productId, { markup: Math.max(0, n) })} style={{ width: 54 }} />%
                    </td>
                    <td>
                      <NumInput className="hw-mini-input" step="0.01" min={0} value={fmtNum(i.sellPrice)} onChange={n => {
                        const newSell = Math.max(0, n);
                        const cost = Number(i.costPrice) || 0;
                        const newMarkup = cost > 0 ? ((newSell - cost) / cost) * 100 : 0;
                        updateItem(i.productId, { markup: newMarkup });
                      }} />
                    </td>
                    <td className="hw-mono">{cs}{fmtNum(i.lineTotal)}</td>
                    <td><button className="hw-icon-btn" onClick={() => removeItem(i.productId)}><X size={14} /></button></td>
                  </tr>
                ))}
                {cart.length === 0 && <tr><td colSpan={7}><EmptyRow text="Search above to add items to this invoice." /></td></tr>}
              </tbody>
            </table>
          </div>
          <div className="hw-hint" style={{ marginTop: 8 }}>Cost price and markup % are for your records only — the printed invoice shows just the sell price and total.</div>
          {overStock.length > 0 && (
            <div className="hw-warn-banner"><AlertTriangle size={14} /> {overStock.length} item{overStock.length > 1 ? "s" : ""} exceed current stock on hand.</div>
          )}
        </div>

        <div className="hw-card hw-summary-card">
          <h3>Summary</h3>
          <div className="hw-summary-row"><span>Subtotal</span><span className="hw-mono">{cs}{fmtNum(subtotal)}</span></div>
          <div className="hw-summary-row">
            <span>Discount</span>
            <NumInput className="hw-mini-input" min={0} value={discount} onChange={n => setDiscount(Math.max(0, n))} style={{ width: 70 }} />
          </div>
          <div className="hw-hint" style={{ marginTop: -8, marginBottom: 8 }}>0–100 = percent, above 100 = flat amount off</div>
          <div className="hw-summary-row hw-summary-total"><span>Total</span><span className="hw-mono">{cs}{fmtNum(total)}</span></div>
          {cart.length > 0 && (
            <div className="hw-profit-box">
              <div className="hw-profit-label">Internal only — not printed</div>
              <div className="hw-summary-row"><span>Cost total</span><span className="hw-mono">{cs}{fmtNum(totalCost)}</span></div>
              <div className="hw-summary-row"><span>Profit</span><span className="hw-mono">{cs}{fmtNum(profit)}</span></div>
            </div>
          )}
          <button className="hw-btn-accent hw-btn-block" disabled={cart.length === 0} onClick={saveInvoice}>
            <Save size={15} /> {editingInvoice ? "Update Invoice" : "Save invoice"}
          </button>
          {editingInvoice && (
            <button className="hw-btn-ghost hw-btn-block" onClick={cancelEdit} style={{ marginTop: '8px' }}>
              Cancel Edit
            </button>
          )}
        </div>
      </div>

      {preview && (
        <InvoicePreviewModal invoice={preview} settings={settings} onClose={() => setPreview(null)} />
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   QUOTATIONS & ESTIMATES VIEW
--------------------------------------------------------- */
function QuotationView({ products, settings, cart, setCart, customerName, setCustomerName, customerEmail, setCustomerEmail, customerPhone, setCustomerPhone, discount, setDiscount, quotations, setQuotations, showToast, setView, onConvertToInvoice, editingQuotation, setEditingQuotation }) {
  const [subView, setSubView] = useState("new"); // "new" | "history"
  const [search, setSearch] = useState("");
  const [preview, setPreview] = useState(null);
  const [query, setQuery] = useState("");
  const cs = settings.currencySymbol;

  const matches = search.trim() ? products.filter(p => p.name.toLowerCase().includes(search.toLowerCase())).slice(0, 12) : [];

  const addToCart = (product) => {
    setSearch("");
    setCart(c => {
      const existing = c.find(i => i.productId === product.id);
      if (existing) {
        return c.map(i => i.productId === product.id ? { ...i, qty: i.qty + 1 } : i);
      }
      return [...c, { productId: product.id, name: product.name, unit: product.unit, qty: 1, costPrice: product.costPrice, markup: product.markup }];
    });
  };

  const updateItem = (productId, patch) => {
    setCart(c => c.map(i => i.productId === productId ? { ...i, ...patch } : i));
  };
  const removeItem = (productId) => setCart(c => c.filter(i => i.productId !== productId));

  const cartWithPrice = cart.map(i => ({ ...i, sellPrice: sellPrice({ costPrice: i.costPrice, markup: i.markup }), lineTotal: i.qty * sellPrice({ costPrice: i.costPrice, markup: i.markup }) }));
  const subtotal = cartWithPrice.reduce((s, i) => s + i.lineTotal, 0);
  const totalCost = cartWithPrice.reduce((s, i) => s + i.qty * i.costPrice, 0);
  const discountAmt = discount > 0 ? (discount <= 100 ? subtotal * (discount / 100) : discount) : 0;
  const total = Math.max(0, subtotal - discountAmt);
  const profit = total - totalCost;

  const cancelEdit = () => {
    setEditingQuotation(null);
    setCart([]);
    setCustomerName("");
    setCustomerEmail("");
    setCustomerPhone("");
    setDiscount(0);
    showToast("Quotation edit cancelled");
  };

  const saveQuotation = () => {
    if (cart.length === 0) return;
    const isEdit = !!editingQuotation;
    const quotationNumber = isEdit ? editingQuotation.quotationNumber : `QT-${String(quotations.length + 1).padStart(4, "0")}`;
    const quoteId = isEdit ? editingQuotation.id : uid();
    const quoteDate = isEdit ? editingQuotation.date : new Date().toISOString();

    const quote = {
      id: quoteId,
      quotationNumber,
      customerName: customerName.trim() || "Walk-in Customer",
      customerEmail: customerEmail.trim(),
      customerPhone: customerPhone.trim(),
      date: quoteDate,
      items: cartWithPrice.map(i => ({ productId: i.productId, name: i.name, unit: i.unit, qty: i.qty, costPrice: i.costPrice, markup: i.markup, price: i.sellPrice, lineTotal: i.lineTotal })),
      subtotal, discount: discountAmt, total, totalCost, profit,
    };

    const nextQuotations = isEdit
      ? quotations.map(q => q.id === quoteId ? quote : q)
      : [quote, ...quotations];

    setQuotations(nextQuotations);
    
    setCart([]);
    setCustomerName("");
    setCustomerEmail("");
    setCustomerPhone("");
    setDiscount(0);
    setEditingQuotation(null);
    setPreview(quote);
    showToast(isEdit ? `${quotationNumber} updated successfully!` : `${quotationNumber} saved locally!`);
  };

  const deleteQuotation = (id) => {
    setQuotations(q => q.filter(x => x.id !== id));
    showToast("Quotation deleted");
  };

  const filteredQuotes = quotations.filter(q =>
    q.quotationNumber.toLowerCase().includes(query.toLowerCase()) ||
    (q.customerName || "").toLowerCase().includes(query.toLowerCase())
  );

  return (
    <div className="hw-view">
      <ViewHeader
        eyebrow={editingQuotation ? "Editing Estimate" : "Estimates Manager"}
        title={editingQuotation ? `Edit Quotation (${editingQuotation.quotationNumber})` : "Quotations"}
        right={
          <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
            {editingQuotation && (
              <button className="hw-btn-ghost" onClick={cancelEdit} style={{ color: "var(--danger)" }}>
                Cancel Edit
              </button>
            )}
            <div style={{ display: "flex", border: "1px solid var(--border)", borderRadius: "6px", overflow: "hidden" }}>
              <button className={`hw-btn-ghost ${subView === "new" ? "active" : ""}`} onClick={() => setSubView("new")} style={{ border: "none", borderRadius: 0, padding: "8px 16px", background: subView === "new" ? "var(--border)" : "transparent" }}>
                {editingQuotation ? "Edit Quote" : "New Quote"}
              </button>
              <button className={`hw-btn-ghost ${subView === "history" ? "active" : ""}`} onClick={() => setSubView("history")} style={{ border: "none", borderRadius: 0, padding: "8px 16px", background: subView === "history" ? "var(--border)" : "transparent" }}>
                Quote History ({quotations.length})
              </button>
            </div>
          </div>
        }
      />

      {subView === "new" ? (
        <div className="hw-invoice-layout">
          <div className="hw-card">
            <div style={{ display: 'flex', gap: '12px' }}>
              <div style={{ flex: 1.2 }}>
                <Field label="Customer name (optional)">
                  <input value={customerName} onChange={e => setCustomerName(e.target.value)} placeholder="Walk-in customer" />
                </Field>
              </div>
              <div style={{ flex: 1 }}>
                <Field label="Customer email (optional)">
                  <input type="email" value={customerEmail} onChange={e => setCustomerEmail(e.target.value)} placeholder="customer@email.com" />
                </Field>
              </div>
              <div style={{ flex: 1 }}>
                <Field label="Customer phone (optional)">
                  <input type="text" value={customerPhone} onChange={e => setCustomerPhone(e.target.value)} placeholder="e.g. 03328898666" />
                </Field>
              </div>
            </div>

            <div className="hw-search-row" style={{ marginTop: 14, position: "relative" }}>
              <Search size={16} />
              <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search products to add to quotation…" />
              {matches.length > 0 && (
                <div className="hw-autocomplete">
                  {matches.map(p => (
                    <button key={p.id} onClick={() => addToCart(p)}>
                      <span>{p.name}</span>
                      <span className="hw-mono hw-muted">{cs}{fmtNum(sellPrice(p))} · {p.quantity} {p.unit} in stock</span>
                    </button>
                  ))}
                </div>
              )}
            </div>

            {cart.length > 0 && (
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 14, marginBottom: -6, fontSize: '12.5px', color: 'var(--ink-soft)', fontWeight: 600 }}>
                <span>Quote Items ({cart.length})</span>
                <span>Total Units: {cart.reduce((s, i) => s + i.qty, 0)}</span>
              </div>
            )}

            <div className="hw-table-wrap" style={{ marginTop: 14 }}>
              <table className="hw-table hw-table-tight">
                <thead>
                  <tr><th>Item</th><th>Qty</th><th>Cost</th><th>Markup</th><th>Sell price</th><th>Total</th><th></th></tr>
                </thead>
                <tbody>
                  {cartWithPrice.map(i => (
                    <tr key={i.productId}>
                      <td>{i.name}<div className="hw-il-meta">{i.unit}</div></td>
                      <td>
                        <NumInput className="hw-mini-input" min={0} value={i.qty} onChange={n => updateItem(i.productId, { qty: Math.max(0, n) })} />
                      </td>
                      <td>
                        <NumInput className="hw-mini-input" step="0.01" min={0} value={i.costPrice} onChange={n => updateItem(i.productId, { costPrice: Math.max(0, n) })} />
                      </td>
                      <td>
                        <NumInput className="hw-mini-input" step="0.1" value={i.markup} onChange={n => updateItem(i.productId, { markup: n })} style={{ width: 54 }} />%
                      </td>
                      <td>
                        <NumInput className="hw-mini-input" step="0.01" min={0} value={fmtNum(i.sellPrice)} onChange={n => {
                          const newSell = Math.max(0, n);
                          const cost = Number(i.costPrice) || 0;
                          const newMarkup = cost > 0 ? ((newSell - cost) / cost) * 100 : 0;
                          updateItem(i.productId, { markup: newMarkup });
                        }} />
                      </td>
                      <td className="hw-mono">{cs}{fmtNum(i.lineTotal)}</td>
                      <td><button className="hw-icon-btn" onClick={() => removeItem(i.productId)}><X size={14} /></button></td>
                    </tr>
                  ))}
                  {cart.length === 0 && <tr><td colSpan={7}><EmptyRow text="Search above to add items to this quotation estimate." /></td></tr>}
                </tbody>
              </table>
            </div>
          </div>

          <div className="hw-card hw-summary-card">
            <h3>Quote Summary</h3>
            <div className="hw-summary-row"><span>Subtotal</span><span className="hw-mono">{cs}{fmtNum(subtotal)}</span></div>
            <div className="hw-summary-row">
              <span>Discount</span>
              <NumInput className="hw-mini-input" min={0} value={discount} onChange={n => setDiscount(Math.max(0, n))} style={{ width: 70 }} />
            </div>
            <div className="hw-hint" style={{ marginTop: -8, marginBottom: 8 }}>0–100 = percent, above 100 = flat amount off</div>
            <div className="hw-summary-row hw-summary-total"><span>Total</span><span className="hw-mono">{cs}{fmtNum(total)}</span></div>
            <button className="hw-btn-accent hw-btn-block" disabled={cart.length === 0} onClick={saveQuotation}>
              <Save size={15} /> {editingQuotation ? "Update Quotation" : "Save & Preview Quote"}
            </button>
            {editingQuotation && (
              <button className="hw-btn-ghost hw-btn-block" onClick={cancelEdit} style={{ marginTop: '8px' }}>
                Cancel Edit
              </button>
            )}
          </div>
        </div>
      ) : (
        <div className="hw-card" style={{ marginTop: 14 }}>
          <div className="hw-search-row" style={{ marginBottom: 14 }}>
            <Search size={16} />
            <input placeholder="Search past quotations by number or customer name…" value={query} onChange={e => setQuery(e.target.value)} />
          </div>
          <div className="hw-table-wrap">
            <table className="hw-table">
              <thead>
                <tr>
                  <th>Quote Number</th>
                  <th>Customer</th>
                  <th>Date</th>
                  <th>Items</th>
                  <th>Total</th>
                  <th style={{ textAlign: "right" }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {filteredQuotes.map(q => (
                  <tr key={q.id}>
                    <td className="hw-mono" style={{ fontWeight: 600 }}>{q.quotationNumber}</td>
                    <td>{q.customerName}</td>
                    <td className="hw-muted">{fmtDate(q.date)}</td>
                    <td className="hw-muted">{q.items.reduce((sum, item) => sum + item.qty, 0)} items</td>
                    <td className="hw-mono">{cs}{fmtNum(q.total)}</td>
                    <td style={{ textAlign: "right" }}>
                      <div style={{ display: "inline-flex", gap: "8px" }}>
                        <button className="hw-btn-ghost" style={{ padding: "6px 10px", fontSize: "12px", display: "inline-flex", alignItems: "center", gap: "4px" }} onClick={() => {
                          setEditingQuotation(q);
                          setCart(q.items.map(i => ({ productId: i.productId, name: i.name, unit: i.unit, qty: i.qty, costPrice: i.costPrice, markup: i.markup })));
                          setCustomerName(q.customerName === "Walk-in Customer" ? "" : (q.customerName || ""));
                          setCustomerEmail(q.customerEmail || "");
                          setCustomerPhone(q.customerPhone || "");
                          setDiscount(q.discount || 0);
                          setSubView("new");
                          showToast(`Loaded ${q.quotationNumber} for editing`);
                        }}>
                          <Pencil size={12} /> Edit
                        </button>
                        <button className="hw-btn-ghost" style={{ padding: "6px 10px", fontSize: "12px" }} onClick={() => setPreview(q)}>
                          View/Print
                        </button>
                        <button className="hw-btn-accent" style={{ padding: "6px 10px", fontSize: "12px" }} onClick={() => onConvertToInvoice(q.items, q.customerName, q.customerEmail, q.customerPhone || "", q.discount)}>
                          Convert to Invoice
                        </button>
                        <button className="hw-icon-btn" onClick={() => deleteQuotation(q.id)} style={{ color: "var(--danger)" }}>
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
                {filteredQuotes.length === 0 && (
                  <tr><td colSpan={6}><EmptyRow text={query ? "No quotations found." : "No quotations saved yet."} /></td></tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {preview && (
        <QuotationPreviewModal 
          quote={preview} 
          settings={settings} 
          onClose={() => setPreview(null)} 
          onEdit={(q) => {
            setEditingQuotation(q);
            setCart(q.items.map(i => ({ productId: i.productId, name: i.name, unit: i.unit, qty: i.qty, costPrice: i.costPrice, markup: i.markup })));
            setCustomerName(q.customerName === "Walk-in Customer" ? "" : (q.customerName || ""));
            setCustomerEmail(q.customerEmail || "");
            setCustomerPhone(q.customerPhone || "");
            setDiscount(q.discount || 0);
            setSubView("new");
            showToast(`Loaded ${q.quotationNumber} for editing`);
          }}
        />
      )}
    </div>
  );
}

function QuotationPreviewModal({ quote, settings, onClose, onEdit }) {
  const cs = settings.currencySymbol;

  const formatPhoneForWhatsapp = (phone) => {
    let cleaned = phone.replace(/\D/g, "");
    if (cleaned.startsWith("0") && cleaned.length === 11) {
      cleaned = "92" + cleaned.slice(1);
    }
    return cleaned;
  };

  const sendWhatsapp = () => {
    if (!quote.customerPhone) return;
    const phone = formatPhoneForWhatsapp(quote.customerPhone);
    const itemsText = quote.items
      .map(i => `• ${i.name} x ${i.qty} (${i.unit}): ${cs}${fmtNum(i.lineTotal)}`)
      .join("\n");
      
    const message = `Hello *${quote.customerName || "Customer"}*,

Here is your quotation estimate from *${settings.shopName}*:
📄 *Quotation:* ${quote.quotationNumber}
📅 *Date:* ${fmtDateTime(quote.date)}
━━━━━━━━━━━━━━━━━━
${itemsText}
━━━━━━━━━━━━━━━━━━
*Subtotal:* ${cs}${fmtNum(quote.subtotal)}
*Discount:* ${cs}${fmtNum(quote.discount)}
*Estimate Total:* ${cs}${fmtNum(quote.total)}

This is an estimate. Prices are valid for 7 days.
Thank you!`;

    const url = `https://wa.me/${phone}?text=${encodeURIComponent(message)}`;
    window.db.openExternalLink(url);
  };

  return (
    <div className="hw-modal-overlay" onClick={onClose}>
      <div className="hw-modal hw-statement-modal" onClick={e => e.stopPropagation()}>
        <div className="hw-statement-card" id="hw-print-area">
          {/* Customizable Statement Header */}
          <div className="hw-excel-header">
            {/* Custom Uploaded Logo (Only shown if uploaded in Settings) */}
            {settings.logoUrl && (
              <div className="hw-excel-logo-badge">
                <img src={settings.logoUrl} alt="Logo" style={{ maxHeight: 52, maxWidth: 110, objectFit: "contain" }} />
              </div>
            )}

            {/* Main Centered Business Details (100% Dynamic from Settings) */}
            <div className="hw-excel-main-center">
              {settings.shopName && <h1 className="hw-excel-org-title">{settings.shopName}</h1>}
              {settings.shopNameUrdu && (
                <div className="hw-excel-org-urdu">{settings.shopNameUrdu}</div>
              )}
              {(settings.phone || settings.address) && (
                <div className="hw-excel-line-bold">
                  {settings.phone && <span>Tel: {settings.phone}</span>}
                  {settings.phone && settings.address && <span> · </span>}
                  {settings.address && <span>{settings.address}</span>}
                </div>
              )}
              {(settings.paymentDetails || settings.bankDetails) && (
                <div className="hw-excel-line-payment">
                  {settings.paymentDetails || settings.bankDetails}
                </div>
              )}
              {settings.email && (
                <div className="hw-excel-line-email">
                  E-mail: {settings.email}
                </div>
              )}
            </div>

            {/* Centered Document Title */}
            <div className="hw-excel-title-bar">
              <h2 className="hw-excel-doc-title">{settings.quotationTitle || "MATERIAL REQUEST"}</h2>
            </div>
          </div>

          {/* Receiver / Billed-To Info Block */}
          <div className="hw-statement-receiver-block">
            <div className="hw-statement-receiver-info">
              <div className="hw-statement-section-label">M/S &amp; CLIENT DETAILS:</div>
              <div className="hw-statement-customer-name">{quote.customerName || "Walk-in Customer"}</div>
              {quote.customerPhone && <div className="hw-statement-customer-sub">📞 Phone: <strong>{quote.customerPhone}</strong></div>}
              {quote.customerEmail && <div className="hw-statement-customer-sub">✉️ Email: {quote.customerEmail}</div>}
            </div>
            <div className="hw-statement-doc-info">
              <div className="hw-statement-doc-row">
                <span className="hw-statement-doc-label">Quote Number:</span>
                <span className="hw-statement-doc-val">{quote.quotationNumber}</span>
              </div>
              <div className="hw-statement-doc-row">
                <span className="hw-statement-doc-label">Date &amp; Time:</span>
                <span className="hw-statement-doc-val">{fmtDateTime(quote.date)}</span>
              </div>
              <div className="hw-statement-doc-row">
                <span className="hw-statement-doc-label">Estimate Validity:</span>
                <span className="hw-statement-doc-val" style={{ color: "#EA580C", fontWeight: 700 }}>7 Days</span>
              </div>
            </div>
          </div>

          {/* Summary Box */}
          <div className="hw-statement-summary-box">
            <div className="hw-statement-summary-col">
              <span className="hw-statement-summary-label">Subtotal</span>
              <span className="hw-statement-summary-value">{cs}{fmtNum(quote.subtotal)}</span>
            </div>
            <div className="hw-statement-summary-col">
              <span className="hw-statement-summary-label">Discount (-)</span>
              <span className="hw-statement-summary-value red">{cs}{fmtNum(quote.discount)}</span>
            </div>
            <div className="hw-statement-summary-col">
              <span className="hw-statement-summary-label">Estimate Total</span>
              <span className="hw-statement-summary-value green">{cs}{fmtNum(quote.total)}</span>
              <span className="hw-statement-summary-sub">(quotation only)</span>
            </div>
            <div className="hw-statement-summary-col">
              <span className="hw-statement-summary-label">Quote Number</span>
              <span className="hw-statement-summary-value" style={{ fontSize: "14px" }}>{quote.quotationNumber}</span>
            </div>
          </div>

          {/* Entries Count Info */}
          <div className="hw-statement-entries-info">
            <span>No. of Items: {quote.items.length} (Total {quote.items.reduce((s, i) => s + i.qty, 0)} units)</span>
            {quote.items.length > 8 && <span style={{ fontSize: '11px', color: 'var(--accent)', fontWeight: 600 }}>Scroll down inside statement to view all items ↓</span>}
          </div>

          {/* Items Table */}
          <div className="hw-statement-table-wrap">
            <table className="hw-statement-table">
              <thead>
                <tr>
                  <th style={{ width: "40px" }}>#</th>
                  <th>Details</th>
                  <th style={{ textAlign: "right" }}>Qty</th>
                  <th style={{ textAlign: "right" }}>Price</th>
                  <th style={{ textAlign: "right" }}>Amount</th>
                </tr>
              </thead>
              <tbody>
                {quote.items.map((i, idx) => (
                  <tr key={idx}>
                    <td>{idx + 1}</td>
                    <td>
                      <span style={{ fontWeight: 600, color: "#111827" }}>{i.name}</span>
                      <span style={{ fontSize: "11px", color: "#6B7280", marginLeft: "6px" }}>({i.unit})</span>
                    </td>
                    <td className="mono" style={{ textAlign: "right" }}>{i.qty}</td>
                    <td className="mono" style={{ textAlign: "right" }}>{cs}{fmtNum(i.price)}</td>
                    <td className="mono green" style={{ textAlign: "right" }}>{cs}{fmtNum(i.lineTotal)}</td>
                  </tr>
                ))}
                {/* Grand Total Row */}
                <tr className="hw-statement-table-totals">
                  <td colSpan={2}>Grand Total</td>
                  <td style={{ textAlign: "right" }}>{quote.items.reduce((s, i) => s + i.qty, 0)}</td>
                  <td></td>
                  <td className="green" style={{ textAlign: "right" }}>{cs}{fmtNum(quote.total)}</td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* Footer Banner */}
          <div className="hw-statement-footer">
            <div className="hw-statement-footer-left">
              <span>{settings.shopName}</span>
              {settings.phone && <span>· Ph: {settings.phone}</span>}
              {settings.whatsapp && <span>· WA: {settings.whatsapp}</span>}
            </div>
            <div className="hw-statement-footer-right">
              <span>Prices are valid for 7 days. Thank you for your business!</span>
            </div>
          </div>
        </div>

        {/* Action buttons */}
        <div className="hw-modal-actions hw-no-print" style={{ padding: "16px 24px", background: "#F9FAFB", borderTop: "1px solid #E5E7EB" }}>
          {onEdit && (
            <button className="hw-btn-ghost" onClick={() => { onClose(); onEdit(quote); }}>
              <Pencil size={14} /> Edit Quotation
            </button>
          )}
          <div style={{ flex: 1 }} />
          {quote.customerPhone && (
            <button className="hw-btn-accent" onClick={sendWhatsapp} style={{ background: "#25D366", borderColor: "#25D366" }}>
              Share WhatsApp
            </button>
          )}
          <button className="hw-btn-ghost" onClick={onClose}>Close</button>
          <button className="hw-btn-accent" onClick={() => window.print()}><Printer size={14} /> Print / Export</button>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   INVOICE HISTORY
--------------------------------------------------------- */
function InvoiceHistory({ invoices, settings, persistInvoices, persistProducts, products, showToast, onEditInvoice }) {
  const [query, setQuery] = useState("");
  const [preview, setPreview] = useState(null);
  const cs = settings.currencySymbol;

  const sorted = [...invoices].filter(i =>
    i.invoiceNumber.toLowerCase().includes(query.toLowerCase()) ||
    (i.customerName || "").toLowerCase().includes(query.toLowerCase())
  ).sort((a, b) => new Date(b.date) - new Date(a.date));

  const deleteInvoice = async (invoice, restock) => {
    if (restock) {
      const nextProducts = products.map(p => {
        const item = invoice.items.find(i => i.productId === p.id);
        return item ? { ...p, quantity: p.quantity + item.qty } : p;
      });
      await persistProducts(nextProducts);
    }
    await persistInvoices(invoices.filter(i => i.id !== invoice.id));
    setPreview(null);
    showToast(restock ? "Invoice deleted, stock restored" : "Invoice deleted");
  };

  return (
    <div className="hw-view">
      <ViewHeader eyebrow={`${invoices.length} invoice${invoices.length === 1 ? "" : "s"}`} title="Invoice History" />
      <div className="hw-search-row">
        <Search size={16} />
        <input placeholder="Search by invoice number or customer…" value={query} onChange={e => setQuery(e.target.value)} />
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(290px, 1fr))', gap: '16px', marginTop: '16px' }}>
        {sorted.map(inv => (
          <div key={inv.id} className="hw-invoice-card" onClick={() => setPreview(inv)}>
            <div className="hw-invoice-card-header">
              <span className="hw-invoice-card-number">#{inv.invoiceNumber}</span>
              {onEditInvoice && (
                <button
                  className="hw-btn-ghost"
                  style={{ padding: "3px 8px", fontSize: "11px", display: "inline-flex", alignItems: "center", gap: "3px" }}
                  onClick={(e) => { e.stopPropagation(); onEditInvoice(inv); }}
                >
                  <Pencil size={11} /> Edit
                </button>
              )}
            </div>
            <div className="hw-invoice-card-body">
              <div className="hw-invoice-card-details">
                <span className="hw-invoice-card-customer">{inv.customerName || "Walk-in Customer"}</span>
                <span className="hw-invoice-card-date">Date: {fmtDate(inv.date)}</span>
              </div>
              <div className="hw-invoice-card-right">
                <span className="hw-invoice-card-total">{cs}{fmtNum(inv.total)}</span>
                <span style={{ fontSize: '11px', color: 'var(--ink-soft)', display: 'flex', alignItems: 'center', gap: '2px', fontWeight: 500 }}>
                  {inv.items.length} items <ChevronRight size={13} />
                </span>
              </div>
            </div>
          </div>
        ))}
        {sorted.length === 0 && (
          <div style={{ gridColumn: '1 / -1' }}>
            <EmptyRow text="No invoices found." />
          </div>
        )}
      </div>
      {preview && (
        <InvoicePreviewModal 
          invoice={preview} 
          settings={settings} 
          onClose={() => setPreview(null)} 
          onDelete={(restock) => deleteInvoice(preview, restock)} 
          onEdit={onEditInvoice}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   DAILY LEDGER — sales vs expenses, day by day
--------------------------------------------------------- */
const EXPENSE_CATEGORIES = ["Rent", "Wages", "Transport", "Utilities", "Supplies", "Misc"];

function DailyLedger({ invoices, expenses, settings, persistExpenses, showToast, products }) {
  const [day, setDay] = useState(todayKey());
  const [editing, setEditing] = useState(null); // entry or "new"
  const [previewInvoice, setPreviewInvoice] = useState(null);
  const cs = settings.currencySymbol;

  const invoicesForDay = useMemo(
    () => invoices.filter(inv => toDateKey(new Date(inv.date)) === day),
    [invoices, day]
  );
  const entriesForDay = useMemo(
    () => expenses.filter(e => e.date === day).sort((a, b) => (b.createdAt || "").localeCompare(a.createdAt || "")),
    [expenses, day]
  );

  const salesTotal = invoicesForDay.reduce((s, i) => s + i.total, 0);
  const cogsTotal = invoicesForDay.reduce((s, i) => s + (i.totalCost || 0), 0);
  const grossProfit = salesTotal - cogsTotal;
  
  // Custom income and expenses
  const customIncomes = entriesForDay.filter(e => e.type === "income");
  const customExpenses = entriesForDay.filter(e => e.type === "expense" || !e.type);
  
  const customIncomeTotal = customIncomes.reduce((s, e) => s + e.amount, 0);
  const expensesTotal = customExpenses.reduce((s, e) => s + e.amount, 0);
  
  const totalRevenue = salesTotal + customIncomeTotal;
  const netProfit = totalRevenue - cogsTotal - expensesTotal;

  const shiftDay = (delta) => {
    const d = new Date(day + "T00:00:00");
    d.setDate(d.getDate() + delta);
    setDay(toDateKey(d));
  };

  const upsertExpense = async (exp) => {
    let next;
    if (expenses.some(e => e.id === exp.id)) {
      next = expenses.map(e => e.id === exp.id ? exp : e);
    } else {
      next = [...expenses, { ...exp, createdAt: new Date().toISOString() }];
    }
    await persistExpenses(next);
    setEditing(null);
    showToast("Ledger entry saved");
  };
  const deleteExpense = async (id) => {
    await persistExpenses(expenses.filter(e => e.id !== id));
    setEditing(null);
    showToast("Ledger entry removed");
  };

  return (
    <div className="hw-view">
      <ViewHeader
        eyebrow="Sales, costs & spending"
        title="Daily Ledger"
        right={<button className="hw-btn-accent" onClick={() => setEditing("new")}><Plus size={16} /> Add entry</button>}
      />

      <div className="hw-day-nav">
        <button className="hw-icon-btn" onClick={() => shiftDay(-1)}><ChevronLeft size={16} /></button>
        <div className="hw-day-label"><CalendarDays size={14} /> {fmtDayLabel(day)}</div>
        <button className="hw-icon-btn" onClick={() => shiftDay(1)} disabled={day >= todayKey()}><ChevronRight size={16} /></button>
        <input type="date" value={day} max={todayKey()} onChange={e => setDay(e.target.value)} className="hw-date-input" />
        {day !== todayKey() && <button className="hw-link" onClick={() => setDay(todayKey())}>Jump to today</button>}
      </div>

      <div className="hw-stat-grid">
        <StatCard 
          label="Total revenue" 
          value={`${cs}${fmtNum(totalRevenue)}`} 
          sub={customIncomeTotal > 0 ? `Invoiced: ${cs}${fmtNum(salesTotal)} · Other: ${cs}${fmtNum(customIncomeTotal)}` : `${invoicesForDay.length} invoice${invoicesForDay.length === 1 ? "" : "s"}`} 
          icon={<Receipt size={18} />} 
          accent="accent" 
        />
        <StatCard label="Cost of goods sold" value={`${cs}${fmtNum(cogsTotal)}`} icon={<ArrowDownCircle size={18} />} accent="steel" />
        <StatCard label="Expenses" value={`${cs}${fmtNum(expensesTotal)}`} sub={`${customExpenses.length} expense${customExpenses.length === 1 ? "" : "s"}`} icon={<Wallet size={18} />} accent={expensesTotal > 0 ? "warn" : "default"} />
        <StatCard label="Net for the day" value={`${cs}${fmtNum(netProfit)}`} sub="revenue − cost − expenses" icon={<ArrowUpCircle size={18} />} accent={netProfit >= 0 ? "ok" : "warn"} />
      </div>

      <div className="hw-card">
        <div className="hw-card-head"><h3>Custom ledger entries on {fmtDayLabel(day).toLowerCase()}</h3></div>
        <div className="hw-table-wrap" style={{ border: "none" }}>
          <table className="hw-table hw-table-tight">
            <thead><tr><th>Description</th><th>Type</th><th>Category</th><th>Amount</th><th></th></tr></thead>
            <tbody>
              {entriesForDay.map(e => {
                const isIncome = e.type === "income";
                return (
                  <tr key={e.id}>
                    <td>{e.description}</td>
                    <td>
                      <span style={{ 
                        display: "inline-block", 
                        fontSize: "11px", 
                        fontWeight: 600, 
                        padding: "2px 6px", 
                        borderRadius: "4px",
                        background: isIncome ? "var(--success-tint)" : "var(--danger-tint)",
                        color: isIncome ? "var(--success)" : "var(--danger)"
                      }}>
                        {isIncome ? "Income (Sold)" : "Expense (Spent)"}
                      </span>
                    </td>
                    <td className="hw-muted">{e.category || "—"}</td>
                    <td className="hw-mono" style={{ 
                      fontWeight: 600, 
                      color: isIncome ? "var(--success)" : "var(--danger)"
                    }}>
                      {isIncome ? "+" : "-"}{cs}{fmtNum(e.amount)}
                    </td>
                    <td>
                      <button className="hw-icon-btn" onClick={() => setEditing(e)}><Pencil size={14} /></button>
                      <button className="hw-icon-btn" onClick={() => deleteExpense(e.id)}><Trash2 size={14} /></button>
                    </td>
                  </tr>
                );
              })}
              {entriesForDay.length === 0 && <tr><td colSpan={5}><EmptyRow text="No custom entries logged for this day yet." /></td></tr>}
            </tbody>
          </table>
        </div>
      </div>

      {invoicesForDay.length > 0 && (
        <div style={{ marginTop: 18 }}>
          <div className="hw-card-head" style={{ marginBottom: 12 }}>
            <h3>Invoiced Sales ({fmtDayLabel(day)})</h3>
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(290px, 1fr))', gap: '14px' }}>
            {invoicesForDay.map(inv => (
              <div key={inv.id} className="hw-invoice-card" onClick={() => setPreviewInvoice(inv)}>
                <div className="hw-invoice-card-header">
                  <span className="hw-invoice-card-number">#{inv.invoiceNumber}</span>
                </div>
                <div className="hw-invoice-card-body">
                  <div className="hw-invoice-card-details">
                    <span className="hw-invoice-card-customer">{inv.customerName || "Walk-in Customer"}</span>
                    <span className="hw-invoice-card-date">{fmtDateTime(inv.date)}</span>
                  </div>
                  <div className="hw-invoice-card-right">
                    <span className="hw-invoice-card-total">{cs}{fmtNum(inv.total)}</span>
                    <span style={{ fontSize: '11px', color: 'var(--ink-soft)', fontWeight: 500 }}>
                      {inv.items.length} items
                    </span>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {editing && (
        <ExpenseModal
          expense={editing === "new" ? null : editing}
          defaultDay={day}
          onClose={() => setEditing(null)}
          onSave={upsertExpense}
          onDelete={editing !== "new" ? () => deleteExpense(editing.id) : null}
          products={products}
          settings={settings}
        />
      )}
      {previewInvoice && (
        <InvoicePreviewModal invoice={previewInvoice} settings={settings} onClose={() => setPreviewInvoice(null)} onDelete={null} />
      )}
    </div>
  );
}

function ExpenseModal({ expense, defaultDay, onClose, onSave, onDelete, products = [], settings }) {
  const [form, setForm] = useState(expense || { id: uid(), date: defaultDay, type: "expense", description: "", category: "Misc", amount: 0 });
  const [showSuggestions, setShowSuggestions] = useState(false);
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const valid = form.description.trim().length > 0 && form.amount > 0;
  
  const type = form.type || "expense";
  const categories = type === "expense" ? EXPENSE_CATEGORIES : ["Other Sales", "Services", "Commission", "Return/Refund", "Misc"];

  // Filter items in inventory matching the typed description
  const matches = form.description.trim() ? products.filter(p => 
    p.name.toLowerCase().includes(form.description.toLowerCase())
  ).slice(0, 5) : [];

  return (
    <ModalShell onClose={onClose} title={expense ? "Edit ledger entry" : "Add ledger entry"}>
      <div className="hw-form-grid">
        <Field label="Entry type" span={2}>
          <div style={{ display: "flex", gap: 10 }}>
            <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", flex: 1, padding: "8px 12px", border: "1px solid var(--border)", borderRadius: 6, background: type === "expense" ? "var(--danger-tint)" : "var(--surface-alt)", color: type === "expense" ? "var(--danger)" : "var(--ink-soft)", fontWeight: 600 }}>
              <input type="radio" name="entry-type" checked={type === "expense"} onChange={() => { setForm(f => ({ ...f, type: "expense", category: "Misc" })); }} style={{ cursor: "pointer" }} />
              Expense (Spent)
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", flex: 1, padding: "8px 12px", border: "1px solid var(--border)", borderRadius: 6, background: type === "income" ? "var(--success-tint)" : "var(--surface-alt)", color: type === "income" ? "var(--success)" : "var(--ink-soft)", fontWeight: 600 }}>
              <input type="radio" name="entry-type" checked={type === "income"} onChange={() => { setForm(f => ({ ...f, type: "income", category: "Misc" })); }} style={{ cursor: "pointer" }} />
              Income (Sold)
            </label>
          </div>
        </Field>
        
        <Field label="Description" span={2}>
          <div style={{ position: "relative" }}>
            <input 
              autoFocus 
              value={form.description} 
              onChange={e => {
                set("description", e.target.value);
                setShowSuggestions(true);
              }} 
              onFocus={() => setShowSuggestions(true)}
              onBlur={() => setTimeout(() => setShowSuggestions(false), 200)}
              placeholder={type === "expense" ? "e.g. Truck fuel, wages, or search product..." : "e.g. Service charge, or search product..."} 
            />
            {showSuggestions && matches.length > 0 && (
              <div className="hw-autocomplete" style={{ top: 38, left: 0, right: 0 }}>
                {matches.map(p => {
                  const sp = sellPrice(p);
                  const priceLabel = type === "income" ? `Sell: ${settings.currencySymbol}${fmtNum(sp)}` : `Cost: ${settings.currencySymbol}${fmtNum(p.costPrice)}`;
                  return (
                    <button 
                      key={p.id} 
                      type="button"
                      onMouseDown={() => {
                        setForm(f => ({
                          ...f,
                          description: p.name,
                          category: p.category || "Misc",
                          amount: type === "income" ? sp : p.costPrice
                        }));
                        setShowSuggestions(false);
                      }}
                      style={{ 
                        width: "100%", 
                        textAlign: "left", 
                        padding: "9px 12px", 
                        border: "none", 
                        borderBottom: "1px solid var(--border)",
                        background: "none", 
                        cursor: "pointer", 
                        display: "flex", 
                        justifyContent: "space-between",
                        fontFamily: "var(--font-body)",
                        fontSize: "13px"
                      }}
                    >
                      <span style={{ fontWeight: 500, color: "var(--ink)" }}>{p.name}</span>
                      <span style={{ fontSize: "11px", color: "var(--ink-soft)" }}>{p.category} · {priceLabel}</span>
                    </button>
                  );
                })}
              </div>
            )}
          </div>
        </Field>

        <Field label="Category">
          <input list="hw-ledger-categories" value={form.category} onChange={e => set("category", e.target.value)} />
          <datalist id="hw-ledger-categories">
            {categories.map(c => <option key={c} value={c} />)}
          </datalist>
        </Field>
        <Field label="Amount">
          <NumInput step="0.01" value={form.amount} onChange={n => set("amount", Math.max(0, n))} />
        </Field>
        <Field label="Date" span={2}>
          <input type="date" value={form.date} max={todayKey()} onChange={e => set("date", e.target.value)} />
        </Field>
      </div>
      <div className="hw-modal-actions">
        {onDelete && <button className="hw-btn-danger" onClick={onDelete}><Trash2 size={14} /> Delete</button>}
        <div style={{ flex: 1 }} />
        <button className="hw-btn-ghost" onClick={onClose}>Cancel</button>
        <button className="hw-btn-accent" disabled={!valid} onClick={() => onSave(form)}><Save size={14} /> Save entry</button>
      </div>
    </ModalShell>
  );
}


/* ---------------------------------------------------------
   INVOICE PREVIEW / PRINT
--------------------------------------------------------- */
function InvoicePreviewModal({ invoice, settings, onClose, onDelete, onEdit }) {
  const cs = settings.currencySymbol;
  const [confirmDelete, setConfirmDelete] = useState(false);

  const formatPhoneForWhatsapp = (phone) => {
    let cleaned = phone.replace(/\D/g, "");
    if (cleaned.startsWith("0") && cleaned.length === 11) {
      cleaned = "92" + cleaned.slice(1);
    }
    return cleaned;
  };

  const sendWhatsapp = () => {
    if (!invoice.customerPhone) return;
    const phone = formatPhoneForWhatsapp(invoice.customerPhone);
    const itemsText = invoice.items
      .map(i => `• ${i.name} x ${i.qty} (${i.unit}): ${cs}${fmtNum(i.lineTotal)}`)
      .join("\n");
      
    const message = `Hello *${invoice.customerName || "Customer"}*,

Here is your invoice statement from *${settings.shopName}*:
📄 *Invoice:* ${invoice.invoiceNumber}
📅 *Date:* ${fmtDateTime(invoice.date)}
━━━━━━━━━━━━━━━━━━
${itemsText}
━━━━━━━━━━━━━━━━━━
*Subtotal:* ${cs}${fmtNum(invoice.subtotal)}
*Discount:* ${cs}${fmtNum(invoice.discount)}
*Total Amount:* ${cs}${fmtNum(invoice.total)}

Thank you for your business!`;

    const url = `https://wa.me/${phone}?text=${encodeURIComponent(message)}`;
    window.db.openExternalLink(url);
  };

  return (
    <div className="hw-modal-overlay" onClick={onClose}>
      <div className="hw-modal hw-statement-modal" onClick={e => e.stopPropagation()}>
        <div className="hw-statement-card" id="hw-print-area">
          {/* Customizable Statement Header */}
          <div className="hw-excel-header">
            {/* Custom Uploaded Logo (Only shown if uploaded in Settings) */}
            {settings.logoUrl && (
              <div className="hw-excel-logo-badge">
                <img src={settings.logoUrl} alt="Logo" style={{ maxHeight: 52, maxWidth: 110, objectFit: "contain" }} />
              </div>
            )}

            {/* Main Centered Business Details (100% Dynamic from Settings) */}
            <div className="hw-excel-main-center">
              {settings.shopName && <h1 className="hw-excel-org-title">{settings.shopName}</h1>}
              {settings.shopNameUrdu && (
                <div className="hw-excel-org-urdu">{settings.shopNameUrdu}</div>
              )}
              {(settings.phone || settings.address) && (
                <div className="hw-excel-line-bold">
                  {settings.phone && <span>Tel: {settings.phone}</span>}
                  {settings.phone && settings.address && <span> · </span>}
                  {settings.address && <span>{settings.address}</span>}
                </div>
              )}
              {(settings.paymentDetails || settings.bankDetails) && (
                <div className="hw-excel-line-payment">
                  {settings.paymentDetails || settings.bankDetails}
                </div>
              )}
              {settings.email && (
                <div className="hw-excel-line-email">
                  E-mail: {settings.email}
                </div>
              )}
            </div>

            {/* Centered Document Title */}
            <div className="hw-excel-title-bar">
              <h2 className="hw-excel-doc-title">{settings.invoiceTitle || "TAX INVOICE"}</h2>
            </div>
          </div>

          {/* Receiver / Billed-To Info Block */}
          <div className="hw-statement-receiver-block">
            <div className="hw-statement-receiver-info">
              <div className="hw-statement-section-label">M/S &amp; CLIENT DETAILS:</div>
              <div className="hw-statement-customer-name">{invoice.customerName || "Walk-in Customer"}</div>
              {invoice.customerPhone && <div className="hw-statement-customer-sub">📞 Phone: <strong>{invoice.customerPhone}</strong></div>}
              {invoice.customerEmail && <div className="hw-statement-customer-sub">✉️ Email: {invoice.customerEmail}</div>}
            </div>
            <div className="hw-statement-doc-info">
              <div className="hw-statement-doc-row">
                <span className="hw-statement-doc-label">Invoice Number:</span>
                <span className="hw-statement-doc-val">{invoice.invoiceNumber}</span>
              </div>
              <div className="hw-statement-doc-row">
                <span className="hw-statement-doc-label">Date &amp; Time:</span>
                <span className="hw-statement-doc-val">{fmtDateTime(invoice.date)}</span>
              </div>
              <div className="hw-statement-doc-row">
                <span className="hw-statement-doc-label">Payment Status:</span>
                <span className="hw-statement-doc-val" style={{ color: "#16A34A", fontWeight: 700 }}>Settled (Paid)</span>
              </div>
            </div>
          </div>

          {/* Summary Box */}
          <div className="hw-statement-summary-box">
            <div className="hw-statement-summary-col">
              <span className="hw-statement-summary-label">Subtotal</span>
              <span className="hw-statement-summary-value">{cs}{fmtNum(invoice.subtotal)}</span>
            </div>
            <div className="hw-statement-summary-col">
              <span className="hw-statement-summary-label">Discount (-)</span>
              <span className="hw-statement-summary-value red">{cs}{fmtNum(invoice.discount)}</span>
            </div>
            <div className="hw-statement-summary-col">
              <span className="hw-statement-summary-label">Net Balance</span>
              <span className="hw-statement-summary-value green">{cs}{fmtNum(invoice.total)}</span>
              <span className="hw-statement-summary-sub">(settled)</span>
            </div>
            <div className="hw-statement-summary-col">
              <span className="hw-statement-summary-label">Ref Number</span>
              <span className="hw-statement-summary-value" style={{ fontSize: "14px" }}>{invoice.invoiceNumber}</span>
            </div>
          </div>

          {/* Entries Count Info */}
          <div className="hw-statement-entries-info">
            <span>No. of Items: {invoice.items.length} (Total {invoice.items.reduce((s, i) => s + i.qty, 0)} units)</span>
            {invoice.items.length > 8 && <span style={{ fontSize: '11px', color: 'var(--accent)', fontWeight: 600 }}>Scroll down inside statement to view all items ↓</span>}
          </div>

          {/* Items Table */}
          <div className="hw-statement-table-wrap">
            <table className="hw-statement-table">
              <thead>
                <tr>
                  <th style={{ width: "40px" }}>#</th>
                  <th>Details</th>
                  <th style={{ textAlign: "right" }}>Qty</th>
                  <th style={{ textAlign: "right" }}>Price</th>
                  <th style={{ textAlign: "right" }}>Amount</th>
                </tr>
              </thead>
              <tbody>
                {invoice.items.map((i, idx) => (
                  <tr key={idx}>
                    <td>{idx + 1}</td>
                    <td>
                      <span style={{ fontWeight: 600, color: "#111827" }}>{i.name}</span>
                      <span style={{ fontSize: "11px", color: "#6B7280", marginLeft: "6px" }}>({i.unit})</span>
                    </td>
                    <td className="mono" style={{ textAlign: "right" }}>{i.qty}</td>
                    <td className="mono" style={{ textAlign: "right" }}>{cs}{fmtNum(i.price)}</td>
                    <td className="mono green" style={{ textAlign: "right" }}>{cs}{fmtNum(i.lineTotal)}</td>
                  </tr>
                ))}
                {/* Grand Total Row */}
                <tr className="hw-statement-table-totals">
                  <td colSpan={2}>Grand Total</td>
                  <td style={{ textAlign: "right" }}>{invoice.items.reduce((s, i) => s + i.qty, 0)}</td>
                  <td></td>
                  <td className="green" style={{ textAlign: "right" }}>{cs}{fmtNum(invoice.total)}</td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* Footer Banner */}
          <div className="hw-statement-footer">
            <div className="hw-statement-footer-left">
              <span>{settings.shopName}</span>
              {settings.phone && <span>· Ph: {settings.phone}</span>}
              {settings.whatsapp && <span>· WA: {settings.whatsapp}</span>}
            </div>
            <div className="hw-statement-footer-right">
              <span>Thank you for your valued business!</span>
            </div>
          </div>
        </div>

        {/* Action buttons */}
        <div className="hw-modal-actions hw-no-print" style={{ padding: "16px 24px", background: "#F9FAFB", borderTop: "1px solid #E5E7EB" }}>
          {onDelete && !confirmDelete && (
            <button className="hw-btn-danger" onClick={() => setConfirmDelete(true)}><Trash2 size={14} /> Delete</button>
          )}
          {confirmDelete && (
            <>
              <button className="hw-btn-danger" onClick={() => onDelete(true)}>Delete &amp; restock items</button>
              <button className="hw-btn-ghost" onClick={() => onDelete(false)}>Delete only</button>
            </>
          )}
          {onEdit && (
            <button className="hw-btn-ghost" onClick={() => { onClose(); onEdit(invoice); }}>
              <Pencil size={14} /> Edit Invoice
            </button>
          )}
          <div style={{ flex: 1 }} />
          {invoice.customerPhone && (
            <button className="hw-btn-accent" onClick={sendWhatsapp} style={{ background: "#25D366", borderColor: "#25D366" }}>
              Share WhatsApp
            </button>
          )}
          <button className="hw-btn-ghost" onClick={onClose}>Close</button>
          <button className="hw-btn-accent" onClick={() => window.print()}><Printer size={14} /> Print / Export</button>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   SETTINGS MODAL
--------------------------------------------------------- */
const POPULAR_BANKS = [
  "Meezan Bank",
  "JazzCash",
  "EasyPaisa",
  "Bank Alfalah",
  "HBL (Habib Bank Limited)",
  "UBL (United Bank Limited)",
  "MCB Bank",
  "Allied Bank (ABL)",
  "Faysal Bank",
  "Bank of Punjab (BOP)",
  "Askari Bank",
  "Standard Chartered",
  "SadaPay",
  "NayaPay",
  "Other Bank / Custom"
];

function SettingsModal({ settings, onClose, onSave }) {
  const [form, setForm] = useState(settings);
  const [bank1Name, setBank1Name] = useState("JazzCash");
  const [bank1Number, setBank1Number] = useState("");
  const [bank1Title, setBank1Title] = useState("");
  
  const [bank2Name, setBank2Name] = useState("Meezan Bank");
  const [bank2Number, setBank2Number] = useState("");
  const [bank2Title, setBank2Title] = useState("");

  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  const handleLogoUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.size > 2 * 1024 * 1024) {
      alert("Please select an image smaller than 2MB.");
      return;
    }
    const reader = new FileReader();
    reader.onload = (ev) => {
      set("logoUrl", ev.target.result);
    };
    reader.readAsDataURL(file);
  };

  const updateCompiledPayment = (b1N, b1Num, b1T, b2N, b2Num, b2T) => {
    const parts = [];
    if (b1Num.trim()) {
      parts.push(`${b1N.toUpperCase()}: ${b1Num.trim()}${b1T.trim() ? ` ${b1T.trim().toUpperCase()}` : ""}`);
    }
    if (b2Num.trim()) {
      parts.push(`${b2N.toUpperCase()}: ${b2Num.trim()}${b2T.trim() ? ` ${b2T.trim().toUpperCase()}` : ""}`);
    }
    if (parts.length > 0) {
      const compiled = parts.join(". ");
      set("paymentDetails", compiled);
      set("bankDetails", compiled);
    }
  };

  const handleSave = () => {
    let finalPayment = form.paymentDetails || form.bankDetails || "";
    if (bank1Number.trim() || bank2Number.trim()) {
      const parts = [];
      if (bank1Number.trim()) {
        parts.push(`${bank1Name.toUpperCase()}: ${bank1Number.trim()}${bank1Title.trim() ? ` ${bank1Title.trim().toUpperCase()}` : ""}`);
      }
      if (bank2Number.trim()) {
        parts.push(`${bank2Name.toUpperCase()}: ${bank2Number.trim()}${bank2Title.trim() ? ` ${bank2Title.trim().toUpperCase()}` : ""}`);
      }
      if (parts.length > 0) {
        finalPayment = parts.join(". ");
      }
    }
    const finalForm = {
      ...form,
      paymentDetails: finalPayment,
      bankDetails: finalPayment
    };
    onSave(finalForm);
  };

  return (
    <ModalShell onClose={onClose} title="Organization & Profile Settings">
      <div className="hw-form-grid">
        {/* Custom Logo Upload Option */}
        <div style={{ gridColumn: "span 2", background: "#F8FAFC", border: "1px solid #E2E8F0", borderRadius: 10, padding: "12px 16px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 16 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            {form.logoUrl ? (
              <img src={form.logoUrl} alt="Logo" style={{ width: 50, height: 50, objectFit: "contain", borderRadius: 8, border: "1px solid #CBD5E1", background: "#FFFFFF" }} />
            ) : (
              <div style={{ width: 50, height: 50, borderRadius: 8, border: "1.5px dashed #94A3B8", display: "flex", alignItems: "center", justifyContent: "center", color: "#94A3B8", fontSize: "11px", fontWeight: 600, background: "#FFFFFF" }}>
                No Logo
              </div>
            )}
            <div>
              <div style={{ fontSize: "13px", fontWeight: 700, color: "#1E293B" }}>Custom Organization Logo</div>
              <div style={{ fontSize: "11.5px", color: "#64748B" }}>Upload your logo to appear on Quotations &amp; Invoices.</div>
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <label className="hw-btn-accent" style={{ height: 34, padding: "0 12px", cursor: "pointer", fontSize: "12px", display: "inline-flex", alignItems: "center", gap: 6 }}>
              <Upload size={13} /> {form.logoUrl ? "Change Logo" : "Upload Logo"}
              <input type="file" accept="image/*" onChange={handleLogoUpload} style={{ display: "none" }} />
            </label>
            {form.logoUrl && (
              <button type="button" className="hw-btn-ghost" onClick={() => set("logoUrl", "")} style={{ height: 34, padding: "0 10px", fontSize: "12px", color: "#DC2626" }}>
                Remove
              </button>
            )}
          </div>
        </div>

        <Field label="Organization / English Name">
          <input value={form.shopName || ""} onChange={e => set("shopName", e.target.value)} placeholder="e.g. CAPITAL HARDWARE TRADING" />
        </Field>
        <Field label="Urdu Name (Optional)">
          <input value={form.shopNameUrdu || ""} onChange={e => set("shopNameUrdu", e.target.value)} placeholder="e.g. کیپیٹل ہارڈ و ئیر ٹریڈنگ" style={{ direction: "rtl", fontFamily: "serif" }} />
        </Field>
        <Field label="Phone Number">
          <input value={form.phone || ""} onChange={e => set("phone", e.target.value)} placeholder="e.g. 0332-8898666" />
        </Field>
        <Field label="Email Address">
          <input value={form.email || ""} onChange={e => set("email", e.target.value)} placeholder="e.g. capitaalht@gmail.com" />
        </Field>
        <Field label="WhatsApp Number">
          <input value={form.whatsapp || ""} onChange={e => set("whatsapp", e.target.value)} placeholder="e.g. 0332-8898666" />
        </Field>
        <Field label="Shop Address / Location">
          <input value={form.address || ""} onChange={e => set("address", e.target.value)} placeholder="e.g. I-16 Islamabad" />
        </Field>

        {/* Single Unified Professional Bank & Payment Details Container */}
        <div style={{ gridColumn: "span 2", background: "#F8FAFC", border: "1px solid #E2E8F0", borderRadius: 10, padding: "14px 16px", marginTop: 4, display: "flex", flexDirection: "column", gap: 12 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", borderBottom: "1px solid #E2E8F0", paddingBottom: 8 }}>
            <span style={{ fontSize: "12px", fontWeight: 700, color: "#1E293B", textTransform: "uppercase", letterSpacing: "0.5px" }}>
              🏦 Online Banking &amp; Payment Accounts
            </span>
            <span style={{ fontSize: "11px", color: "#64748B" }}>
              Shown on Quotations &amp; Invoices
            </span>
          </div>

          {/* Bank Account 1 */}
          <div style={{ background: "#FFFFFF", border: "1px solid #E2E8F0", borderRadius: 8, padding: "12px 14px", display: "flex", flexDirection: "column", gap: 10 }}>
            <div style={{ fontSize: "11px", fontWeight: 700, color: "#EA580C", textTransform: "uppercase", letterSpacing: "0.5px" }}>
              Primary Account (JazzCash / EasyPaisa / Bank)
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <Field label="Bank / Wallet Provider">
                <select value={bank1Name} onChange={e => { setBank1Name(e.target.value); updateCompiledPayment(e.target.value, bank1Number, bank1Title, bank2Name, bank2Number, bank2Title); }}>
                  {POPULAR_BANKS.map(b => (
                    <option key={b} value={b}>{b}</option>
                  ))}
                </select>
              </Field>
              <Field label="Account Title (optional)">
                <input 
                  value={bank1Title} 
                  onChange={e => { setBank1Title(e.target.value); updateCompiledPayment(bank1Name, bank1Number, e.target.value, bank2Name, bank2Number, bank2Title); }} 
                  placeholder="e.g. Raja Shahid" 
                />
              </Field>
            </div>
            <Field label="Account / Phone / IBAN Number">
              <input 
                value={bank1Number} 
                onChange={e => { setBank1Number(e.target.value); updateCompiledPayment(bank1Name, e.target.value, bank1Title, bank2Name, bank2Number, bank2Title); }} 
                placeholder="e.g. 03461270679" 
              />
            </Field>
          </div>

          {/* Bank Account 2 */}
          <div style={{ background: "#FFFFFF", border: "1px solid #E2E8F0", borderRadius: 8, padding: "12px 14px", display: "flex", flexDirection: "column", gap: 10 }}>
            <div style={{ fontSize: "11px", fontWeight: 700, color: "#475569", textTransform: "uppercase", letterSpacing: "0.5px" }}>
              Secondary Account (Optional - Meezan / Alfalah / etc.)
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <Field label="Bank / Wallet Provider">
                <select value={bank2Name} onChange={e => { setBank2Name(e.target.value); updateCompiledPayment(bank1Name, bank1Number, bank1Title, e.target.value, bank2Number, bank2Title); }}>
                  {POPULAR_BANKS.map(b => (
                    <option key={b} value={b}>{b}</option>
                  ))}
                </select>
              </Field>
              <Field label="Account Title (optional)">
                <input 
                  value={bank2Title} 
                  onChange={e => { setBank2Title(e.target.value); updateCompiledPayment(bank1Name, bank1Number, bank1Title, bank2Name, bank2Number, e.target.value); }} 
                  placeholder="e.g. Capital Hardware" 
                />
              </Field>
            </div>
            <Field label="Account / Phone / IBAN Number">
              <input 
                value={bank2Number} 
                onChange={e => { setBank2Number(e.target.value); updateCompiledPayment(bank1Name, bank1Number, bank1Title, bank2Name, e.target.value, bank2Title); }} 
                placeholder="e.g. 0829-010-3838087" 
              />
            </Field>
          </div>

          {/* Formatted Statement Preview / Direct Edit */}
          <div style={{ borderTop: "1px dashed #CBD5E1", paddingTop: 10 }}>
            <Field label="Payment Line Preview (Customizable)">
              <input 
                value={form.paymentDetails || form.bankDetails || ""} 
                onChange={e => { set("paymentDetails", e.target.value); set("bankDetails", e.target.value); }} 
                placeholder="e.g. JAZZ CASH: 0307-8898663 RAJA SHAHID. MEEZAN BANK: 0829-010-3838087 CAPITAL HARDWARE" 
                style={{ background: "#FFFFFF", fontWeight: 600, color: "#0F172A" }}
              />
              <span className="hw-hint" style={{ marginTop: 4 }}>
                This is the exact line printed between telephone and email on your statements.
              </span>
            </Field>
          </div>
        </div>

        <Field label="Quotation Document Heading">
          <input value={form.quotationTitle || ""} onChange={e => set("quotationTitle", e.target.value)} placeholder="e.g. MATERIAL REQUEST" />
        </Field>
        <Field label="Invoice Document Heading">
          <input value={form.invoiceTitle || ""} onChange={e => set("invoiceTitle", e.target.value)} placeholder="e.g. TAX INVOICE" />
        </Field>

        <Field label="Currency symbol">
          <input value={form.currencySymbol || ""} onChange={e => set("currencySymbol", e.target.value)} maxLength={6} />
        </Field>
        <Field label="Default reorder level">
          <NumInput value={form.lowStockDefault || 5} onChange={n => set("lowStockDefault", Math.max(0, n))} />
        </Field>
      </div>
      <div className="hw-modal-actions" style={{ marginTop: 20 }}>
        <div style={{ flex: 1 }} />
        <button className="hw-btn-ghost" onClick={onClose}>Cancel</button>
        <button className="hw-btn-accent" onClick={handleSave}><Save size={14} /> Save settings</button>
      </div>
    </ModalShell>
  );
}

/* ---------------------------------------------------------
   SHARED UI ATOMS
--------------------------------------------------------- */
function ModalShell({ title, onClose, children }) {
  return (
    <div className="hw-modal-overlay" onClick={onClose}>
      <div className="hw-modal" onClick={e => e.stopPropagation()}>
        <div className="hw-modal-head">
          <h3>{title}</h3>
          <button className="hw-icon-btn" onClick={onClose}><X size={16} /></button>
        </div>
        {children}
      </div>
    </div>
  );
}
function Field({ label, children, span }) {
  return (
    <label className="hw-field" style={span ? { gridColumn: `span ${span}` } : undefined}>
      <span className="hw-field-label">{label}</span>
      {children}
    </label>
  );
}

/* Number input that lets the field go visually empty while typing (instead of
   snapping to "0" on backspace and then getting a stray leading zero) and
   only commits back to a real number on blur or once a valid number is typed. */
function NumInput({ value, onChange, ...rest }) {
  const [text, setText] = useState(value === undefined || value === null ? "" : String(value));
  const focused = useRef(false);

  useEffect(() => {
    if (!focused.current) setText(value === undefined || value === null ? "" : String(value));
  }, [value]);

  return (
    <input
      type="number"
      inputMode="decimal"
      value={text}
      onFocus={(e) => { focused.current = true; e.target.select(); }}
      onChange={(e) => {
        const v = e.target.value;
        setText(v);
        if (v === "" || v === "-" || v === "." || v === "-.") return;
        const num = Number(v);
        if (!Number.isNaN(num)) onChange(num);
      }}
      onBlur={() => {
        focused.current = false;
        if (text === "" || Number.isNaN(Number(text))) {
          setText("0");
          onChange(0);
        } else {
          setText(String(Number(text)));
        }
      }}
      {...rest}
    />
  );
}

/* ---------------------------------------------------------
   STYLE
--------------------------------------------------------- */
function Style() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');

      .hw-root {
        --bg: #F3F4F6;
        --surface: #FFFFFF;
        --surface-alt: #F9FAFB;
        --border: #E5E7EB;
        --ink: #111827;
        --ink-soft: #4B5563;
        --accent: #EA580C;
        --accent-dark: #C2410C;
        --accent-tint: #FFEDD5;
        --steel: #6B7280;
        --steel-tint: #F3F4F6;
        --success: #10B981;
        --success-tint: #D1FAE5;
        --danger: #EF4444;
        --danger-tint: #FEE2E2;
        --warn: #F59E0B;
        --warn-tint: #FEF3C7;
        --font-display: 'Inter', sans-serif;
        --font-body: 'Inter', sans-serif;
        --font-mono: 'IBM Plex Mono', monospace;

        display: flex;
        min-height: 640px;
        height: 100%;
        background: var(--bg);
        color: var(--ink);
        font-family: var(--font-body);
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid var(--border);
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.04);
      }
      .hw-root * { box-sizing: border-box; }
      @keyframes spin { to { transform: rotate(360deg); } }

      /* SIDEBAR */
      .hw-sidebar {
        width: 216px;
        flex-shrink: 0;
        background: #111827;
        color: #F3F4F6;
        display: flex;
        flex-direction: column;
        padding: 20px 14px;
        border-right: 1px solid #1F2937;
      }
      .hw-brand { display: flex; align-items: center; gap: 10px; padding: 4px 8px 22px; }
      .hw-brand-mark { font-size: 20px; width: 34px; height: 34px; border-radius: 8px; background: var(--accent); color: #FFFFFF; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
      .hw-brand-mark.small { width: 26px; height: 26px; font-size: 15px; border-radius: 6px; }
      .hw-brand-name { font-family: var(--font-display); font-weight: 600; font-size: 16px; line-height: 1.15; letter-spacing: 0.2px; color: #FFFFFF; }
      .hw-brand-sub { font-size: 11px; color: #9CA3AF; margin-top: 1px; }
      .hw-nav { display: flex; flex-direction: column; gap: 3px; flex: 1; }
      .hw-navbtn {
        display: flex; align-items: center; gap: 10px;
        padding: 9px 10px; border-radius: 7px; border: none; background: transparent; color: #9CA3AF;
        font-family: var(--font-body); font-size: 13.5px; font-weight: 500; cursor: pointer; text-align: left;
        transition: all 0.15s ease;
      }
      .hw-navbtn:hover { background: #1F2937; color: #FFFFFF; }
      .hw-navbtn.active { background: var(--accent); color: #FFFFFF; font-weight: 600; }
      .hw-navbtn span:first-of-type { flex: 1; }
      .hw-navbadge { background: var(--danger); color: #fff; font-size: 10.5px; font-weight: 700; border-radius: 20px; padding: 1px 6px; }
      .hw-navbtn.active .hw-navbadge { background: #FFFFFF; color: var(--accent); }
      .hw-navchevron { opacity: 0.6; }
      .hw-settings-btn {
        display: flex; align-items: center; gap: 8px; padding: 9px 10px; border-radius: 7px;
        background: transparent; border: 1px solid #374151; color: #9CA3AF; font-size: 13px; cursor: pointer; font-family: var(--font-body);
        transition: all 0.15s ease;
      }
      .hw-settings-btn:hover { background: #1F2937; color: #FFFFFF; }

      /* MAIN */
      .hw-main { flex: 1; overflow-y: auto; padding: 26px 30px 40px; }
      .hw-view-head { display: flex; align-items: flex-end; justify-content: space-between; margin-bottom: 20px; gap: 12px; flex-wrap: wrap; }
      .hw-eyebrow { font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.09em; color: var(--ink-soft); font-weight: 600; margin-bottom: 2px; }
      .hw-title { font-family: var(--font-display); font-size: 28px; font-weight: 700; letter-spacing: -0.5px; margin: 0; }

      /* STAT GRID */
      .hw-stat-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 22px; }
      .hw-stat { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 16px 18px; position: relative; box-shadow: 0 1px 3px rgba(0,0,0,0.02); }
      .hw-stat-icon { width: 28px; height: 28px; border-radius: 7px; display: flex; align-items: center; justify-content: center; background: var(--surface-alt); color: var(--ink-soft); margin-bottom: 10px; }
      .hw-stat-steel .hw-stat-icon { background: var(--steel-tint); color: var(--steel); }
      .hw-stat-warn .hw-stat-icon { background: var(--warn-tint); color: var(--warn); }
      .hw-stat-ok .hw-stat-icon { background: var(--success-tint); color: var(--success); }
      .hw-stat-accent .hw-stat-icon { background: var(--accent-tint); color: var(--accent-dark); }
      .hw-stat-value { font-family: var(--font-display); font-size: 24px; font-weight: 700; line-height: 1; letter-spacing: -0.3px; }
      .hw-stat-label { font-size: 12.5px; color: var(--ink-soft); margin-top: 5px; font-weight: 500; }
      .hw-stat-sub { font-size: 11px; color: var(--ink-soft); margin-top: 2px; }

      /* DASH COLUMNS */
      .hw-dash-cols { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
      .hw-card { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 18px 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.02); }
      .hw-card-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }
      .hw-card-head h3 { font-family: var(--font-display); font-size: 16px; font-weight: 700; margin: 0; color: var(--ink); }
      .hw-link { border: none; background: none; color: var(--accent); font-size: 12.5px; font-weight: 600; display: flex; align-items: center; gap: 2px; cursor: pointer; padding: 0; transition: color 0.15s ease; }
      .hw-link:hover { color: var(--accent-dark); }
      .hw-simple-list { list-style: none; margin: 0; padding: 0; }
      .hw-simple-list li { display: flex; align-items: center; gap: 8px; padding: 10px 0; border-bottom: 1px solid var(--border); font-size: 13.0px; }
      .hw-simple-list li:last-child { border-bottom: none; }
      .hw-hazard-dot { width: 7px; height: 7px; border-radius: 50%; background: var(--warn); flex-shrink: 0; }
      .hw-il-name { font-weight: 500; }
      .hw-il-meta { color: var(--ink-soft); font-size: 11.5px; margin-left: auto; }
      .hw-inv-num { font-family: var(--font-mono); font-size: 11.5px; font-weight: 600; color: var(--accent); background: var(--accent-tint); border: 1px solid rgba(234, 88, 12, 0.2); padding: 2px 8px; border-radius: 6px; }
      .hw-empty { color: var(--ink-soft); font-size: 13px; padding: 20px 2px; text-align: center; }

      /* SEARCH */
      .hw-search-row { display: flex; align-items: center; gap: 8px; background: var(--surface); border: 1px solid var(--border); border-radius: 8px; padding: 8px 12px; margin-bottom: 14px; color: var(--ink-soft); box-shadow: 0 1px 2px rgba(0,0,0,0.02); }
      .hw-search-row input { border: none; outline: none; flex: 1; font-size: 13.5px; font-family: var(--font-body); background: transparent; color: var(--ink); }

      /* BULK ACTIONS BAR */
      .hw-bulk-actions-bar {
        display: flex;
        align-items: center;
        gap: 12px;
        background: var(--accent-tint);
        border: 1px solid #FED7AA;
        border-radius: 8px;
        padding: 10px 16px;
        margin-bottom: 14px;
        color: var(--accent-dark);
        font-size: 13.5px;
        font-weight: 500;
        flex-wrap: wrap;
      }
      .hw-bulk-actions-bar span {
        margin-right: auto;
        font-weight: 700;
        color: var(--accent-dark);
      }

      /* DAY NAV */
      .hw-day-nav { display: flex; align-items: center; gap: 10px; margin-bottom: 16px; flex-wrap: wrap; }
      .hw-day-label { display: flex; align-items: center; gap: 6px; font-family: var(--font-display); font-size: 16px; font-weight: 700; background: var(--surface); border: 1px solid var(--border); border-radius: 8px; padding: 6px 14px; }
      .hw-date-input { border: 1px solid var(--border); border-radius: 8px; padding: 6px 10px; font-family: var(--font-body); font-size: 12.5px; background: var(--surface); color: var(--ink-soft); }

      /* TABLE */
      .hw-table-wrap { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; overflow: auto; box-shadow: 0 1px 3px rgba(0,0,0,0.02); }
      .hw-table { width: 100%; border-collapse: collapse; font-size: 13.5px; }
      .hw-table th { text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.06em; color: var(--ink-soft); font-weight: 600; padding: 12px 14px; border-bottom: 1px solid var(--border); background: var(--surface-alt); position: sticky; top: 0; z-index: 2; }
      .hw-table td { padding: 12px 14px; border-bottom: 1px solid var(--border); vertical-align: middle; }
      .hw-table tbody tr:last-child td { border-bottom: none; }
      .hw-table tbody tr:hover { background: var(--surface-alt); }
      .hw-row-low { background: var(--warn-tint) !important; }
      .hw-prod-name { font-weight: 500; }
      .hw-low-tag { display: inline-flex; align-items: center; gap: 3px; font-size: 10.5px; color: var(--warn); font-weight: 600; margin-top: 2px; }
      .hw-muted { color: var(--ink-soft); }
      .hw-mono { font-family: var(--font-mono); font-size: 12.5px; }
      .hw-qty-cell { display: inline-flex; align-items: center; gap: 6px; }
      .hw-qty-cell button { width: 20px; height: 20px; border-radius: 5px; border: 1px solid var(--border); background: var(--surface-alt); display: flex; align-items: center; justify-content: center; cursor: pointer; color: var(--ink); }
      .hw-qty-cell button:disabled { opacity: 0.4; cursor: not-allowed; }
      .hw-icon-btn { border: none; background: transparent; color: var(--ink-soft); cursor: pointer; padding: 5px; border-radius: 6px; display: inline-flex; }
      .hw-icon-btn:hover { background: var(--surface-alt); color: var(--ink); }
      .hw-mini-input { width: 56px; padding: 4px 6px; border: 1px solid var(--border); border-radius: 5px; font-family: var(--font-mono); font-size: 12.5px; background: var(--surface); }

      /* INVOICE CARD STYLE */
      .hw-invoice-card {
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 12px;
        overflow: hidden;
        display: flex;
        flex-direction: column;
        cursor: pointer;
        transition: all 0.2s ease;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
      }
      .hw-invoice-card:hover {
        border-color: var(--accent);
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05), 0 4px 6px -2px rgba(0, 0, 0, 0.02);
      }
      .hw-invoice-card-header {
        background: var(--surface-alt);
        border-bottom: 1px solid var(--border);
        padding: 8px 14px;
        display: flex;
        align-items: center;
        justify-content: space-between;
      }
      .hw-invoice-card-number {
        font-family: var(--font-mono);
        font-weight: 700;
        font-size: 11.5px;
        color: var(--accent);
      }
      .hw-invoice-card-body {
        padding: 16px 14px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex: 1;
      }
      .hw-invoice-card-details {
        display: flex;
        flex-direction: column;
        gap: 4px;
      }
      .hw-invoice-card-customer {
        font-weight: 700;
        font-size: 14.5px;
        color: var(--ink);
      }
      .hw-invoice-card-date {
        font-size: 11.5px;
        color: var(--ink-soft);
      }
      .hw-invoice-card-right {
        display: flex;
        flex-direction: column;
        align-items: flex-end;
        gap: 4px;
      }
      .hw-invoice-card-total {
        font-size: 19px;
        font-weight: 700;
        color: var(--ink);
        letter-spacing: -0.3px;
      }

      /* BUTTONS */
      .hw-btn-accent { display: inline-flex; align-items: center; gap: 6px; background: var(--accent); color: #FFFFFF; border: none; padding: 9px 15px; border-radius: 8px; font-weight: 600; font-size: 13px; cursor: pointer; font-family: var(--font-body); transition: background-color 0.15s ease; }
      .hw-btn-accent:hover { background: var(--accent-dark); }
      .hw-btn-accent:disabled { opacity: 0.45; cursor: not-allowed; }
      .hw-btn-ghost { background: transparent; border: 1px solid var(--border); color: var(--ink); padding: 9px 15px; border-radius: 8px; font-size: 13px; font-weight: 500; cursor: pointer; font-family: var(--font-body); transition: background-color 0.15s ease; }
      .hw-btn-ghost:hover { background: var(--surface-alt); }
      .hw-btn-danger { display: inline-flex; align-items: center; gap: 6px; background: var(--danger-tint); color: var(--danger); border: 1px solid #fecaca; padding: 9px 15px; border-radius: 8px; font-size: 13px; font-weight: 600; cursor: pointer; font-family: var(--font-body); }
      .hw-btn-block { width: 100%; justify-content: center; margin-top: 12px; }

      /* MODALS */
      .hw-modal-overlay { position: fixed; inset: 0; background: rgba(17, 24, 39, 0.6); backdrop-filter: blur(4px); display: flex; align-items: center; justify-content: center; z-index: 50; padding: 20px; overflow-y: auto; }
      .hw-modal { background: var(--surface); border-radius: 12px; width: 100%; max-width: 540px; padding: 22px 24px; max-height: 88vh; overflow-y: auto; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.1), 0 10px 10px -5px rgba(0,0,0,0.04); }
      .hw-modal-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; }
      .hw-modal-head h3 { font-family: var(--font-display); font-size: 18px; font-weight: 700; margin: 0; }
      .hw-form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
      .hw-field { display: flex; flex-direction: column; gap: 6px; font-size: 12.5px; }
      .hw-field-label { color: var(--ink-soft); font-weight: 600; }
      .hw-field input, .hw-field select { border: 1px solid #D1D5DB; border-radius: 7px; padding: 8px 10px; font-size: 13.5px; font-family: var(--font-body); background: #FFFFFF; color: var(--ink); outline: none; transition: all 0.15s ease; }
      .hw-field input:focus, .hw-field select:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(234, 88, 12, 0.15); }
      .hw-hint { font-size: 11px; color: var(--ink-soft); }
      .hw-modal-actions { display: flex; align-items: center; gap: 8px; margin-top: 18px; flex-wrap: wrap; }
      .hw-warn-banner { display: flex; align-items: center; gap: 7px; background: var(--warn-tint); color: var(--warn); font-size: 12.5px; font-weight: 600; padding: 8px 12px; border-radius: 8px; margin-top: 10px; }

      /* INVOICE LAYOUT */
      .hw-invoice-layout { display: grid; grid-template-columns: 1.6fr 1fr; gap: 14px; align-items: start; }
      .hw-invoice-layout .hw-table-wrap { max-height: 480px; overflow-y: auto; }
      .hw-summary-card { position: sticky; top: 20px; z-index: 5; }
      .hw-summary-card h3 { font-family: var(--font-display); font-size: 17px; font-weight: 600; margin: 0 0 12px; }
      .hw-summary-row { display: flex; align-items: center; justify-content: space-between; padding: 6px 0; font-size: 13.5px; }
      .hw-summary-total { border-top: 1px solid var(--border); margin-top: 6px; padding-top: 10px; font-family: var(--font-display); font-size: 19px; font-weight: 600; }
      .hw-profit-box { background: var(--steel-tint); border-radius: 8px; padding: 8px 12px 4px; margin-top: 12px; }
      .hw-profit-box .hw-summary-row { color: var(--steel); font-size: 12.5px; padding: 4px 0; }
      .hw-profit-label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.06em; color: var(--steel); font-weight: 700; opacity: 0.8; margin-bottom: 2px; }
      .hw-autocomplete { position: absolute; top: 42px; left: 0; right: 0; background: var(--surface); border: 1px solid var(--border); border-radius: 8px; box-shadow: 0 8px 20px rgba(0,0,0,0.08); z-index: 10; overflow: hidden; }
      .hw-autocomplete button { width: 100%; display: flex; align-items: center; justify-content: space-between; gap: 10px; padding: 9px 12px; border: none; background: none; text-align: left; cursor: pointer; font-size: 13px; border-bottom: 1px solid var(--border); }
      .hw-autocomplete button:last-child { border-bottom: none; }
      .hw-autocomplete button:hover { background: var(--surface-alt); }
      .hw-autocomplete button:disabled { opacity: 0.4; cursor: not-allowed; }
      .hw-table-tight td, .hw-table-tight th { padding: 8px 12px; }

      /* TOAST */
      .hw-toast { position: fixed; bottom: 22px; left: 50%; transform: translateX(-50%); background: var(--ink); color: #fff; padding: 9px 16px; border-radius: 8px; font-size: 13px; display: flex; align-items: center; gap: 7px; z-index: 100; box-shadow: 0 8px 20px rgba(0,0,0,0.25); }
      .hw-toast-err { background: var(--danger); }

      /* RECEIPT */
      .hw-receipt-modal { max-width: 400px; padding: 0; overflow: visible; background: transparent; }
      .hw-receipt { background: var(--surface); border-radius: 4px; padding: 22px 22px 14px; font-family: var(--font-body); position: relative; }
      .hw-receipt-head { display: flex; align-items: center; gap: 8px; margin-bottom: 10px; }
      .hw-receipt-shop { font-family: var(--font-display); font-weight: 600; font-size: 16px; flex: 1; }
      .hw-receipt-inv { font-family: var(--font-mono); font-size: 11px; background: var(--surface-alt); padding: 2px 7px; border-radius: 4px; }
      .hw-receipt-meta { display: flex; justify-content: space-between; font-size: 11.5px; color: var(--ink-soft); margin-bottom: 8px; }
      .hw-receipt-divider { border-top: 1px dashed var(--border); margin: 8px 0; }
      .hw-receipt-table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
      .hw-receipt-table th { text-align: left; font-size: 10px; text-transform: uppercase; color: var(--ink-soft); padding: 4px 4px; letter-spacing: 0.05em; }
      .hw-receipt-table td { padding: 5px 4px; }
      .hw-receipt-table td:not(:first-child) { text-align: right; }
      .hw-receipt-table th:not(:first-child) { text-align: right; }
      .hw-receipt-totals > div { display: flex; justify-content: space-between; font-size: 13px; padding: 3px 4px; }
      .hw-receipt-total-final { font-family: var(--font-display); font-size: 18px; font-weight: 600; border-top: 1px solid var(--border); margin-top: 4px; padding-top: 8px !important; }
      .hw-receipt-foot { text-align: center; font-size: 11.5px; color: var(--ink-soft); margin-top: 14px; font-style: italic; }
      .hw-receipt-teeth { height: 10px; margin: 10px -22px -14px; background: repeating-linear-gradient(90deg, transparent 0 6px, var(--bg) 6px 12px); }

      /* STATEMENT REPORT TEMPLATE */
      .hw-statement-modal { max-width: 850px; padding: 0; width: 100%; max-height: 92vh; display: flex; flex-direction: column; background: #FFFFFF; border-radius: 12px; overflow: hidden; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.35); }
      .hw-statement-card { background: #FFFFFF; overflow-y: auto; flex: 1; min-height: 0; font-family: var(--font-body); display: flex; flex-direction: column; width: 100%; position: relative; }
      .hw-statement-modal .hw-modal-actions { flex-shrink: 0; margin-top: 0; background: #F9FAFB; border-top: 1px solid #E5E7EB; padding: 12px 20px; z-index: 10; }
      
      /* AUTHENTIC EXCEL INVOICE HEADER (MATCHING USER'S PREFERRED EXCEL STYLE) */
      .hw-excel-header { background: #FFFFFF; padding: 22px 24px 14px; position: relative; border-bottom: 2px solid #1E293B; width: 100%; box-sizing: border-box; }
      .hw-excel-logo-badge { position: absolute; left: 24px; top: 20px; display: flex; align-items: center; gap: 8px; }
      .hw-excel-logo-label { display: flex; flex-direction: column; gap: 1px; }
      .hw-excel-main-center { display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center; width: 100%; margin: 0 auto; gap: 3px; }
      .hw-excel-org-title { font-size: 25px; font-weight: 900; color: #000000; letter-spacing: 0.5px; text-transform: uppercase; margin: 0 0 2px 0; font-family: var(--font-display); line-height: 1.15; text-align: center; }
      .hw-excel-org-urdu { font-size: 23px; font-weight: 700; color: #000000; direction: rtl; font-family: 'Jameel Noori Nastaleeq', 'Noto Nastaliq Urdu', 'Segoe UI', serif; margin: -2px 0 2px 0; line-height: 1.3; text-align: center; }
      .hw-excel-line-bold { font-size: 13px; font-weight: 800; color: #000000; margin: 1px 0; text-align: center; }
      .hw-excel-line-payment { font-size: 12px; font-weight: 800; color: #000000; text-transform: uppercase; letter-spacing: 0.3px; max-width: 680px; margin: 1px 0; text-align: center; }
      .hw-excel-line-email { font-size: 12.5px; font-weight: 700; color: #000000; margin: 1px 0; text-align: center; }
      .hw-excel-title-bar { text-align: center; margin-top: 14px; padding-top: 6px; width: 100%; display: flex; justify-content: center; }
      .hw-excel-doc-title { font-size: 19px; font-weight: 900; color: #000000; letter-spacing: 2px; text-transform: uppercase; margin: 0; font-family: var(--font-display); text-align: center; }

      /* Receiver / Billed-To Block */
      .hw-statement-receiver-block { display: flex; justify-content: space-between; align-items: flex-start; padding: 12px 20px; background: #FAFBFD; border-bottom: 1px solid #E5E7EB; flex-shrink: 0; gap: 16px; }
      .hw-statement-receiver-info { display: flex; flex-direction: column; gap: 2px; }
      .hw-statement-section-label { font-size: 10px; font-weight: 700; color: #6B7280; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 2px; }
      .hw-statement-customer-name { font-size: 16px; font-weight: 700; color: #111827; }
      .hw-statement-customer-sub { font-size: 12px; color: #4B5563; }
      .hw-statement-doc-info { display: flex; flex-direction: column; gap: 3px; text-align: right; }
      .hw-statement-doc-row { font-size: 12px; color: #4B5563; }
      .hw-statement-doc-label { color: #6B7280; margin-right: 6px; }
      .hw-statement-doc-val { font-weight: 600; color: #111827; font-family: var(--font-mono); }

      /* Dedicated Payment Details Bar */
      .hw-statement-payment-bar { margin: 10px 20px 0; background: #FFFBEB; border: 1px dashed #F59E0B; border-radius: 6px; padding: 8px 14px; display: flex; align-items: center; gap: 12px; font-size: 12px; flex-shrink: 0; }
      .hw-statement-payment-tag { font-weight: 800; color: #B45309; font-size: 10.5px; letter-spacing: 0.5px; text-transform: uppercase; flex-shrink: 0; }
      .hw-statement-payment-text { font-weight: 600; color: #92400E; word-break: break-word; }

      /* Summary Box */
      .hw-statement-summary-box { border: 1px solid #E5E7EB; border-radius: 8px; margin: 12px 20px 10px; display: grid; grid-template-columns: repeat(4, 1fr); background: #FAFBFD; overflow: hidden; flex-shrink: 0; }
      .hw-statement-summary-col { padding: 8px 12px; display: flex; flex-direction: column; gap: 2px; border-right: 1px solid #E5E7EB; }
      .hw-statement-summary-col:last-child { border-right: none; }
      .hw-statement-summary-label { font-size: 10px; font-weight: 600; color: #6B7280; text-transform: uppercase; letter-spacing: 0.5px; }
      .hw-statement-summary-value { font-size: 16px; font-weight: 700; color: #111827; font-family: var(--font-mono); }
      .hw-statement-summary-value.red { color: #DC2626; }
      .hw-statement-summary-value.green { color: #16A34A; }
      .hw-statement-summary-sub { font-size: 9.5px; color: #9CA3AF; }

      /* Entries Count Info */
      .hw-statement-entries-info { font-size: 12px; font-weight: 700; color: #111827; margin: 0 20px 6px; display: flex; justify-content: space-between; align-items: center; flex-shrink: 0; }

      /* Statement Table */
      .hw-statement-table-wrap { border: 1px solid #E5E7EB; border-radius: 8px; margin: 0 20px 16px; overflow-x: auto; overflow-y: visible; }
      .hw-statement-table { width: 100%; border-collapse: collapse; font-size: 12.5px; text-align: left; }
      .hw-statement-table th { background: #F3F4F6; padding: 8px 12px; font-weight: 600; color: #374151; font-size: 11.5px; border-bottom: 1px solid #E5E7EB; position: sticky; top: 0; z-index: 2; }
      .hw-statement-table td { padding: 7px 12px; border-bottom: 1px solid #F3F4F6; color: #4B5563; }
      .hw-statement-table tbody tr:last-child td { border-bottom: none; }
      .hw-statement-table td.mono { font-family: var(--font-mono); font-weight: 500; }
      .hw-statement-table td.green { color: #16A34A; font-weight: 600; }
      .hw-statement-table td.red { color: #DC2626; font-weight: 600; }

      /* Grand Total Row */
      .hw-statement-table-totals { background: #F9FAFB; border-top: 2px solid #E5E7EB; font-weight: 700 !important; color: #111827 !important; }
      .hw-statement-table-totals td { font-weight: 700; color: #111827; padding: 8px 12px; }

      /* Footer Banner */
      .hw-statement-footer { background: #EA580C; color: #FFFFFF; padding: 9px 20px; display: flex; align-items: center; justify-content: space-between; font-size: 11.5px; font-weight: 500; flex-shrink: 0; }
      .hw-statement-footer-left { display: flex; align-items: center; gap: 8px; }
      .hw-statement-footer-right { display: flex; align-items: center; gap: 4px; font-size: 11px; opacity: 0.95; }
      .hw-statement-btn-install { background: #FFFFFF; color: #EA580C; border: none; padding: 2px 6px; border-radius: 4px; font-size: 9.5px; font-weight: 700; text-transform: uppercase; }

      /* CUSTOM SCROLLBARS */
      .hw-main::-webkit-scrollbar,
      .hw-table-wrap::-webkit-scrollbar,
      .hw-statement-card::-webkit-scrollbar,
      .hw-modal::-webkit-scrollbar,
      .hw-modal-overlay::-webkit-scrollbar {
        width: 8px;
        height: 8px;
      }
      .hw-main::-webkit-scrollbar-track,
      .hw-table-wrap::-webkit-scrollbar-track,
      .hw-statement-card::-webkit-scrollbar-track,
      .hw-modal::-webkit-scrollbar-track,
      .hw-modal-overlay::-webkit-scrollbar-track {
        background: #F3F4F6;
        border-radius: 4px;
      }
      .hw-main::-webkit-scrollbar-thumb,
      .hw-table-wrap::-webkit-scrollbar-thumb,
      .hw-statement-card::-webkit-scrollbar-thumb,
      .hw-modal::-webkit-scrollbar-thumb,
      .hw-modal-overlay::-webkit-scrollbar-thumb {
        background: #CBD5E1;
        border-radius: 4px;
      }
      .hw-main::-webkit-scrollbar-thumb:hover,
      .hw-table-wrap::-webkit-scrollbar-thumb:hover,
      .hw-statement-card::-webkit-scrollbar-thumb:hover,
      .hw-modal::-webkit-scrollbar-thumb:hover,
      .hw-modal-overlay::-webkit-scrollbar-thumb:hover {
        background: #94A3B8;
      }

      @media print {
        .hw-statement-modal { max-height: none !important; overflow: visible !important; box-shadow: none !important; }
        .hw-statement-card { overflow: visible !important; height: auto !important; }
        .hw-statement-banner { background-color: #EA580C !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        .hw-statement-footer { background-color: #EA580C !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        .hw-statement-summary-box { background-color: #FAFBFD !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        .hw-statement-table th { background-color: #F3F4F6 !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        .hw-statement-table-totals { background-color: #F9FAFB !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        body * { visibility: hidden; }
        #hw-print-area, #hw-print-area * { visibility: visible; }
        #hw-print-area { position: fixed; top: 0; left: 0; width: 100%; }
        .hw-no-print { display: none !important; }
      }

      @media (max-width: 760px) {
        .hw-root { flex-direction: column; }
        .hw-sidebar { width: 100%; flex-direction: row; align-items: center; padding: 12px 14px; }
        .hw-nav { flex-direction: row; }
        .hw-navbtn span:first-of-type { display: none; }
        .hw-settings-btn span { display: none; }
        .hw-stat-grid { grid-template-columns: 1fr 1fr; }
        .hw-dash-cols, .hw-invoice-layout, .hw-form-grid { grid-template-columns: 1fr; }
      }
    `}</style>
  );
}

export function LoadingOverlay({ message = "Loading..." }) {
  return (
    <div style={{
      position: "fixed",
      top: 0, left: 0, right: 0, bottom: 0,
      background: "rgba(38, 36, 32, 0.7)",
      backdropFilter: "blur(4px)",
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
      zIndex: 9999,
      fontFamily: "var(--font-body)",
      color: "#FFFFFF"
    }}>
      <div style={{
        background: "#262420",
        border: "1px solid #D9720B",
        borderRadius: "16px",
        padding: "32px 48px",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        boxShadow: "0 10px 40px rgba(0, 0, 0, 0.4)",
        animation: "hwScaleIn 0.3s cubic-bezier(0.34, 1.56, 0.64, 1)"
      }}>
        <style>{`
          @keyframes hwScaleIn {
            from { transform: scale(0.9); opacity: 0; }
            to { transform: scale(1); opacity: 1; }
          }
          @keyframes hwSpinPulse {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .hw-loading-spinner {
            width: 52px;
            height: 52px;
            border: 4px solid rgba(217, 114, 11, 0.15);
            border-top-color: #D9720B;
            border-radius: 50%;
            animation: hwSpinPulse 0.9s cubic-bezier(0.55, 0.085, 0.68, 0.53) infinite;
            margin-bottom: 20px;
            box-shadow: 0 0 15px rgba(217, 114, 11, 0.2);
          }
        `}</style>
        <div className="hw-loading-spinner"></div>
        <div style={{ fontSize: "16px", fontWeight: "600", letterSpacing: "0.2px", color: "#F6F3EC" }}>{message}</div>
      </div>
    </div>
  );
}
