import { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";
// ─── SUPABASE ─────────────────────────────────────────────────────────────────
const SUPABASE_URL = "https://wvydsbjpgdadftqbkygr.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Ind2eWRzYmpwZ2RhZGZ0cWJreWdyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODEyNTYwMDIsImV4cCI6MjA5NjgzMjAwMn0.Z-VhIS4niHDpRT_SATW2u-n6botZXVCCjmIELQ30XJ8";
const supabase     = createClient(SUPABASE_URL, SUPABASE_KEY, {
  auth: {
    storage: window.localStorage,
    autoRefreshToken: true,
    persistSession: true,
    detectSessionInUrl: false,
  }
});
const NOTIFY_URL   = `${SUPABASE_URL}/functions/v1/notify-consume`;
const CLAUDE_URL   = `${SUPABASE_URL}/functions/v1/claude-proxy`;

// ─── CELLAR PHYSICAL STRUCTURE ───────────────────────────────────────────────
// Each row is a vertical rail from floor to ceiling.
// Shelf capacities per row type, listed TOP → BOTTOM (shelf index 0 = top)
const SHELF_STRUCTURE = {
  full:   Array(21).fill(3),                          // Rows 1,2,3,7,8,9 — 21 shelves × 3
  mid:    [...Array(5).fill(3), ...Array(2).fill(2), ...Array(2).fill(1)], // Rows 4,5,6
  magnum: Array(4).fill(1),                           // Magnum rows (2 rows × 4 shelves × 1)
};

const ROW_DEFS = [
  { id: "Row 1",    type: "full",   label: "R1" },
  { id: "Row 2",    type: "full",   label: "R2" },
  { id: "Row 3",    type: "full",   label: "R3" },
  { id: "Row 4",    type: "mid",    label: "R4" },
  { id: "Row 5",    type: "mid",    label: "R5" },
  { id: "Row 6",    type: "mid",    label: "R6" },
  { id: "Row 7",    type: "full",   label: "R7" },
  { id: "Row 8",    type: "full",   label: "R8" },
  { id: "Row 9",    type: "full",   label: "R9" },
  { id: "Magnum A", type: "magnum", label: "MA" },
  { id: "Magnum B", type: "magnum", label: "MB" },
];

const ROW_CAPACITY = Object.fromEntries(
  ROW_DEFS.map(r => [r.id, SHELF_STRUCTURE[r.type].reduce((a, b) => a + b, 0)])
);
const TOTAL_CAPACITY = Object.values(ROW_CAPACITY).reduce((a, b) => a + b, 0); // 449

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const STICKER = {
  red:    { label: "Needs to Age",  hex: "#c0392b", text: "#fff" },
  yellow: { label: "$50–$100",      hex: "#d4a017", text: "#1a1a1a" },
  blue:   { label: "$100+",         hex: "#2471a3", text: "#fff" },
  green:  { label: "Under $50",     hex: "#1e8449", text: "#fff" },
  black:  { label: "Past Peak",     hex: "#2c2c2c", text: "#aaa" },
};

const VARIETALS = ["Cabernet Sauvignon","Pinot Noir","Merlot","Syrah/Shiraz","Zinfandel","Malbec","Tempranillo","Sangiovese","Grenache","Nebbiolo","Chardonnay","Sauvignon Blanc","Pinot Grigio","Riesling","Viognier","Rosé","Champagne/Sparkling","Dessert/Port","Other"];
const REGIONS   = ["Napa Valley","Sonoma","Bordeaux","Burgundy","Rhône Valley","Tuscany","Piedmont","Rioja","Barossa Valley","Willamette Valley","Washington State","Argentina","Other"];
const OCCASIONS = ["Weeknight Dinner","Date Night","Celebration","Large Gathering","Wine Pairing","Casual Sipping","Gift"];

const VINTAGE_DATA = {
  "Napa Valley": {2012:98,2013:94,2014:96,2015:92,2016:97,2017:95,2018:99,2019:96,2020:88,2021:93,2022:95},
  "Bordeaux":    {2012:90,2013:88,2014:92,2015:97,2016:98,2017:93,2018:96,2019:95,2020:92,2021:90,2022:91},
  "Burgundy":    {2012:91,2013:89,2014:93,2015:95,2016:96,2017:90,2018:94,2019:98,2020:95,2021:92,2022:93},
  "Tuscany":     {2012:93,2013:90,2014:88,2015:92,2016:94,2017:89,2018:95,2019:93,2020:90,2021:91,2022:92},
  "Sonoma":      {2012:94,2013:92,2014:95,2015:91,2016:96,2017:93,2018:97,2019:94,2020:87,2021:92,2022:94},
};

// ─── FULL CELLAR — wine-inventoryV2.xlsx ──────────────────────────────────────
const SAMPLE_WINES = [
  { id:1000, winery:"Venge Vineyards", name:"Silenceaux", varietal:"Cabernet Sauvignon", vintage:2021, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:0, slot:0, sticker:"red", price:79.99, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1001, winery:"Venge Vineyards", name:"Silenceaux", varietal:"Cabernet Sauvignon", vintage:2021, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:0, slot:1, sticker:"red", price:79.99, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1002, winery:"Venge Vineyards", name:"Silenceaux", varietal:"Cabernet Sauvignon", vintage:2021, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:0, slot:2, sticker:"red", price:79.99, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1003, winery:"Beringer Vineyards", name:"Beringer Private Reserve Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2017, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:1, slot:0, sticker:"blue", price:170.0, drinkFrom:2022, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1004, winery:"Beringer Vineyards", name:"Beringer Private Reserve Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2017, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:1, slot:1, sticker:"blue", price:170.0, drinkFrom:2022, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1005, winery:"Beringer Vineyards", name:"Beringer Private Reserve Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2017, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:1, slot:2, sticker:"blue", price:170.0, drinkFrom:2022, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1006, winery:"DAOU Vineyards", name:"DAOU Estate Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2018, region:"Adelaida District, Paso Robles, California", country:"USA", row:"Row 1", shelf:2, slot:0, sticker:"blue", price:132.0, drinkFrom:2021, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1007, winery:"DAOU Vineyards", name:"DAOU Reserve Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 1", shelf:2, slot:1, sticker:"red", price:49.99, drinkFrom:2024, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1008, winery:"Beringer Vineyards", name:"Beringer Private Reserve Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2017, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:2, slot:2, sticker:"blue", price:170.0, drinkFrom:2022, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1009, winery:"Law Estate Wines", name:"Beyond Category", varietal:"Tempranillo", vintage:2022, region:"Paso Robles", country:"USA", row:"Row 1", shelf:3, slot:0, sticker:"red", price:80.0, drinkFrom:2025, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1010, winery:"Law Estate Wines", name:"Beyond Category", varietal:"Tempranillo", vintage:2022, region:"Paso Robles", country:"USA", row:"Row 1", shelf:3, slot:1, sticker:"red", price:80.0, drinkFrom:2025, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1011, winery:"Law Estate Wines", name:"Beyond Category", varietal:"Tempranillo", vintage:2022, region:"Paso Robles", country:"USA", row:"Row 1", shelf:3, slot:2, sticker:"red", price:80.0, drinkFrom:2025, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1012, winery:"Law Estate Wines", name:"Audacious", varietal:"Cabernet Sauvignon", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 1", shelf:4, slot:0, sticker:"red", price:90.0, drinkFrom:2025, drinkTo:2037, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1013, winery:"Law Estate Wines", name:"Audacious", varietal:"Cabernet Sauvignon", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 1", shelf:4, slot:1, sticker:"red", price:90.0, drinkFrom:2025, drinkTo:2037, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1014, winery:"Law Estate Wines", name:"Audacious", varietal:"Cabernet Sauvignon", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 1", shelf:4, slot:2, sticker:"red", price:90.0, drinkFrom:2025, drinkTo:2037, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1015, winery:"Turnbull Wine Cellars", name:"Bonne Vivante", varietal:"Cabernet Sauvignon", vintage:2021, region:"Oakville, Napa Valley, California", country:"USA", row:"Row 1", shelf:5, slot:0, sticker:"red", price:77.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1016, winery:"Conn Creek", name:"Class of '73", varietal:"Cabernet Sauvignon", vintage:2021, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:5, slot:1, sticker:"blue", price:140.0, drinkFrom:2025, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1017, winery:"Conn Creek", name:"Class of '73", varietal:"Cabernet Sauvignon", vintage:2021, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:5, slot:2, sticker:"blue", price:140.0, drinkFrom:2025, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1018, winery:"Conn Creek", name:"Class of '73", varietal:"Cabernet Sauvignon", vintage:2021, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:6, slot:0, sticker:"blue", price:140.0, drinkFrom:2025, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1019, winery:"Conn Creek", name:"Class of '73", varietal:"Cabernet Sauvignon", vintage:2021, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:6, slot:1, sticker:"blue", price:140.0, drinkFrom:2025, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1020, winery:"Conn Creek", name:"Class of '73", varietal:"Cabernet Sauvignon", vintage:2021, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:6, slot:2, sticker:"blue", price:140.0, drinkFrom:2025, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1021, winery:"Will Harlan", name:"The Mascot", varietal:"Cabernet Sauvignon", vintage:2014, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:7, slot:0, sticker:"blue", price:149.0, drinkFrom:2018, drinkTo:2029, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1022, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 1", shelf:7, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1023, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 1", shelf:7, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1024, winery:"Spring Mountain Vineyard", name:"Spring Mountain Vinyard Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2018, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:8, slot:0, sticker:"blue", price:100.0, drinkFrom:2022, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1025, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 1", shelf:8, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1026, winery:"Spring Mountain Vineyard", name:"Spring Mountain Vinyard Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2018, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:8, slot:2, sticker:"blue", price:100.0, drinkFrom:2022, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1027, winery:"Caldwell Vineyard", name:"Society of Smugglers", varietal:"Cabernet Sauvignon", vintage:2019, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:9, slot:0, sticker:"blue", price:200.0, drinkFrom:2025, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1028, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 1", shelf:9, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1029, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 1", shelf:9, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1030, winery:"Alexander Valley, Sonoma County", name:"The Setting", varietal:"Cabernet Sauvignon", vintage:2022, region:"Sonoma County", country:"USA", row:"Row 1", shelf:10, slot:0, sticker:"yellow", price:90.0, drinkFrom:2026, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1031, winery:"Alexander Valley, Sonoma County", name:"The Setting", varietal:"Cabernet Sauvignon", vintage:2022, region:"Sonoma County", country:"USA", row:"Row 1", shelf:10, slot:1, sticker:"yellow", price:90.0, drinkFrom:2026, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1032, winery:"Alexander Valley, Sonoma County", name:"The Setting", varietal:"Cabernet Sauvignon", vintage:2022, region:"Sonoma County", country:"USA", row:"Row 1", shelf:10, slot:2, sticker:"yellow", price:90.0, drinkFrom:2026, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1033, winery:"Alexander Valley, Sonoma County", name:"The Setting", varietal:"Cabernet Sauvignon", vintage:2022, region:"Sonoma County", country:"USA", row:"Row 1", shelf:11, slot:0, sticker:"yellow", price:90.0, drinkFrom:2026, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1034, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 1", shelf:11, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1035, winery:"Alexander Valley, Sonoma County", name:"The Setting", varietal:"Cabernet Sauvignon", vintage:2022, region:"Sonoma County", country:"USA", row:"Row 1", shelf:11, slot:2, sticker:"yellow", price:90.0, drinkFrom:2026, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1036, winery:"Castello di Amorosa", name:"Napa Valley Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2019, region:"Napa Valley", country:"USA", row:"Row 1", shelf:12, slot:0, sticker:"yellow", price:69.0, drinkFrom:20232031, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1037, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 1", shelf:12, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1038, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 1", shelf:12, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1039, winery:"DeLille Cellars", name:"Four Flags Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2012, region:"Red Mountain & Columbia Valley, Washington", country:"USA", row:"Row 1", shelf:13, slot:0, sticker:"yellow", price:70.0, drinkFrom:2025, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1040, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 1", shelf:13, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1041, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 1", shelf:13, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1042, winery:"Leeuwin Estate", name:"Art Series Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2011, region:"Margeret River, Western Australia", country:"Australia", row:"Row 1", shelf:14, slot:0, sticker:"yellow", price:79.0, drinkFrom:2025, drinkTo:2031, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1043, winery:"Leeuwin Estate", name:"Art Series Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2011, region:"Margeret River, Western Australia", country:"Australia", row:"Row 1", shelf:14, slot:1, sticker:"yellow", price:79.0, drinkFrom:2025, drinkTo:2031, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1044, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 1", shelf:14, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1045, winery:"Frank Family, Napa Valley", name:"Frank Famliy Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2023, region:"Napa Valley", country:"USA", row:"Row 1", shelf:15, slot:0, sticker:"yellow", price:60.0, drinkFrom:2026, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1046, winery:"Frank Family, Napa Valley", name:"Frank Famliy Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2023, region:"Napa Valley", country:"USA", row:"Row 1", shelf:15, slot:1, sticker:"yellow", price:60.0, drinkFrom:2026, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1047, winery:"Frank Family, Napa Valley", name:"Frank Famliy Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2023, region:"Napa Valley", country:"USA", row:"Row 1", shelf:15, slot:2, sticker:"yellow", price:60.0, drinkFrom:2026, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1048, winery:"Frank Family, Napa Valley", name:"Frank Famliy Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2023, region:"Napa Valley", country:"USA", row:"Row 1", shelf:16, slot:0, sticker:"yellow", price:60.0, drinkFrom:2026, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1049, winery:"Frank Family, Napa Valley", name:"Frank Famliy Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2023, region:"Napa Valley", country:"USA", row:"Row 1", shelf:16, slot:1, sticker:"yellow", price:60.0, drinkFrom:2026, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1050, winery:"Frank Family, Napa Valley", name:"Frank Famliy Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2023, region:"Napa Valley", country:"USA", row:"Row 1", shelf:16, slot:2, sticker:"yellow", price:60.0, drinkFrom:2026, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1051, winery:"Chappellet Vineyard", name:"Chappellet Mountain Cuvée Proprietary Blend", varietal:"Cabernet Sauvignon", vintage:2023, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:17, slot:0, sticker:"green", price:40.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1052, winery:"Duckhorn Portfolio", name:"Decoy Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2021, region:"California", country:"USA", row:"Row 1", shelf:17, slot:1, sticker:"green", price:25.0, drinkFrom:2023, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1053, winery:"Ciel du Cheval Vineyard", name:"Cadence", varietal:"Cabernet Sauvignon", vintage:2018, region:"Red Mountain, Washington", country:"USA", row:"Row 1", shelf:17, slot:2, sticker:"green", price:45.0, drinkFrom:2026, drinkTo:2038, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1054, winery:"Chappellet Vineyard", name:"Chappellet Mountain Cuvée Proprietary Blend", varietal:"Cabernet Sauvignon", vintage:2024, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:18, slot:0, sticker:"green", price:49.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1055, winery:"Chappellet Vineyard", name:"Chappellet Mountain Cuvée Proprietary Blend", varietal:"Cabernet Sauvignon", vintage:2024, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:18, slot:1, sticker:"green", price:49.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1056, winery:"Chappellet Vineyard", name:"Chappellet Mountain Cuvée Proprietary Blend", varietal:"Cabernet Sauvignon", vintage:2024, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:18, slot:2, sticker:"green", price:49.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1057, winery:"Hill Family Estate", name:"Hill Family", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:19, slot:0, sticker:"green", price:45.0, drinkFrom:2025, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1058, winery:"Hill Family Estate", name:"Hill Family", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:19, slot:1, sticker:"green", price:45.0, drinkFrom:2025, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1059, winery:"Hill Family Estate", name:"Hill Family", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:19, slot:2, sticker:"green", price:45.0, drinkFrom:2025, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1060, winery:"Hill Family Estate", name:"Hill Family", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:20, slot:0, sticker:"green", price:45.0, drinkFrom:2025, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1061, winery:"Hill Family Estate", name:"Hill Family", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:20, slot:1, sticker:"green", price:45.0, drinkFrom:2025, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1062, winery:"Hill Family Estate", name:"Hill Family", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley, California", country:"USA", row:"Row 1", shelf:20, slot:2, sticker:"green", price:45.0, drinkFrom:2025, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1063, winery:"Post and Beam", name:"Post and Beam Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", country:"USA", row:"Row 2", shelf:0, slot:0, sticker:"red", price:50.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1064, winery:"Post and Beam", name:"Post and Beam Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", country:"USA", row:"Row 2", shelf:0, slot:1, sticker:"red", price:50.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1065, winery:"Post and Beam", name:"Post and Beam Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", country:"USA", row:"Row 2", shelf:0, slot:2, sticker:"red", price:50.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1066, winery:"Mai Soli Wines", name:"Mai Soli Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2020, region:"Alexander Valley, Sonoma County, California", country:"USA", row:"Row 2", shelf:1, slot:0, sticker:"red", price:65.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1067, winery:"Mai Soli Wines", name:"Mai Soli Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2020, region:"Alexander Valley, Sonoma County, California", country:"USA", row:"Row 2", shelf:1, slot:1, sticker:"red", price:65.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1068, winery:"Mai Soli Wines", name:"Mai Soli Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2020, region:"Alexander Valley, Sonoma County, California", country:"USA", row:"Row 2", shelf:1, slot:2, sticker:"red", price:65.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1069, winery:"Educated Guess", name:"Educated Guess Reserve Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2021, region:"Oakville, Napa Valley", country:"USA", row:"Row 2", shelf:2, slot:0, sticker:"red", price:60.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1070, winery:"Bella Union", name:"Bella Union Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", country:"USA", row:"Row 2", shelf:2, slot:1, sticker:"red", price:80.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1071, winery:"Mai Soli Wines", name:"Mai Soli Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2020, region:"Alexander Valley, Sonoma County, California", country:"USA", row:"Row 2", shelf:2, slot:2, sticker:"red", price:65.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1072, winery:"Bella Union", name:"Bella Union Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", country:"USA", row:"Row 2", shelf:3, slot:0, sticker:"red", price:80.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1073, winery:"Bella Union", name:"Bella Union Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", country:"USA", row:"Row 2", shelf:3, slot:1, sticker:"red", price:80.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1074, winery:"Bella Union", name:"Bella Union Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", country:"USA", row:"Row 2", shelf:3, slot:2, sticker:"red", price:80.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1075, winery:"Bella Union", name:"Bella Union Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", country:"USA", row:"Row 2", shelf:4, slot:0, sticker:"red", price:80.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1076, winery:"Bella Union", name:"Bella Union Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", country:"USA", row:"Row 2", shelf:4, slot:1, sticker:"red", price:80.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1077, winery:"Bella Union", name:"Bella Union Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", country:"USA", row:"Row 2", shelf:4, slot:2, sticker:"red", price:80.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1078, winery:"Bella Union", name:"Bella Union Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", country:"USA", row:"Row 2", shelf:5, slot:0, sticker:"red", price:80.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1079, winery:"Bella Union", name:"Bella Union Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", country:"USA", row:"Row 2", shelf:5, slot:1, sticker:"red", price:80.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1080, winery:"Bella Union", name:"Bella Union Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", country:"USA", row:"Row 2", shelf:5, slot:2, sticker:"red", price:80.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1081, winery:"Odette Estate Winery", name:"Adaptation Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2019, region:"Napa Valley, California", country:"USA", row:"Row 2", shelf:6, slot:0, sticker:"red", price:60.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1082, winery:"Odette Estate Winery", name:"Adaptation Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2019, region:"Napa Valley, California", country:"USA", row:"Row 2", shelf:6, slot:1, sticker:"red", price:60.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1083, winery:"Odette Estate Winery", name:"Adaptation Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2019, region:"Napa Valley, California", country:"USA", row:"Row 2", shelf:6, slot:2, sticker:"red", price:60.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1084, winery:"Law Estate Wines", name:"Sagacious", varietal:"Grenache", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:7, slot:0, sticker:"red", price:90.0, drinkFrom:2025, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1085, winery:"Law Estate Wines", name:"Sagacious", varietal:"Grenache", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:7, slot:1, sticker:"red", price:90.0, drinkFrom:2025, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1086, winery:"Law Estate Wines", name:"Sagacious", varietal:"Grenache", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:7, slot:2, sticker:"red", price:90.0, drinkFrom:2025, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1087, winery:"Law Estate Wines", name:"Beguiling ", varietal:"Grenache", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:8, slot:0, sticker:"red", price:90.0, drinkFrom:2025, drinkTo:2037, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1088, winery:"Law Estate Wines", name:"Beguiling ", varietal:"Grenache", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:8, slot:1, sticker:"red", price:90.0, drinkFrom:2025, drinkTo:2037, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1089, winery:"Law Estate Wines", name:"Beguiling ", varietal:"Grenache", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:8, slot:2, sticker:"red", price:90.0, drinkFrom:2025, drinkTo:2037, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1090, winery:"Law Estate Wines", name:"First Tracks", varietal:"Other", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:9, slot:0, sticker:"yellow", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1091, winery:"Law Estate Wines", name:"First Tracks", varietal:"Other", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:9, slot:1, sticker:"yellow", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1092, winery:"Law Estate Wines", name:"First Tracks", varietal:"Other", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:9, slot:2, sticker:"yellow", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1093, winery:"Law Estate Wines", name:"First Tracks", varietal:"Other", vintage:2019, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:10, slot:0, sticker:"yellow", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1094, winery:"Law Estate Wines", name:"First Tracks", varietal:"Other", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:10, slot:1, sticker:"yellow", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1095, winery:"Law Estate Wines", name:"First Tracks", varietal:"Other", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:10, slot:2, sticker:"yellow", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1096, winery:"Dunham Cellars", name:"Dunham Cellars", varietal:"Cabernet Sauvignon", vintage:2007, region:"Washington", country:"USA", row:"Row 2", shelf:11, slot:0, sticker:"yellow", price:70.0, drinkFrom:2026, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1097, winery:"Dunham Cellars", name:"Dunham Cellars", varietal:"Cabernet Sauvignon", vintage:2007, region:"Washington", country:"USA", row:"Row 2", shelf:11, slot:1, sticker:"yellow", price:70.0, drinkFrom:2026, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1098, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 2", shelf:11, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1099, winery:"Penfolds", name:"Bin 389 Cabernet Shiraz", varietal:"Cabernet Sauvignon", vintage:2022, region:"South Australia", country:"Australia", row:"Row 2", shelf:12, slot:0, sticker:"yellow", price:75.0, drinkFrom:2025, drinkTo:2050, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1100, winery:"Penfolds", name:"Bin 389 Cabernet Shiraz", varietal:"Cabernet Sauvignon", vintage:2022, region:"South Australia", country:"Australia", row:"Row 2", shelf:12, slot:1, sticker:"yellow", price:75.0, drinkFrom:2025, drinkTo:2050, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1101, winery:"Penfolds", name:"Bin 389 Cabernet Shiraz", varietal:"Cabernet Sauvignon", vintage:2022, region:"South Australia", country:"Australia", row:"Row 2", shelf:12, slot:2, sticker:"yellow", price:75.0, drinkFrom:2025, drinkTo:2050, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1102, winery:"Penfolds", name:"Bin 600 Cabernet Shiraz", varietal:"Cabernet Sauvignon", vintage:2018, region:"California", country:"USA", row:"Row 2", shelf:13, slot:0, sticker:"yellow", price:50.0, drinkFrom:2022, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1103, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 2", shelf:13, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1104, winery:"Penfolds", name:"Bin 600 Cabernet Shiraz", varietal:"Cabernet Sauvignon", vintage:2018, region:"California", country:"USA", row:"Row 2", shelf:13, slot:2, sticker:"yellow", price:50.0, drinkFrom:2022, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1105, winery:"Penfolds", name:"Bin 704 Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2019, region:"Napa Valley, California", country:"USA", row:"Row 2", shelf:14, slot:0, sticker:"yellow", price:65.0, drinkFrom:2023, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1106, winery:"Penfolds", name:"Bin 704 Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2019, region:"Napa Valley, California", country:"USA", row:"Row 2", shelf:14, slot:1, sticker:"yellow", price:65.0, drinkFrom:2023, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1107, winery:"Penfolds", name:"Bin 389 Cabernet Shiraz", varietal:"Cabernet Sauvignon", vintage:2019, region:"South Australia", country:"Australia", row:"Row 2", shelf:14, slot:2, sticker:"yellow", price:70.0, drinkFrom:2023, drinkTo:2038, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1108, winery:"Law Estate Wines", name:"First Tracks", varietal:"Other", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:15, slot:0, sticker:"yellow", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1109, winery:"Law Estate Wines", name:"First Tracks", varietal:"Other", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:15, slot:1, sticker:"yellow", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1110, winery:"Law Estate Wines", name:"First Tracks", varietal:"Other", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:15, slot:2, sticker:"yellow", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1111, winery:"Law Estate Wines", name:"First Tracks", varietal:"Other", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:16, slot:0, sticker:"yellow", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1112, winery:"Law Estate Wines", name:"First Tracks", varietal:"Other", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:16, slot:1, sticker:"yellow", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1113, winery:"Law Estate Wines", name:"First Tracks", varietal:"Other", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 2", shelf:16, slot:2, sticker:"yellow", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1114, winery:"Silver Oak Cellars", name:"Silver Oak Cabernet Sauvignon Alexander Valley", varietal:"Cabernet Sauvignon", vintage:2019, region:"Alexander Valley, Sonoma County, California", country:"USA", row:"Row 2", shelf:17, slot:0, sticker:"yellow", price:85.0, drinkFrom:2023, drinkTo:2038, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1115, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 2", shelf:17, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1116, winery:"Elderton Wines", name:"Elderton Ode to Lorraine", varietal:"Cabernet Sauvignon", vintage:2018, region:"Barrosa Valley", country:"Australia", row:"Row 2", shelf:17, slot:2, sticker:"yellow", price:55.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1117, winery:"Greenwing (Duckhorn)", name:"Greenwing Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Columbia Valley, Washington", country:"USA", row:"Row 2", shelf:18, slot:0, sticker:"green", price:25.0, drinkFrom:2024, drinkTo:2029, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1118, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 2", shelf:18, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1119, winery:"Marietta Cellars", name:"Game Trail", varietal:"Cabernet Sauvignon", vintage:2021, region:"Yorkville Highlands, California", country:"USA", row:"Row 2", shelf:18, slot:2, sticker:"green", price:50.0, drinkFrom:2025, drinkTo:2039, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1120, winery:"Marietta Cellars", name:"Game Trail", varietal:"Cabernet Sauvignon", vintage:2021, region:"Yorkville Highlands, California", country:"USA", row:"Row 2", shelf:19, slot:0, sticker:"green", price:50.0, drinkFrom:2025, drinkTo:2039, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1121, winery:"Marietta Cellars", name:"Game Trail", varietal:"Cabernet Sauvignon", vintage:2021, region:"Yorkville Highlands, California", country:"USA", row:"Row 2", shelf:19, slot:1, sticker:"green", price:50.0, drinkFrom:2025, drinkTo:2039, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1122, winery:"Marietta Cellars", name:"Game Trail", varietal:"Cabernet Sauvignon", vintage:2021, region:"Yorkville Highlands, California", country:"USA", row:"Row 2", shelf:19, slot:2, sticker:"green", price:50.0, drinkFrom:2025, drinkTo:2039, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1123, winery:"Cornwell Vineyards", name:"Courtship", varietal:"Cabernet Sauvignon", vintage:2021, region:"Sonoma County, California", country:"USA", row:"Row 2", shelf:20, slot:0, sticker:"green", price:50.0, drinkFrom:2026, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1124, winery:"Cornwell Vineyards", name:"Courtship", varietal:"Cabernet Sauvignon", vintage:2021, region:"Sonoma County, California", country:"USA", row:"Row 2", shelf:20, slot:1, sticker:"green", price:50.0, drinkFrom:2026, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1125, winery:"Cornwell Vineyards", name:"Courtship", varietal:"Cabernet Sauvignon", vintage:2021, region:"Sonoma County, California", country:"USA", row:"Row 2", shelf:20, slot:2, sticker:"green", price:50.0, drinkFrom:2026, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1126, winery:"Cass Winery", name:"Estate CAB Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2017, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:0, slot:0, sticker:"blue", price:175.0, drinkFrom:2022, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1127, winery:"Caymus Vineyards", name:"50th Anniversary Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", country:"USA", row:"Row 3", shelf:0, slot:1, sticker:"blue", price:105.0, drinkFrom:2025, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1128, winery:"Caymus Vineyards", name:"50th Anniversary Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", country:"USA", row:"Row 3", shelf:0, slot:2, sticker:"blue", price:105.0, drinkFrom:2025, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1129, winery:"Turley Wine Cellars", name:"Cedarman Zinfandel", varietal:"Zinfandel", vintage:2021, region:"Napa Valley", country:"USA", row:"Row 3", shelf:1, slot:0, sticker:"red", price:55.0, drinkFrom:2025, drinkTo:2041, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1130, winery:"Justin Vineyards", name:"Isosceles", varietal:"Cabernet Sauvignon", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:1, slot:1, sticker:"red", price:85.0, drinkFrom:2025, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1131, winery:"Justin Vineyards", name:"Savant", varietal:"Cabernet Sauvignon", vintage:2019, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:1, slot:2, sticker:"yellow", price:55.0, drinkFrom:2021, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1132, winery:"Produttori del Barbaresco", name:"Barbaresco", varietal:"Nebbiolo", vintage:2020, region:"Barbaresco DOCG", country:"Italy", row:"Row 3", shelf:2, slot:0, sticker:"red", price:46.0, drinkFrom:2023, drinkTo:2045, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1133, winery:"Justin Vineyards", name:"Isosceles", varietal:"Cabernet Sauvignon", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:2, slot:1, sticker:"red", price:85.0, drinkFrom:2025, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1134, winery:"Justin Vineyards", name:"Isosceles", varietal:"Cabernet Sauvignon", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:2, slot:2, sticker:"red", price:85.0, drinkFrom:2025, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1135, winery:"Law Estate Wines", name:"Beguiling ", varietal:"Syrah/Shiraz", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:3, slot:0, sticker:"red", price:85.0, drinkFrom:2025, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1136, winery:"Law Estate Wines", name:"Beguiling ", varietal:"Syrah/Shiraz", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:3, slot:1, sticker:"red", price:85.0, drinkFrom:2025, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1137, winery:"Law Estate Wines", name:"Beguiling ", varietal:"Syrah/Shiraz", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:3, slot:2, sticker:"red", price:85.0, drinkFrom:2025, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1138, winery:"Penfolds", name:"Bin 389 Cabernet Shiraz", varietal:"Cabernet Sauvignon", vintage:2021, region:"South Australia", country:"Australia", row:"Row 3", shelf:4, slot:0, sticker:"red", price:70.0, drinkFrom:2025, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1139, winery:"Penfolds", name:"Bin 389 Cabernet Shiraz", varietal:"Cabernet Sauvignon", vintage:2021, region:"South Australia", country:"Australia", row:"Row 3", shelf:4, slot:1, sticker:"red", price:70.0, drinkFrom:2025, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1140, winery:"Penfolds", name:"Bin 389 Cabernet Shiraz", varietal:"Cabernet Sauvignon", vintage:2019, region:"South Australia", country:"Australia", row:"Row 3", shelf:4, slot:2, sticker:"red", price:70.0, drinkFrom:2025, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1141, winery:"Conn Creek Winary", name:"Conn Creek Cabernet Sauvignon Class of '73", varietal:"Cabernet Sauvignon", vintage:2021, region:"Napa Valley", country:"USA", row:"Row 3", shelf:5, slot:0, sticker:"red", price:150.0, drinkFrom:2028, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1142, winery:"Conn Creek Winary", name:"Conn Creek Cabernet Sauvignon Class of '73", varietal:"Cabernet Sauvignon", vintage:2021, region:"Napa Valley", country:"USA", row:"Row 3", shelf:5, slot:1, sticker:"red", price:150.0, drinkFrom:2028, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1143, winery:"Conn Creek Winary", name:"Conn Creek Cabernet Sauvignon Class of '73", varietal:"Cabernet Sauvignon", vintage:2021, region:"Napa Valley", country:"USA", row:"Row 3", shelf:5, slot:2, sticker:"red", price:150.0, drinkFrom:2028, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1144, winery:"Law Estate Wines", name:"Intrepid by Law", varietal:"Syrah/Shiraz", vintage:2022, region:"", country:"", row:"Row 3", shelf:6, slot:0, sticker:"red", price:115.0, drinkFrom:2027, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1145, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 3", shelf:6, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1146, winery:"Conn Creek Winary", name:"Conn Creek Cabernet Sauvignon Class of '73", varietal:"Cabernet Sauvignon", vintage:2021, region:"Napa Valley", country:"USA", row:"Row 3", shelf:6, slot:2, sticker:"red", price:150.0, drinkFrom:2028, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1147, winery:"Law Estate Wines", name:"Prima Law", varietal:"Syrah/Shiraz", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:7, slot:0, sticker:"blue", price:115.0, drinkFrom:2025, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1148, winery:"Law Estate Wines", name:"Intrepid by Law", varietal:"Syrah/Shiraz", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:7, slot:1, sticker:"red", price:115.0, drinkFrom:2027, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1149, winery:"Law Estate Wines", name:"The Nines", varietal:"Grenache", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:7, slot:2, sticker:"red", price:115.0, drinkFrom:2027, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1150, winery:"Law Estate Wines", name:"Prima Law", varietal:"Syrah/Shiraz", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:8, slot:0, sticker:"blue", price:115.0, drinkFrom:2025, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1151, winery:"Law Estate Wines", name:"Prima Law", varietal:"Syrah/Shiraz", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:8, slot:1, sticker:"blue", price:115.0, drinkFrom:2025, drinkTo:2036, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1152, winery:"Law Estate Wines", name:"Intrepid by Law", varietal:"Syrah/Shiraz", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:8, slot:2, sticker:"red", price:115.0, drinkFrom:2027, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1153, winery:"Law Estate Wines", name:"Intrepid by Law", varietal:"Syrah/Shiraz", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:9, slot:0, sticker:"red", price:115.0, drinkFrom:2027, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1154, winery:"Law Estate Wines", name:"Intrepid by Law", varietal:"Syrah/Shiraz", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:9, slot:1, sticker:"red", price:115.0, drinkFrom:2027, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1155, winery:"Law Estate Wines", name:"Intrepid by Law", varietal:"Syrah/Shiraz", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:9, slot:2, sticker:"red", price:115.0, drinkFrom:2027, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1156, winery:"Chateau Barreyre", name:"Chateau Barreyre Bordeaux Superieur", varietal:"Other", vintage:2016, region:"Bordeaux", country:"France", row:"Row 3", shelf:10, slot:0, sticker:"green", price:25.0, drinkFrom:2022, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1157, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 3", shelf:10, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1158, winery:"Law Estate Wines", name:"Intrepid by Law", varietal:"Syrah/Shiraz", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:10, slot:2, sticker:"red", price:115.0, drinkFrom:2027, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1159, winery:"Justin Vineyards", name:"Isosceles", varietal:"Cabernet Sauvignon", vintage:2018, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:11, slot:0, sticker:"yellow", price:85.0, drinkFrom:2020, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1160, winery:"Justin Vineyards", name:"Isosceles", varietal:"Cabernet Sauvignon", vintage:2018, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:11, slot:1, sticker:"yellow", price:85.0, drinkFrom:2020, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1161, winery:"Justin Vineyards", name:"Savant", varietal:"Cabernet Sauvignon", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:11, slot:2, sticker:"red", price:55.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1162, winery:"The Prisoner Wine Co.", name:"The Prisoner Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley, California", country:"USA", row:"Row 3", shelf:12, slot:0, sticker:"yellow", price:52.0, drinkFrom:2023, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1163, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 3", shelf:12, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1164, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 3", shelf:12, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1165, winery:"Yount Ridge Cellars", name:"Yount Ridge", varietal:"Cabernet Sauvignon", vintage:2023, region:"Napa Valley, California", country:"USA", row:"Row 3", shelf:13, slot:0, sticker:"yellow", price:65.0, drinkFrom:2024, drinkTo:2047, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1166, winery:"Yount Ridge Cellars", name:"Yount Ridge", varietal:"Cabernet Sauvignon", vintage:2023, region:"Napa Valley, California", country:"USA", row:"Row 3", shelf:13, slot:1, sticker:"yellow", price:65.0, drinkFrom:2024, drinkTo:2047, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1167, winery:"Yount Ridge Cellars", name:"Yount Ridge", varietal:"Cabernet Sauvignon", vintage:2023, region:"Napa Valley, California", country:"USA", row:"Row 3", shelf:13, slot:2, sticker:"yellow", price:65.0, drinkFrom:2024, drinkTo:2047, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1168, winery:"Yount Ridge Cellars", name:"Yount Ridge", varietal:"Cabernet Sauvignon", vintage:2023, region:"Napa Valley, California", country:"USA", row:"Row 3", shelf:14, slot:0, sticker:"yellow", price:65.0, drinkFrom:2024, drinkTo:2047, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1169, winery:"Yount Ridge Cellars", name:"Yount Ridge", varietal:"Cabernet Sauvignon", vintage:2023, region:"Napa Valley, California", country:"USA", row:"Row 3", shelf:14, slot:1, sticker:"yellow", price:65.0, drinkFrom:2024, drinkTo:2047, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1170, winery:"Yount Ridge Cellars", name:"Yount Ridge", varietal:"Cabernet Sauvignon", vintage:2023, region:"Napa Valley, California", country:"USA", row:"Row 3", shelf:14, slot:2, sticker:"yellow", price:65.0, drinkFrom:2024, drinkTo:2047, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1171, winery:"Austin Hope", name:"Austin Hope Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2023, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:15, slot:0, sticker:"green", price:55.0, drinkFrom:2025, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1172, winery:"Austin Hope", name:"Austin Hope Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2023, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:15, slot:1, sticker:"green", price:55.0, drinkFrom:2025, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1173, winery:"Austin Hope", name:"Austin Hope Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2023, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:15, slot:2, sticker:"green", price:55.0, drinkFrom:2025, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1174, winery:"Austin Hope", name:"Austin Hope Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2023, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:16, slot:0, sticker:"green", price:55.0, drinkFrom:2025, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1175, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 3", shelf:16, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1176, winery:"Austin Hope", name:"Austin Hope Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2023, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:16, slot:2, sticker:"green", price:55.0, drinkFrom:2025, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1177, winery:"Sextant Wines", name:"Kamal Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2017, region:"Paso Robles, California", country:"USA", row:"Row 3", shelf:17, slot:0, sticker:"green", price:45.0, drinkFrom:2020, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1178, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 3", shelf:17, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1179, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 3", shelf:17, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1180, winery:"Emerson Brown Wines", name:"Emerson Brown", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley, California", country:"USA", row:"Row 3", shelf:18, slot:0, sticker:"red", price:50.0, drinkFrom:2025, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1181, winery:"Emerson Brown Wines", name:"Emerson Brown", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley, California", country:"USA", row:"Row 3", shelf:18, slot:1, sticker:"red", price:50.0, drinkFrom:2025, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1182, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 3", shelf:18, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1183, winery:"Emerson Brown Wines", name:"Emerson Brown", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley, California", country:"USA", row:"Row 3", shelf:19, slot:0, sticker:"red", price:50.0, drinkFrom:2025, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1184, winery:"Emerson Brown Wines", name:"Emerson Brown", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley, California", country:"USA", row:"Row 3", shelf:19, slot:1, sticker:"red", price:50.0, drinkFrom:2025, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1185, winery:"Emerson Brown Wines", name:"Emerson Brown", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley, California", country:"USA", row:"Row 3", shelf:19, slot:2, sticker:"red", price:50.0, drinkFrom:2025, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1186, winery:"Emerson Brown Wines", name:"Emerson Brown", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley, California", country:"USA", row:"Row 3", shelf:20, slot:0, sticker:"red", price:50.0, drinkFrom:2025, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1187, winery:"Emerson Brown Wines", name:"Emerson Brown", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley, California", country:"USA", row:"Row 3", shelf:20, slot:1, sticker:"red", price:50.0, drinkFrom:2025, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1188, winery:"Emerson Brown Wines", name:"Emerson Brown", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley, California", country:"USA", row:"Row 3", shelf:20, slot:2, sticker:"red", price:50.0, drinkFrom:2025, drinkTo:2040, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1189, winery:"Sinegal Estate", name:"Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2021, region:"Napa Valley, California", country:"USA", row:"Row 4", shelf:0, slot:0, sticker:"red", price:85.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1190, winery:"Sinegal Estate", name:"Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2021, region:"Napa Valley, California", country:"USA", row:"Row 4", shelf:0, slot:1, sticker:"red", price:85.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1191, winery:"Sinegal Estate", name:"Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2021, region:"Napa Valley, California", country:"USA", row:"Row 4", shelf:0, slot:2, sticker:"red", price:85.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1192, winery:"Michele Satta", name:"Piastraia Bolgheri Superiore", varietal:"Merlot", vintage:2015, region:"Bolgheri, Tuscany", country:"Italy", row:"Row 4", shelf:1, slot:0, sticker:"yellow", price:50.0, drinkFrom:2018, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1193, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 4", shelf:1, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1194, winery:"L'Aventure", name:"Optimus Estate Proprietary Red", varietal:"Cabernet Sauvignon", vintage:2010, region:"Paso Robles", country:"USA", row:"Row 4", shelf:1, slot:2, sticker:"yellow", price:80.0, drinkFrom:2015, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1195, winery:"Crous St-Martin", name:"Crous St-Martin “Les Espaliers” Gigondas", varietal:"Syrah/Shiraz", vintage:2023, region:"Gigondas, Rhone Valley", country:"France", row:"Row 4", shelf:2, slot:0, sticker:"green", price:28.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1196, winery:"Crous St-Martin", name:"Crous St-Martin “Les Espaliers” Gigondas", varietal:"Syrah/Shiraz", vintage:2023, region:"Gigondas, Rhone Valley", country:"France", row:"Row 4", shelf:2, slot:1, sticker:"green", price:28.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1197, winery:"Crous St-Martin", name:"Crous St-Martin “Les Espaliers” Gigondas", varietal:"Syrah/Shiraz", vintage:2023, region:"Gigondas, Rhone Valley", country:"France", row:"Row 4", shelf:2, slot:2, sticker:"green", price:28.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1198, winery:"Domaine Saint‑Damien", name:"Domaine Saint‑Damien Gigondas", varietal:"Other", vintage:2017, region:"Southern Rhône", country:"France", row:"Row 4", shelf:3, slot:0, sticker:"green", price:45.0, drinkFrom:2017, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1199, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 4", shelf:3, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1200, winery:"Vena Cava", name:"Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2020, region:"Valle de Guadalupe", country:"Mexico", row:"Row 4", shelf:3, slot:2, sticker:"green", price:38.0, drinkFrom:2022, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1201, winery:"Michel Rolland", name:"Clos de los Siete", varietal:"Other", vintage:2018, region:"Valle de Uco, Mendoza", country:"Argentina", row:"Row 4", shelf:4, slot:0, sticker:"green", price:25.0, drinkFrom:2021, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1202, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 4", shelf:4, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1203, winery:"Michel Rolland", name:"Clos de los Siete", varietal:"Other", vintage:2018, region:"Valle de Uco, Mendoza", country:"Argentina", row:"Row 4", shelf:4, slot:2, sticker:"green", price:25.0, drinkFrom:2021, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1204, winery:"Hartwell Estate", name:"Misté Hill Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2008, region:"Stags Leap, Napa Valley", country:"USA", row:"Row 4", shelf:5, slot:0, sticker:"black", price:120.0, drinkFrom:2012, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1205, winery:"Silverado Vineyards", name:"Silverado Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2018, region:"Napa Valley", country:"USA", row:"Row 4", shelf:5, slot:1, sticker:"green", price:47.0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1206, winery:"Vacqueyras", name:"Domaine Saint-Pierre Vacqueyras", varietal:"Other", vintage:2007, region:"Vacqueyras, Southern Rhone", country:"France", row:"Row 4", shelf:6, slot:0, sticker:"black", price:30.0, drinkFrom:2020, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1207, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 4", shelf:6, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1208, winery:"Caymus Vineyards", name:"Caymus Cabernet (1L)", varietal:"Cabernet Sauvignon", vintage:2017, region:"Napa Valley, California", country:"USA", row:"Row 4", shelf:7, slot:0, sticker:"blue", price:110.0, drinkFrom:2020, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1209, winery:"Pride Mountain Vineyards", name:"Pride Cabernet", varietal:"Cabernet Sauvignon", vintage:2018, region:"Spring Mountain, CA", country:"USA", row:"Row 4", shelf:8, slot:0, sticker:"red", price:85.0, drinkFrom:2022, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1210, winery:"Law Estate Wines", name:"The Nies", varietal:"Grenache", vintage:2022, region:"Dry Creek Valley, Sonoma County", country:"USA", row:"Row 5", shelf:0, slot:0, sticker:"yellow", price:115.0, drinkFrom:2023, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1211, winery:"Law Estate Wines", name:"The Nies", varietal:"Grenache", vintage:2022, region:"Dry Creek Valley, Sonoma County", country:"USA", row:"Row 5", shelf:0, slot:1, sticker:"yellow", price:115.0, drinkFrom:2023, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1212, winery:"Law Estate Wines", name:"The Nies", varietal:"Grenache", vintage:2022, region:"Dry Creek Valley, Sonoma County", country:"USA", row:"Row 5", shelf:0, slot:2, sticker:"yellow", price:115.0, drinkFrom:2023, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1213, winery:"Law Estate Wines", name:"The Nies", varietal:"Grenache", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 5", shelf:1, slot:0, sticker:"yellow", price:115.0, drinkFrom:2023, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1214, winery:"Law Estate Wines", name:"The Nies", varietal:"Grenache", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 5", shelf:1, slot:1, sticker:"yellow", price:115.0, drinkFrom:2023, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1215, winery:"Law Estate Wines", name:"The Nies", varietal:"Grenache", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 5", shelf:1, slot:2, sticker:"yellow", price:115.0, drinkFrom:2023, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1216, winery:"Caymus-Suisun", name:"Grand Durif", varietal:"Other", vintage:2018, region:"Susiun Valley", country:"USA", row:"Row 5", shelf:2, slot:0, sticker:"yellow", price:49.95, drinkFrom:2026, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1217, winery:"A. Rafanelli", name:"Zinfandel", varietal:"Zinfandel", vintage:2017, region:"Dry Creek Valley, Sonoma County", country:"USA", row:"Row 5", shelf:2, slot:1, sticker:"yellow", price:55.0, drinkFrom:2020, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1218, winery:"Law Estate Wines", name:"The Nies", varietal:"Grenache", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 5", shelf:2, slot:2, sticker:"yellow", price:115.0, drinkFrom:2023, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1219, winery:"Cotes-du-Rhone Villages", name:"Les Garrigues", varietal:"Syrah/Shiraz", vintage:2016, region:"Côtes‑du‑Rhône Villages", country:"France", row:"Row 5", shelf:3, slot:0, sticker:"green", price:35.0, drinkFrom:2024, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1220, winery:"Austin Hope", name:"Austin Hope Appelation Series", varietal:"Syrah/Shiraz", vintage:2021, region:"Santa Barbara County, CA", country:"USA", row:"Row 5", shelf:3, slot:1, sticker:"green", price:45.0, drinkFrom:2024, drinkTo:2029, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1221, winery:"Austin Hope", name:"Austin Hope Appelation Series", varietal:"Syrah/Shiraz", vintage:2021, region:"Santa Barbara County, CA", country:"USA", row:"Row 5", shelf:3, slot:2, sticker:"green", price:45.0, drinkFrom:2024, drinkTo:2029, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1222, winery:"Seghesio Family Vineyards", name:"Seghesio Zinfandel", varietal:"Zinfandel", vintage:2020, region:"Sonoma County", country:"USA", row:"Row 5", shelf:4, slot:0, sticker:"green", price:25.0, drinkFrom:2022, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1223, winery:"Tackitt Family Vineyards", name:"Wild Rover Petite Sirah", varietal:"Zinfandel", vintage:2019, region:"Paso Robles, California", country:"USA", row:"Row 5", shelf:4, slot:1, sticker:"green", price:42.0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1224, winery:"Foggy Canyon", name:"Foggy Canyon Syrah", varietal:"Syrah/Shiraz", vintage:2017, region:"Russian River Valley, Sonoma County", country:"USA", row:"Row 5", shelf:4, slot:2, sticker:"green", price:45.0, drinkFrom:2022, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1225, winery:"Keys Creek Winery", name:"Keys Creek Winery Syrah", varietal:"Syrah/Shiraz", vintage:2015, region:"North San Diego County", country:"USA", row:"Row 5", shelf:5, slot:0, sticker:"black", price:40.0, drinkFrom:2018, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1226, winery:"Keys Creek Winery", name:"Keys Creek Winery Syrah", varietal:"Syrah/Shiraz", vintage:2015, region:"North San Diego County", country:"USA", row:"Row 5", shelf:5, slot:1, sticker:"black", price:40.0, drinkFrom:2018, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1227, winery:"Penfolds", name:"RWT Shiraz BIN 798", varietal:"Syrah/Shiraz", vintage:2016, region:"Barossa Valley", country:"Australia", row:"Row 5", shelf:6, slot:0, sticker:"blue", price:150.0, drinkFrom:2020, drinkTo:2035, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1228, winery:"Cotes-du-Rhone Villages", name:"Les Garrigues", varietal:"Syrah/Shiraz", vintage:2016, region:"Côtes‑du‑Rhône Villages", country:"France", row:"Row 5", shelf:6, slot:1, sticker:"green", price:35.0, drinkFrom:2024, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1229, winery:"By Farr (Tout Près)", name:"Tout Près Pinot Noir", varietal:"Pinot Noir", vintage:2020, region:"Geelong, Victoria", country:"Australia", row:"Row 5", shelf:7, slot:0, sticker:"blue", price:120.0, drinkFrom:2023, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1230, winery:"Yao Ming", name:"Cabernet Sauvignon", varietal:"Cabernet Sauvignon", vintage:2015, region:"Napa Valley", country:"USA", row:"Row 5", shelf:8, slot:0, sticker:"blue", price:95.0, drinkFrom:2020, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1231, winery:"Law Estate Wines", name:"Sagacious", varietal:"Syrah/Shiraz", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 6", shelf:0, slot:0, sticker:"red", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1232, winery:"Law Estate Wines", name:"Sagacious", varietal:"Syrah/Shiraz", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 6", shelf:0, slot:1, sticker:"red", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1233, winery:"Law Estate Wines", name:"Sagacious", varietal:"Syrah/Shiraz", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 6", shelf:0, slot:2, sticker:"red", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1234, winery:"Hickinbotham Clarendon", name:"Brooks Road Shiraz", varietal:"Syrah/Shiraz", vintage:2015, region:"McLaren Vale", country:"Australia", row:"Row 6", shelf:1, slot:0, sticker:"yellow", price:85.0, drinkFrom:2020, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1235, winery:"Hickinbotham Clarendon", name:"Brooks Road Shiraz", varietal:"Syrah/Shiraz", vintage:2015, region:"McLaren Vale", country:"Australia", row:"Row 6", shelf:1, slot:1, sticker:"yellow", price:85.0, drinkFrom:2020, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1236, winery:"Hickinbotham Clarendon", name:"Brooks Road Shiraz", varietal:"Syrah/Shiraz", vintage:2015, region:"McLaren Vale", country:"Australia", row:"Row 6", shelf:1, slot:2, sticker:"yellow", price:85.0, drinkFrom:2020, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1237, winery:"Nicorra", name:"Nicorra", varietal:"Syrah/Shiraz", vintage:2013, region:"Central Coast, California", country:"USA", row:"Row 6", shelf:2, slot:0, sticker:"green", price:45.0, drinkFrom:2015, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1238, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 6", shelf:2, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1239, winery:"Michael David Winery", name:"Petite Petit", varietal:"Other", vintage:2020, region:"Lodi", country:"USA", row:"Row 6", shelf:2, slot:2, sticker:"green", price:16.0, drinkFrom:2022, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1240, winery:"Z Wine", name:"Julius Shiraz", varietal:"Syrah/Shiraz", vintage:2019, region:"Barossa Valley", country:"Australia", row:"Row 6", shelf:3, slot:0, sticker:"green", price:40.0, drinkFrom:2021, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1241, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 6", shelf:3, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1242, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 6", shelf:3, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1243, winery:"Castello di Amorosa", name:"Napa Valley Sangiovese", varietal:"Merlot", vintage:2021, region:"Napa Valley", country:"USA", row:"Row 6", shelf:4, slot:0, sticker:"green", price:40.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1244, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 6", shelf:4, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1245, winery:"Castello di Amorosa", name:"Napa Valley Sangiovese", varietal:"Merlot", vintage:2021, region:"Napa Valley", country:"USA", row:"Row 6", shelf:4, slot:2, sticker:"green", price:40.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1246, winery:"Austin Hope", name:"The Magic Sun (Avery #3)", varietal:"Syrah/Shiraz", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 6", shelf:5, slot:0, sticker:"green", price:65.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1247, winery:"Justin Vineyards", name:"Right Angle", varietal:"Other", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 6", shelf:5, slot:1, sticker:"green", price:35.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1248, winery:"Montes", name:"Purple Angel", varietal:"Other", vintage:2015, region:"Colchagua Valley", country:"Chile", row:"Row 6", shelf:6, slot:0, sticker:"blue", price:70.0, drinkFrom:2018, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1249, winery:"North Coast Blend", name:"Syrah/Sirah Blend", varietal:"Syrah/Shiraz", vintage:2019, region:"North Coast, California", country:"USA", row:"Row 6", shelf:6, slot:1, sticker:"green", price:20.0, drinkFrom:2022, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1250, winery:"BRAND Napa Valley", name:"Proprietary Blend", varietal:"Other", vintage:2017, region:"Napa Valley, California", country:"USA", row:"Row 6", shelf:7, slot:0, sticker:"blue", price:150.0, drinkFrom:2022, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1251, winery:"Penfolds", name:"Grange Bin 95", varietal:"Syrah/Shiraz", vintage:2013, region:"South Australia", country:"Australia", row:"Row 6", shelf:8, slot:0, sticker:"blue", price:750.0, drinkFrom:2023, drinkTo:2043, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1252, winery:"Justin Vineyards", name:"Justin Syrah", varietal:"Syrah/Shiraz", vintage:2023, region:"Paso Robles, California", country:"USA", row:"Row 7", shelf:0, slot:0, sticker:"red", price:35.0, drinkFrom:2025, drinkTo:2031, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1253, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:0, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1254, winery:"Justin Vineyards", name:"Right Angle", varietal:"Other", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 7", shelf:0, slot:2, sticker:"red", price:35.0, drinkFrom:2025, drinkTo:2031, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1255, winery:"Law Estate Wines", name:"Aspire", varietal:"Syrah/Shiraz", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 7", shelf:1, slot:0, sticker:"red", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1256, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:1, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1257, winery:"Law Estate Wines", name:"Aspire", varietal:"Syrah/Shiraz", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 7", shelf:1, slot:2, sticker:"red", price:88.0, drinkFrom:2024, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1258, winery:"Aonair", name:"Aonair Heritage Clone", varietal:"Grenache", vintage:2020, region:"Napa Valley", country:"USA", row:"Row 7", shelf:2, slot:0, sticker:"red", price:68.0, drinkFrom:2023, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1259, winery:"Barreri & Rovati", name:"Barolo Riserva", varietal:"Nebbiolo", vintage:2016, region:"Piedmont – Barolo", country:"Italy", row:"Row 7", shelf:2, slot:1, sticker:"red", price:65.0, drinkFrom:2022, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1260, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:2, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1261, winery:"Carmelo Rodero", name:"Reserva Ribera del Duero", varietal:"Tempranillo", vintage:2018, region:"Ribera del Duero", country:"Spain", row:"Row 7", shelf:3, slot:0, sticker:"yellow", price:60.0, drinkFrom:2022, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1262, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:3, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1263, winery:"Carmelo Rodero", name:"Reserva Ribera del Duero", varietal:"Tempranillo", vintage:2018, region:"Ribera del Duero", country:"Spain", row:"Row 7", shelf:3, slot:2, sticker:"yellow", price:60.0, drinkFrom:2022, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1264, winery:"Cass Winery", name:"Rockin’ Ted", varietal:"Syrah/Shiraz", vintage:2017, region:"Paso Robles", country:"USA", row:"Row 7", shelf:4, slot:0, sticker:"yellow", price:35.0, drinkFrom:2020, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1265, winery:"Cass Winery", name:"Rockin’ Ted", varietal:"Syrah/Shiraz", vintage:2017, region:"Paso Robles", country:"USA", row:"Row 7", shelf:4, slot:1, sticker:"yellow", price:35.0, drinkFrom:2020, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1266, winery:"Cass Winery", name:"Rockin’ Ted", varietal:"Syrah/Shiraz", vintage:2017, region:"Paso Robles", country:"USA", row:"Row 7", shelf:4, slot:2, sticker:"yellow", price:35.0, drinkFrom:2020, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1267, winery:"Domaine des Sénéchaux", name:"Châteauneuf-du-Pape", varietal:"Syrah/Shiraz", vintage:2010, region:"Rhône Valley – Châteauneuf-du-Pape", country:"France", row:"Row 7", shelf:5, slot:0, sticker:"yellow", price:65.0, drinkFrom:2015, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1268, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:5, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1269, winery:"Cass Winery", name:"Rockin’ Ted", varietal:"Syrah/Shiraz", vintage:2017, region:"Paso Robles", country:"USA", row:"Row 7", shelf:5, slot:2, sticker:"yellow", price:35.0, drinkFrom:2020, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1270, winery:"Agricola Fontanelle", name:"Stiglianese II Toscana", varietal:"Sangiovese", vintage:2019, region:"Tuscany", country:"Italy", row:"Row 7", shelf:6, slot:0, sticker:"green", price:30.0, drinkFrom:2022, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1271, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:7, slot:0, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1272, winery:"Christopher Cameron Vineyards", name:"Christopher Cameron Merlot", varietal:"Merlot", vintage:2021, region:"Dry Creek Valley", country:"USA", row:"Row 7", shelf:7, slot:1, sticker:"green", price:20.0, drinkFrom:2022, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1273, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:7, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1274, winery:"De’ Ricci", name:"Il Vignone", varietal:"Sangiovese", vintage:2020, region:"Tuscany", country:"Italy", row:"Row 7", shelf:8, slot:0, sticker:"green", price:32.0, drinkFrom:2021, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1275, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:8, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1276, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:8, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1277, winery:"Pinot noir", name:"Elysian Red", varietal:"Other", vintage:2019, region:"Santa Barbara County", country:"USA", row:"Row 7", shelf:9, slot:0, sticker:"green", price:25.0, drinkFrom:2022, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1278, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:9, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1279, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:9, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1280, winery:"San Cassiano", name:"Tecnica Rosso Veneto", varietal:"Other", vintage:2018, region:"Veneto", country:"Italy", row:"Row 7", shelf:10, slot:0, sticker:"green", price:28.0, drinkFrom:2021, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1281, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:10, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1282, winery:"San Cassiano", name:"Tecnica Rosso Veneto", varietal:"Other", vintage:2018, region:"Veneto", country:"Italy", row:"Row 7", shelf:10, slot:2, sticker:"green", price:28.0, drinkFrom:2021, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1283, winery:"Charles Tyrand", name:"Rochegude Côtes-du-Rhône", varietal:"Syrah/Shiraz", vintage:2020, region:"Côtes-du-Rhône", country:"France", row:"Row 7", shelf:11, slot:0, sticker:"green", price:22.0, drinkFrom:2022, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1284, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:11, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1285, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:11, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1286, winery:"Castello di Amorosa", name:"Sangiovese", varietal:"Sangiovese", vintage:2018, region:"Napa Valley, California", country:"USA", row:"Row 7", shelf:12, slot:0, sticker:"green", price:39.0, drinkFrom:2021, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1287, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:12, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1288, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:12, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1289, winery:"Château de Saint Cosme", name:"Château de Saint Cosme Crozes‑Hermitage", varietal:"Syrah/Shiraz", vintage:2023, region:"Northern Rhône", country:"France", row:"Row 7", shelf:13, slot:0, sticker:"green", price:35.0, drinkFrom:2024, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1290, winery:"Château de Saint Cosme", name:"Château de Saint Cosme Crozes‑Hermitage", varietal:"Syrah/Shiraz", vintage:2023, region:"Northern Rhône", country:"France", row:"Row 7", shelf:13, slot:1, sticker:"green", price:35.0, drinkFrom:2024, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1291, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:13, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1292, winery:"Christopher Cameron Vineyards", name:"Christopher Cameron Merlot", varietal:"Merlot", vintage:2017, region:"Dry Creek Valley", country:"USA", row:"Row 7", shelf:14, slot:0, sticker:"green", price:20.0, drinkFrom:2022, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1293, winery:"Christopher Cameron Vineyards", name:"Christopher Cameron Merlot", varietal:"Merlot", vintage:2015, region:"Dry Creek Valley", country:"USA", row:"Row 7", shelf:14, slot:1, sticker:"green", price:20.0, drinkFrom:2020, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1294, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:14, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1295, winery:"Rusack Vineyards", name:"Syrah Estate", varietal:"Syrah/Shiraz", vintage:2019, region:"Ballard Canyon, Santa Barbara", country:"USA", row:"Row 7", shelf:15, slot:0, sticker:"green", price:45.0, drinkFrom:2022, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1296, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:15, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1297, winery:"Rusack Vineyards", name:"Syrah Estate", varietal:"Syrah/Shiraz", vintage:2019, region:"Ballard Canyon, Santa Barbara", country:"USA", row:"Row 7", shelf:15, slot:2, sticker:"green", price:45.0, drinkFrom:2022, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1298, winery:"Rusack Vineyards", name:"Syrah Estate", varietal:"Syrah/Shiraz", vintage:2019, region:"Ballard Canyon, Santa Barbara", country:"USA", row:"Row 7", shelf:16, slot:0, sticker:"green", price:45.0, drinkFrom:2022, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1299, winery:"Rusack Vineyards", name:"Syrah Estate", varietal:"Syrah/Shiraz", vintage:2019, region:"Ballard Canyon, Santa Barbara", country:"USA", row:"Row 7", shelf:16, slot:1, sticker:"green", price:45.0, drinkFrom:2022, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1300, winery:"Rusack Vineyards", name:"Syrah Estate", varietal:"Syrah/Shiraz", vintage:2019, region:"Ballard Canyon, Santa Barbara", country:"USA", row:"Row 7", shelf:16, slot:2, sticker:"green", price:45.0, drinkFrom:2022, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1301, winery:"Campaperi", name:"Rosso Toscano Campaperi", varietal:"Sangiovese", vintage:2019, region:"Tuscany", country:"Italy", row:"Row 7", shelf:17, slot:0, sticker:"green", price:11.0, drinkFrom:2021, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1302, winery:"Campaperi", name:"Rosso Toscano Campaperi", varietal:"Sangiovese", vintage:2019, region:"Tuscany", country:"Italy", row:"Row 7", shelf:17, slot:1, sticker:"green", price:11.0, drinkFrom:2021, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1303, winery:"Campaperi", name:"Rosso Toscano Campaperi", varietal:"Sangiovese", vintage:2019, region:"Tuscany", country:"Italy", row:"Row 7", shelf:17, slot:2, sticker:"green", price:11.0, drinkFrom:2021, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1304, winery:"Campaperi", name:"Rosso Toscano Campaperi", varietal:"Sangiovese", vintage:2019, region:"Tuscany", country:"Italy", row:"Row 7", shelf:18, slot:0, sticker:"green", price:11.0, drinkFrom:2021, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1305, winery:"Campaperi", name:"Rosso Toscano Campaperi", varietal:"Sangiovese", vintage:2019, region:"Tuscany", country:"Italy", row:"Row 7", shelf:18, slot:1, sticker:"green", price:11.0, drinkFrom:2021, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1306, winery:"Campaperi", name:"Rosso Toscano Campaperi", varietal:"Sangiovese", vintage:2019, region:"Tuscany", country:"Italy", row:"Row 7", shelf:18, slot:2, sticker:"green", price:11.0, drinkFrom:2021, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1307, winery:"Campaperi", name:"Rosso Toscano Campaperi", varietal:"Sangiovese", vintage:2019, region:"Tuscany", country:"Italy", row:"Row 7", shelf:19, slot:0, sticker:"green", price:11.0, drinkFrom:2021, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1308, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:19, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1309, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:19, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1310, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:20, slot:0, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1311, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:20, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1312, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 7", shelf:20, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1313, winery:"By Farr", name:"Tout Près by Farr Pinot Noir", varietal:"Pinot Noir", vintage:2020, region:"Geelong, Victoria", country:"Australia", row:"Row 8", shelf:0, slot:0, sticker:"blue", price:115.0, drinkFrom:2023, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1314, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 8", shelf:0, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1315, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 8", shelf:0, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1316, winery:"Justin Vineyards", name:"Justification", varietal:"Merlot", vintage:2022, region:"Paso Robles, California", country:"USA", row:"Row 8", shelf:1, slot:0, sticker:"red", price:70.0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1317, winery:"L'Aventure", name:"L'Aventure Estate Cuvée", varietal:"Other", vintage:2017, region:"Paso Robles", country:"USA", row:"Row 8", shelf:1, slot:1, sticker:"blue", price:118.0, drinkFrom:2021, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1318, winery:"Occidental Wines", name:"Freestone-Occidental Pinot Noir", varietal:"Pinot Noir", vintage:2022, region:"Sonoma Coast – Freestone-Occidental", country:"USA", row:"Row 8", shelf:1, slot:2, sticker:"red", price:80.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1319, winery:"Occidental Wines", name:"Freestone-Occidental Pinot Noir", varietal:"Pinot Noir", vintage:2022, region:"Sonoma Coast – Freestone-Occidental", country:"USA", row:"Row 8", shelf:2, slot:0, sticker:"red", price:80.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1320, winery:"Occidental Wines", name:"Freestone-Occidental Pinot Noir", varietal:"Pinot Noir", vintage:2022, region:"Sonoma Coast – Freestone-Occidental", country:"USA", row:"Row 8", shelf:2, slot:1, sticker:"red", price:80.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1321, winery:"Occidental Wines", name:"Freestone-Occidental Pinot Noir", varietal:"Pinot Noir", vintage:2022, region:"Sonoma Coast – Freestone-Occidental", country:"USA", row:"Row 8", shelf:2, slot:2, sticker:"red", price:80.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1322, winery:"Monticello Estate", name:"Monticello Cabernet Franc", varietal:"Cabernet Franc", vintage:2019, region:"Napa Valley, California", country:"USA", row:"Row 8", shelf:3, slot:0, sticker:"yellow", price:58.0, drinkFrom:2023, drinkTo:2031, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1323, winery:"Costa Azul", name:"Costa Azul Cabernet Franc", varietal:"Cabernet Franc", vintage:2022, region:"Clements Hills, Lodi, California", country:"USA", row:"Row 8", shelf:3, slot:1, sticker:"red", price:59.0, drinkFrom:2024, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1324, winery:"Louis Jadot", name:"Domaine Gagey Cent Vignes Beaune 1er Cru", varietal:"Pinot Noir", vintage:2015, region:"Beaune 1er Cru Appellation", country:"France", row:"Row 8", shelf:3, slot:2, sticker:"blue", price:102.0, drinkFrom:2025, drinkTo:2045, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1325, winery:"Justin Vineyards", name:"Justification", varietal:"Merlot", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Row 8", shelf:4, slot:0, sticker:"yellow", price:70.0, drinkFrom:2021, drinkTo:2031, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1326, winery:"Justin Vineyards", name:"Justification", varietal:"Merlot", vintage:2018, region:"Paso Robles, California", country:"USA", row:"Row 8", shelf:4, slot:1, sticker:"yellow", price:70.0, drinkFrom:2020, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1327, winery:"Justin Vineyards", name:"Justification", varietal:"Merlot", vintage:2019, region:"Paso Robles, California", country:"USA", row:"Row 8", shelf:4, slot:2, sticker:"yellow", price:70.0, drinkFrom:2021, drinkTo:2031, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1328, winery:"Austin Hope", name:"Austin Hope Cellar Select Graciano", varietal:"Other", vintage:2019, region:"Paso Robles, California", country:"USA", row:"Row 8", shelf:5, slot:0, sticker:"yellow", price:90.0, drinkFrom:2023, drinkTo:2034, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1329, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 8", shelf:5, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1330, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 8", shelf:5, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1331, winery:"DAOU Vineyards", name:"Daou Petaluma Gap", varietal:"Pinot Noir", vintage:2019, region:"Paso Robles", country:"USA", row:"Row 8", shelf:6, slot:0, sticker:"yellow", price:91.0, drinkFrom:2026, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1332, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 8", shelf:6, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1333, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 8", shelf:6, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1334, winery:"Hourglass", name:"HG III", varietal:"Other", vintage:2018, region:"Napa Valley", country:"USA", row:"Row 8", shelf:7, slot:0, sticker:"yellow", price:50.0, drinkFrom:2021, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1335, winery:"Duckhorn Winery", name:"Goldeneye", varietal:"Pinot Noir", vintage:2017, region:"Anderson Valley, CA", country:"USA", row:"Row 8", shelf:7, slot:1, sticker:"yellow", price:62.0, drinkFrom:2020, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1336, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 8", shelf:7, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1337, winery:"RAEN Winery", name:"Royal St. Robert Pinot Noir", varietal:"Pinot Noir", vintage:2019, region:"Sonoma Coast", country:"USA", row:"Row 8", shelf:8, slot:0, sticker:"yellow", price:75.0, drinkFrom:2021, drinkTo:2029, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1338, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 8", shelf:8, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1339, winery:"RAEN Winery", name:"Royal St. Robert Pinot Noir", varietal:"Pinot Noir", vintage:2019, region:"Sonoma Coast", country:"USA", row:"Row 8", shelf:8, slot:2, sticker:"yellow", price:75.0, drinkFrom:2021, drinkTo:2029, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1340, winery:"Costa Azul", name:"Costa Azul Pinot Noir", varietal:"Pinot Noir", vintage:2018, region:"Santa Lucia Highlands", country:"USA", row:"Row 8", shelf:9, slot:0, sticker:"blue", price:100.0, drinkFrom:2022, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1341, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 8", shelf:9, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1342, winery:"Costa Azul", name:"Costa Azul Pinot Noir", varietal:"Pinot Noir", vintage:2019, region:"Santa Lucia Highlands", country:"USA", row:"Row 8", shelf:9, slot:2, sticker:"blue", price:100.0, drinkFrom:2023, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1343, winery:"Pietrose", name:"Pietrose Magnum", varietal:"Other", vintage:2013, region:"Unknown", country:"Unknown", row:"Row 8", shelf:10, slot:0, sticker:"yellow", price:70.0, drinkFrom:2017, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1344, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 8", shelf:10, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1345, winery:"Pietrose", name:"Pietrose Magnum", varietal:"Other", vintage:2013, region:"Unknown", country:"Unknown", row:"Row 8", shelf:10, slot:2, sticker:"yellow", price:70.0, drinkFrom:2017, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1346, winery:"Poggio Leano", name:"Poggio Leano", varietal:"Sangiovese", vintage:2022, region:"Temecula, California", country:"USA", row:"Row 8", shelf:12, slot:0, sticker:"green", price:35.0, drinkFrom:2024, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1347, winery:"Buena Vista", name:"Buena Vista Winemaker’s Cuvée Pinot", varietal:"Pinot Noir", vintage:2019, region:"Sonoma County", country:"USA", row:"Row 8", shelf:12, slot:1, sticker:"green", price:14.0, drinkFrom:2022, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1348, winery:"Christopher Cameron Vineyards", name:"Christopher Cameron Malbec", varietal:"Malbec", vintage:2022, region:"Dry Creek Valley, Sonoma Vounty", country:"USA", row:"Row 8", shelf:12, slot:2, sticker:"green", price:38.0, drinkFrom:2024, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1349, winery:"Justin Vineyards", name:"Justin Merlot", varietal:"Merlot", vintage:2019, region:"Paso Robles, California", country:"USA", row:"Row 8", shelf:13, slot:0, sticker:"green", price:30.0, drinkFrom:2019, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1350, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 8", shelf:13, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1351, winery:"The Paring", name:"The Paring Red", varietal:"Other", vintage:2014, region:"California", country:"USA", row:"Row 8", shelf:13, slot:2, sticker:"green", price:28.0, drinkFrom:2019, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1352, winery:"Fallbrook Winery", name:"5150 Dangerously Delicious Red Wine", varietal:"Other", vintage:2016, region:"North San Diego County", country:"USA", row:"Row 8", shelf:14, slot:0, sticker:"black", price:20.0, drinkFrom:2020, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1353, winery:"Pinot noir", name:"Elysian Red Blend", varietal:"Other", vintage:2019, region:"Santa Barbara County", country:"USA", row:"Row 8", shelf:14, slot:1, sticker:"green", price:42.0, drinkFrom:2022, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1354, winery:"Alamos", name:"Malbec", varietal:"Malbec", vintage:2022, region:"Mendoza", country:"Argentina", row:"Row 8", shelf:14, slot:2, sticker:"green", price:15.0, drinkFrom:2023, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1355, winery:"Wagner Family of Wine", name:"Red Schooner Voyage 7", varietal:"Malbec", vintage:2020, region:"Napa Valley (fruit from Argentina)", country:"USA", row:"Row 8", shelf:15, slot:0, sticker:"green", price:47.49, drinkFrom:2022, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1356, winery:"DeLoach Vineyards", name:"De Loach", varietal:"Pinot Noir", vintage:2013, region:"Santa Rosa, Callifornia", country:"USA", row:"Row 8", shelf:15, slot:1, sticker:"black", price:27.0, drinkFrom:2016, drinkTo:2020, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1357, winery:"Austin Hope", name:"Pinot Noir Cellar Select", varietal:"Pinot Noir", vintage:2022, region:"Central Coast, California", country:"USA", row:"Row 8", shelf:15, slot:2, sticker:"green", price:45.0, drinkFrom:2024, drinkTo:2028, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1358, winery:"Vicente Faria", name:"Animus Douro Red", varietal:"Other", vintage:2016, region:"Douro Valley", country:"Portugal", row:"Row 8", shelf:16, slot:0, sticker:"green", price:15.0, drinkFrom:2018, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1359, winery:"Oak Mountain Winery", name:"Malbec", varietal:"Malbec", vintage:2019, region:"Temecula Valley", country:"USA", row:"Row 8", shelf:16, slot:1, sticker:"green", price:20.0, drinkFrom:2021, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1360, winery:"Domaine Manigley", name:"Rully Les Chauchoux Monopole", varietal:"Pinot Noir", vintage:2017, region:"Burgundy – Rully", country:"France", row:"Row 8", shelf:16, slot:2, sticker:"green", price:40.0, drinkFrom:2021, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1361, winery:"Domaine Besancenot", name:"Beaune Theurons Premier Cru", varietal:"Pinot Noir", vintage:1994, region:"Burgundy – Beaune", country:"France", row:"Row 8", shelf:17, slot:0, sticker:"black", price:130.0, drinkFrom:2025, drinkTo:2010, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1362, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 8", shelf:17, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1363, winery:"Domaine Besancenot", name:"Beaune Theurons Premier Cru", varietal:"Pinot Noir", vintage:1994, region:"Burgundy – Beaune", country:"France", row:"Row 8", shelf:17, slot:2, sticker:"black", price:130.0, drinkFrom:2025, drinkTo:2010, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1364, winery:"Claudia", name:"Tesela Red Blend", varietal:"Other", vintage:2018, region:"Ribera del Duero", country:"Spain", row:"Row 8", shelf:18, slot:0, sticker:"green", price:25.0, drinkFrom:2020, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1365, winery:"Decoy (Duckhorn Portfolio)", name:"Merlot", varietal:"Merlot", vintage:2019, region:"California", country:"USA", row:"Row 8", shelf:18, slot:1, sticker:"green", price:20.0, drinkFrom:2022, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1366, winery:"Laird Family Estate", name:"Jillian’s Blend", varietal:"Other", vintage:2017, region:"Napa Valley", country:"USA", row:"Row 8", shelf:18, slot:2, sticker:"green", price:25.0, drinkFrom:2020, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1367, winery:"Niner Wine Estates", name:"Niner Pinot Noir", varietal:"Pinot Noir", vintage:2018, region:"Edna Valley", country:"USA", row:"Row 8", shelf:19, slot:0, sticker:"green", price:35.0, drinkFrom:2022, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1368, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 8", shelf:19, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1369, winery:"Niner Wine Estates", name:"Niner Pinot Noir", varietal:"Pinot Noir", vintage:2018, region:"Edna Valley", country:"USA", row:"Row 8", shelf:19, slot:2, sticker:"green", price:35.0, drinkFrom:2022, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1370, winery:"Twomey Cellars", name:"Twomey Pinot Noir", varietal:"Pinot Noir", vintage:2021, region:"Russian River Valley", country:"USA", row:"Row 9", shelf:0, slot:0, sticker:"red", price:30.0, drinkFrom:2023, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1371, winery:"Twomey Cellars", name:"Twomey Pinot Noir", varietal:"Pinot Noir", vintage:2021, region:"Russian River Valley", country:"USA", row:"Row 9", shelf:0, slot:1, sticker:"red", price:30.0, drinkFrom:2023, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1372, winery:"Twomey Cellars", name:"Twomey Pinot Noir", varietal:"Pinot Noir", vintage:2021, region:"Russian River Valley", country:"USA", row:"Row 9", shelf:0, slot:2, sticker:"red", price:30.0, drinkFrom:2023, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1373, winery:"Twomey Cellars", name:"Twomey Pinot Noir", varietal:"Pinot Noir", vintage:2021, region:"Russian River Valley", country:"USA", row:"Row 9", shelf:1, slot:0, sticker:"red", price:30.0, drinkFrom:2023, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1374, winery:"Twomey Cellars", name:"Twomey Pinot Noir", varietal:"Pinot Noir", vintage:2021, region:"Russian River Valley", country:"USA", row:"Row 9", shelf:1, slot:1, sticker:"red", price:30.0, drinkFrom:2023, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1375, winery:"Twomey Cellars", name:"Twomey Pinot Noir", varietal:"Pinot Noir", vintage:2021, region:"Russian River Valley", country:"USA", row:"Row 9", shelf:1, slot:2, sticker:"red", price:30.0, drinkFrom:2023, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1376, winery:"Law Estate Wines", name:"Soph by Law White Wine", varietal:"Other", vintage:2023, region:"Paso Robles", country:"USA", row:"Row 9", shelf:2, slot:0, sticker:"yellow", price:88.0, drinkFrom:2025, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1377, winery:"Law Estate Wines", name:"Soph by Law White Wine", varietal:"Other", vintage:2023, region:"Paso Robles", country:"USA", row:"Row 9", shelf:2, slot:1, sticker:"yellow", price:88.0, drinkFrom:2025, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1378, winery:"Law Estate Wines", name:"Soph by Law White Wine", varietal:"Other", vintage:2023, region:"Paso Robles", country:"USA", row:"Row 9", shelf:2, slot:2, sticker:"yellow", price:88.0, drinkFrom:2025, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1379, winery:"Law Estate Wines", name:"Rosé", varietal:"Grenache", vintage:2025, region:"Paso Robles", country:"USA", row:"Row 9", shelf:3, slot:0, sticker:"yellow", price:88.0, drinkFrom:2024, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1380, winery:"Law Estate Wines", name:"Soph by Law White Wine", varietal:"Other", vintage:2022, region:"Paso Robles", country:"USA", row:"Row 9", shelf:3, slot:1, sticker:"yellow", price:88.0, drinkFrom:2025, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1381, winery:"Law Estate Wines", name:"Soph by Law White Wine", varietal:"Other", vintage:2022, region:"Paso Robles", country:"USA", row:"Row 9", shelf:3, slot:2, sticker:"yellow", price:88.0, drinkFrom:2025, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1382, winery:"Law Estate Wines", name:"Soph by Law White Wine", varietal:"Other", vintage:2024, region:"Paso Robles", country:"USA", row:"Row 9", shelf:4, slot:0, sticker:"yellow", price:88.0, drinkFrom:2025, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1383, winery:"Law Estate Wines", name:"Soph by Law White Wine", varietal:"Other", vintage:2024, region:"Paso Robles", country:"USA", row:"Row 9", shelf:4, slot:1, sticker:"yellow", price:88.0, drinkFrom:2025, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1384, winery:"Law Estate Wines", name:"Soph by Law White Wine", varietal:"Other", vintage:2024, region:"Paso Robles", country:"USA", row:"Row 9", shelf:4, slot:2, sticker:"yellow", price:88.0, drinkFrom:2025, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1385, winery:"Croix Estate", name:"Croix Narrow Gauge Pinot Noir", varietal:"Pinot Noir", vintage:2016, region:"Fulton, California", country:"USA", row:"Row 9", shelf:5, slot:0, sticker:"yellow", price:60.0, drinkFrom:2017, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1386, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 9", shelf:5, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1387, winery:"Croix Estate", name:"Croix Narrow Gauge Pinot Noir", varietal:"Pinot Noir", vintage:2016, region:"Fulton, California", country:"USA", row:"Row 9", shelf:5, slot:2, sticker:"yellow", price:60.0, drinkFrom:2017, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1388, winery:"Law Estate Wines", name:"Chardonnay Law", varietal:"Chardonnay", vintage:2024, region:"Paso Robles", country:"USA", row:"Row 9", shelf:6, slot:0, sticker:"yellow", price:88.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1389, winery:"Law Estate Wines", name:"Chardonnay Law", varietal:"Chardonnay", vintage:2024, region:"Paso Robles", country:"USA", row:"Row 9", shelf:6, slot:1, sticker:"yellow", price:88.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1390, winery:"Law Estate Wines", name:"Chardonnay Law", varietal:"Chardonnay", vintage:2024, region:"Paso Robles", country:"USA", row:"Row 9", shelf:6, slot:2, sticker:"yellow", price:88.0, drinkFrom:2024, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1391, winery:"Martin Ray", name:"Martin Ray Pinot Noir", varietal:"Pinot Noir", vintage:2017, region:"Sonoma Coast", country:"USA", row:"Row 9", shelf:7, slot:0, sticker:"green", price:30.0, drinkFrom:2021, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1392, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 9", shelf:7, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1393, winery:"Kendall-Jackson", name:"Grand Reserve Chardonnay", varietal:"Chardonnay", vintage:2019, region:"Santa Barbara County", country:"USA", row:"Row 9", shelf:7, slot:2, sticker:"green", price:22.0, drinkFrom:2021, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1394, winery:"Bluebell", name:"Reserve Bluebell Pinot Noir", varietal:"Pinot Noir", vintage:2016, region:"Willamette Valley", country:"USA", row:"Row 9", shelf:8, slot:0, sticker:"green", price:35.0, drinkFrom:2020, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1395, winery:"Bluebell", name:"Reserve Bluebell Pinot Noir", varietal:"Pinot Noir", vintage:2016, region:"Willamette Valley", country:"USA", row:"Row 9", shelf:8, slot:1, sticker:"green", price:35.0, drinkFrom:2020, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1396, winery:"Bluebell", name:"Reserve Bluebell Pinot Noir", varietal:"Pinot Noir", vintage:2016, region:"Willamette Valley", country:"USA", row:"Row 9", shelf:8, slot:2, sticker:"green", price:35.0, drinkFrom:2020, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1397, winery:"Merry Edwards Winery", name:"Merry Edwards Sauvignon Blanc", varietal:"Sauvignon Blanc", vintage:2019, region:"Russian River Valley", country:"USA", row:"Row 9", shelf:9, slot:0, sticker:"green", price:40.0, drinkFrom:2020, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1398, winery:"Merry Edwards Winery", name:"Merry Edwards Sauvignon Blanc", varietal:"Sauvignon Blanc", vintage:2019, region:"Russian River Valley", country:"USA", row:"Row 9", shelf:9, slot:1, sticker:"green", price:40.0, drinkFrom:2020, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1399, winery:"Merry Edwards Winery", name:"Merry Edwards Sauvignon Blanc", varietal:"Sauvignon Blanc", vintage:2019, region:"Russian River Valley", country:"USA", row:"Row 9", shelf:9, slot:2, sticker:"green", price:40.0, drinkFrom:2020, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1400, winery:"Merry Edwards Winery", name:"Merry Edwards Sauvignon Blanc", varietal:"Sauvignon Blanc", vintage:2019, region:"Russian River Valley", country:"USA", row:"Row 9", shelf:10, slot:0, sticker:"green", price:40.0, drinkFrom:2020, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1401, winery:"Merry Edwards Winery", name:"Merry Edwards Sauvignon Blanc", varietal:"Sauvignon Blanc", vintage:2019, region:"Russian River Valley", country:"USA", row:"Row 9", shelf:10, slot:1, sticker:"green", price:40.0, drinkFrom:2020, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1402, winery:"Merry Edwards Winery", name:"Merry Edwards Sauvignon Blanc", varietal:"Sauvignon Blanc", vintage:2019, region:"Russian River Valley", country:"USA", row:"Row 9", shelf:10, slot:2, sticker:"green", price:40.0, drinkFrom:2020, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1403, winery:"Rombauer Vineyards", name:"Rombauer Carneros Chardonnay", varietal:"Chardonnay", vintage:2021, region:"Carneros", country:"USA", row:"Row 9", shelf:11, slot:0, sticker:"green", price:40.0, drinkFrom:2022, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1404, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 9", shelf:11, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1405, winery:"Rombauer Vineyards", name:"Rombauer Carneros Chardonnay", varietal:"Chardonnay", vintage:2021, region:"Carneros", country:"USA", row:"Row 9", shelf:11, slot:2, sticker:"green", price:40.0, drinkFrom:2022, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1406, winery:"Rombauer Vineyards", name:"Rombauer Carneros Chardonnay", varietal:"Chardonnay", vintage:2021, region:"Carneros", country:"USA", row:"Row 9", shelf:12, slot:0, sticker:"green", price:40.0, drinkFrom:2022, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1407, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 9", shelf:12, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1408, winery:"Rombauer Vineyards", name:"Rombauer Carneros Chardonnay", varietal:"Chardonnay", vintage:2021, region:"Carneros", country:"USA", row:"Row 9", shelf:12, slot:2, sticker:"green", price:40.0, drinkFrom:2022, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1409, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 9", shelf:13, slot:0, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1410, winery:"Austin Hope", name:"Austin Hope Chardonnay", varietal:"Chardonnay", vintage:2022, region:"Central Coast", country:"USA", row:"Row 9", shelf:13, slot:1, sticker:"green", price:49.0, drinkFrom:2023, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1411, winery:"Hope Family Wines", name:"Treana Blanc", varietal:"Other", vintage:2020, region:"Paso Robles", country:"USA", row:"Row 9", shelf:13, slot:2, sticker:"black", price:25.0, drinkFrom:2021, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1412, winery:"Kendall-Jackson", name:"Vintner's Reserve Chardonnay", varietal:"Chardonnay", vintage:2019, region:"California", country:"USA", row:"Row 9", shelf:14, slot:0, sticker:"green", price:13.0, drinkFrom:2021, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1413, winery:"Kendall-Jackson", name:"Vintner's Reserve Chardonnay", varietal:"Chardonnay", vintage:2020, region:"California", country:"USA", row:"Row 9", shelf:14, slot:1, sticker:"green", price:13.0, drinkFrom:2021, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1414, winery:"Kendall-Jackson", name:"Vintner's Reserve Chardonnay", varietal:"Chardonnay", vintage:2020, region:"California", country:"USA", row:"Row 9", shelf:14, slot:2, sticker:"green", price:13.0, drinkFrom:2021, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1415, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 9", shelf:15, slot:0, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1416, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 9", shelf:15, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1417, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 9", shelf:15, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1418, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 9", shelf:16, slot:0, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1419, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 9", shelf:16, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1420, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 9", shelf:16, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1421, winery:"La Marca", name:"La Marca Prosecco", varietal:"Rosé", vintage:2020, region:"Veneto", country:"Italy", row:"Row 9", shelf:17, slot:0, sticker:"green", price:14.0, drinkFrom:2022, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1422, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 9", shelf:17, slot:1, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1423, winery:"", name:"", varietal:"Other", vintage:2020, region:"", country:"", row:"Row 9", shelf:17, slot:2, sticker:"green", price:0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1424, winery:"Segura Viudas", name:"Segura Viudas Brut", varietal:"Champagne/Sparkling", vintage:2021, region:"Penedès", country:"Spain", row:"Row 9", shelf:18, slot:0, sticker:"green", price:10.0, drinkFrom:2023, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1425, winery:"Champagne Taittinger", name:"Taittinger Brut La Française", varietal:"Pinot Noir", vintage:2021, region:"Reims, Champagne", country:"France", row:"Row 9", shelf:18, slot:1, sticker:"yellow", price:55.0, drinkFrom:2023, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1426, winery:"Lancelot-Pienne", name:"Lancelot-Pienne Accord Majeur Brut", varietal:"Champagne/Sparkling", vintage:2020, region:"Champagne", country:"France", row:"Row 9", shelf:18, slot:2, sticker:"yellow", price:50.0, drinkFrom:2022, drinkTo:2027, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1427, winery:"Korbel", name:"Korbel Brut", varietal:"Chardonnay", vintage:2021, region:"Sonoma County, CA", country:"USA", row:"Row 9", shelf:19, slot:0, sticker:"green", price:15.0, drinkFrom:2023, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1428, winery:"Mumm Napa", name:"Mumm Napa Brut Prestige", varietal:"Pinot Noir", vintage:2021, region:"Napa Valley, California", country:"USA", row:"Row 9", shelf:19, slot:1, sticker:"green", price:20.0, drinkFrom:2023, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1429, winery:"Mumm Napa", name:"Brut Rosé", varietal:"Rosé", vintage:2020, region:"Napa Valley, California", country:"USA", row:"Row 9", shelf:19, slot:2, sticker:"green", price:20.0, drinkFrom:2022, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1430, winery:"Riondo Collezione", name:"Riondo Collezione", varietal:"Rosé", vintage:2020, region:"Veneto", country:"Italy", row:"Row 9", shelf:20, slot:0, sticker:"green", price:12.0, drinkFrom:2022, drinkTo:2025, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1431, winery:"Veuve Clicquot Ponsardin", name:"Veuve Clicquot", varietal:"Champagne/Sparkling", vintage:2020, region:"Reims, Champagne", country:"France", row:"Row 9", shelf:20, slot:1, sticker:"green", price:50.0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1432, winery:"Lois Roederer", name:"Louis Roederer Collection 242", varietal:"Champagne/Sparkling", vintage:2020, region:"Reims, Champagne", country:"France", row:"Row 9", shelf:20, slot:2, sticker:"green", price:50.0, drinkFrom:2022, drinkTo:2030, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1433, winery:"Ramey Wine Cellar", name:"Ramey Magnum", varietal:"Cabernet Sauvignon", vintage:2014, region:"Napa Valley, California", country:"USA", row:"Magnum A", shelf:0, slot:0, sticker:"yellow", price:60.0, drinkFrom:2028, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1434, winery:"Tenute Silvio Nardi", name:"Tenute Silvio Nardi Magnum", varietal:"Nebbiolo", vintage:2005, region:"Tuscany, Italy", country:"Italy", row:"Magnum A", shelf:1, slot:0, sticker:"yellow", price:99.0, drinkFrom:2013, drinkTo:2015, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1435, winery:"Justin Vineyards", name:"Isosceles Magnum", varietal:"Other", vintage:2021, region:"Paso Robles, California", country:"USA", row:"Magnum A", shelf:2, slot:0, sticker:"yellow", price:85.0, drinkFrom:2025, drinkTo:2038, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1436, winery:"Wagner Family of Wine", name:"Red Schooner Voyage 9 Magnum", varietal:"Malbec", vintage:2009, region:"Napa Valley, California", country:"USA", row:"Magnum A", shelf:3, slot:0, sticker:"yellow", price:60.0, drinkFrom:2022, drinkTo:2026, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1437, winery:"Ramey Wine Cellar", name:"Ramey Magnum", varietal:"Cabernet Sauvignon", vintage:2014, region:"Napa Valley, California", country:"USA", row:"Magnum B", shelf:0, slot:0, sticker:"yellow", price:60.0, drinkFrom:2028, drinkTo:2033, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1438, winery:"Tenute Silvio Nardi", name:"Tenute Silvio Nardi Magnum", varietal:"Nebbiolo", vintage:2005, region:"Tuscany, Italy", country:"Italy", row:"Magnum B", shelf:1, slot:0, sticker:"yellow", price:99.0, drinkFrom:2013, drinkTo:2015, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1439, winery:"DAOU Vineyards", name:"Daou Reserve Magnum", varietal:"Cabernet Sauvignon", vintage:2017, region:"Paso Robles, California", country:"USA", row:"Magnum B", shelf:2, slot:0, sticker:"yellow", price:99.0, drinkFrom:2022, drinkTo:2047, quantity:1, rating:0, notes:"", occasion:[] },
  { id:1440, winery:"Chateau de Marsan", name:"Cotes de Bordeaux", varietal:"Other", vintage:2022, region:"", country:"", row:"Magnum B", shelf:3, slot:0, sticker:"green", price:0, drinkFrom:2024, drinkTo:2032, quantity:1, rating:0, notes:"", occasion:[] },
];

// Total bottles: 317

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const currentYear = new Date().getFullYear();
const isRTD      = w => w.drinkFrom <= currentYear && w.drinkTo >= currentYear;
const isNeedsAge = w => w.drinkFrom > currentYear;
const isPastPeak = w => w.drinkTo < currentYear;

// Rich drink status — used in detail modal and wine list
function drinkStatus(w) {
  if (w.drinkTo < currentYear) {
    const yrs = currentYear - w.drinkTo;
    return { label: `Past Peak (${yrs}yr${yrs!==1?"s":""} ago)`, short: "Past Peak", color: "#a93226", icon: "⚠️" };
  }
  if (w.drinkFrom > currentYear) {
    const yrs = w.drinkFrom - currentYear;
    return { label: `Cellar ${yrs} more yr${yrs!==1?"s":""}`, short: "Needs Aging", color: "#d4a017", icon: "⏳" };
  }
  const yrsLeft = w.drinkTo - currentYear;
  if (yrsLeft <= 1) return { label: "Drink Now — peak ending soon!", short: "Drink Now!", color: "#c0392b", icon: "🚨" };
  if (yrsLeft <= 3) return { label: `Ready — ${yrsLeft} yrs left`, short: "Drink Soon", color: "#e67e22", icon: "🔔" };
  return { label: `Ready — peak through ${w.drinkTo}`, short: "Ready", color: "#1e8449", icon: "✓" };
}
function autoSticker(price, drinkFrom, drinkTo) {
  if (drinkTo < currentYear) return "black";
  if (drinkFrom > currentYear) return "red";
  if (price >= 100) return "blue";
  if (price >= 50) return "yellow";
  return "green";
}
function scoreColor(s) {
  if (s >= 97) return "#7b0d1e"; if (s >= 94) return "#a93226";
  if (s >= 90) return "#ca6f1e"; if (s >= 87) return "#b7950b";
  return "#616a6b";
}

// ─── SLOT KEY ────────────────────────────────────────────────────────────────
const slotKey = (row, shelf, slot) => `${row}__${shelf}__${slot}`;

// Build a slot map from wines array
function buildSlotMap(wines) {
  const map = {};
  wines.forEach(w => { map[slotKey(w.row, w.shelf, w.slot)] = w; });
  return map;
}

// ─── STYLES ───────────────────────────────────────────────────────────────────
const C = {
  bg:       "#0a0604",
  surface:  "#140a02",
  card:     "#1c0e04",
  border:   "#2e1608",
  gold:     "#c9a84c",
  goldDim:  "#7a6028",
  text:     "#e8dcc8",
  muted:    "#7a6848",
  dim:      "#3a2810",
};
const inp  = { background:"#0a0604", border:"1px solid #2e1608", color:"#e8dcc8", padding:"9px 12px", borderRadius:6, fontSize:13, fontFamily:"inherit", width:"100%", boxSizing:"border-box", outline:"none" };
const chip = (on, accent="#c9a84c") => ({ background: on ? `${accent}22` : "#1c0e04", border:`1px solid ${on ? accent : "#2e1608"}`, color: on ? accent : "#7a6848", padding:"5px 13px", borderRadius:20, cursor:"pointer", fontSize:12, fontFamily:"inherit", transition:"all 0.15s" });
const btn  = (disabled=false, color="#8b1a1a") => ({ background: disabled ? "#3a2810" : color, border:"none", color: disabled ? "#7a6848" : "#fff", padding:"11px 22px", borderRadius:7, cursor: disabled ? "not-allowed" : "pointer", fontSize:14, fontFamily:"inherit", fontWeight:700, letterSpacing:.5, transition:"opacity 0.15s" });

// ─── NAV ──────────────────────────────────────────────────────────────────────
const NAV = [
  {id:"map",     icon:"🏛️", label:"Cellar Map"},
  {id:"cellar",  icon:"🍷", label:"My Cellar"},
  {id:"add",     icon:"＋",  label:"Add Wine"},
  {id:"search",  icon:"🔍", label:"Search"},
  {id:"vintage", icon:"📅", label:"Vintages"},
  {id:"list",    icon:"📋", label:"Wine List"},
  {id:"charts",  icon:"📊", label:"Analytics"},
];

// ─── AUTH GATE ────────────────────────────────────────────────────────────────
function AuthGate({ onAuth }) {
  const [email,    setEmail]    = useState("");
  const [password, setPassword] = useState("");
  const [loading,  setLoading]  = useState(false);
  const [error,    setError]    = useState("");
  const [mode,     setMode]     = useState("login"); // login | signup

  const handleSubmit = async () => {
    setLoading(true); setError("");
    try {
      const { data, error: authError } = mode === "login"
        ? await supabase.auth.signInWithPassword({ email, password })
        : await supabase.auth.signUp({ email, password });
      if (authError) throw authError;
      if (data.session) onAuth(data.session);
      else if (mode === "signup") setError("Check your email to confirm your account, then log in.");
    } catch (e) {
      setError(e.message || "Authentication failed");
    }
    setLoading(false);
  };

  return (
    <div style={{ fontFamily:"'Palatino Linotype',serif", background:"#0a0604", minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
      <div style={{ background:"#1c0e04", border:"1px solid #2e1608", borderRadius:16, padding:40, maxWidth:400, width:"100%" }}>
        <div style={{ textAlign:"center", marginBottom:32 }}>
          <div style={{ fontSize:48, marginBottom:8 }}>🍷</div>
          <div style={{ fontSize:24, fontWeight:700, letterSpacing:3, color:"#c9a84c", textTransform:"uppercase" }}>CaveKeeper</div>
          <div style={{ fontSize:12, color:"#7a6848", letterSpacing:2, textTransform:"uppercase", marginTop:4 }}>Personal Wine Cellar</div>
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
          <div>
            <label style={{ fontSize:11, color:"#7a6848", display:"block", marginBottom:5, textTransform:"uppercase", letterSpacing:1 }}>Email</label>
            <input type="email" value={email} onChange={e=>setEmail(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleSubmit()} placeholder="household@email.com" style={{ background:"#0a0604", border:"1px solid #2e1608", color:"#e8dcc8", padding:"10px 12px", borderRadius:7, fontSize:14, fontFamily:"inherit", width:"100%", boxSizing:"border-box", outline:"none" }}/>
          </div>
          <div>
            <label style={{ fontSize:11, color:"#7a6848", display:"block", marginBottom:5, textTransform:"uppercase", letterSpacing:1 }}>Password</label>
            <input type="password" value={password} onChange={e=>setPassword(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleSubmit()} placeholder="••••••••" style={{ background:"#0a0604", border:"1px solid #2e1608", color:"#e8dcc8", padding:"10px 12px", borderRadius:7, fontSize:14, fontFamily:"inherit", width:"100%", boxSizing:"border-box", outline:"none" }}/>
          </div>
          {error && <div style={{ background:"#2d0a00", border:"1px solid #5a1a00", color:"#e07060", padding:"10px 12px", borderRadius:7, fontSize:13 }}>{error}</div>}
          <button onClick={handleSubmit} disabled={loading||!email||!password} style={{ background: loading||!email||!password ? "#2e1608" : "#8b1a1a", border:"none", color: loading||!email||!password ? "#7a6848" : "#fff", padding:"13px", borderRadius:8, cursor: loading||!email||!password ? "not-allowed" : "pointer", fontSize:15, fontFamily:"inherit", fontWeight:700, letterSpacing:1, marginTop:4 }}>
            {loading ? "..." : mode === "login" ? "Sign In" : "Create Account"}
          </button>
          <div style={{ textAlign:"center", fontSize:13, color:"#7a6848" }}>
            {mode === "login" ? (
              <span>First time? <span onClick={()=>{setMode("signup");setError("");}} style={{ color:"#c9a84c", cursor:"pointer" }}>Create account</span></span>
            ) : (
              <span>Have an account? <span onClick={()=>{setMode("login");setError("");}} style={{ color:"#c9a84c", cursor:"pointer" }}>Sign in</span></span>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── SEEDING HELPER — uploads SAMPLE_WINES to Supabase once ───────────────────
async function seedDatabase(wines) {
  const rows = wines.map(w => ({
    id:         w.id,
    winery:     w.winery,
    name:       w.name,
    varietal:   w.varietal,
    vintage:    w.vintage,
    region:     w.region,
    country:    w.country,
    row:        w.row,
    shelf:      w.shelf,
    slot:       w.slot,
    sticker:    w.sticker,
    price:      w.price,
    drink_from: w.drinkFrom,
    drink_to:   w.drinkTo,
    rating:     w.rating || 0,
    notes:      w.notes  || "",
    occasion:   w.occasion || [],
  }));
  // upsert in batches of 100 to avoid request size limits
  for (let i = 0; i < rows.length; i += 100) {
    const batch = rows.slice(i, i + 100);
    const { error } = await supabase.from("wines").upsert(batch, { onConflict: "id" });
    if (error) console.error("Seed error batch", i, error);
  }
}

// ─── DB → APP SHAPE ───────────────────────────────────────────────────────────
function dbToWine(r) {
  return {
    id:        r.id,
    winery:    r.winery    || "",
    name:      r.name      || "",
    varietal:  r.varietal  || "Other",
    vintage:   r.vintage   || 2020,
    region:    r.region    || "",
    country:   r.country   || "",
    row:       r.row       || "Row 1",
    shelf:     r.shelf     || 0,
    slot:      r.slot      || 0,
    sticker:   r.sticker   || "green",
    price:     r.price     || 0,
    drinkFrom: r.drink_from || 2022,
    drinkTo:   r.drink_to   || 2032,
    rating:    r.rating    || 0,
    notes:     r.notes     || "",
    occasion:  r.occasion  || [],
    quantity:  1,
  };
}

function wineToDb(w) {
  return {
    id:         w.id,
    winery:     w.winery,
    name:       w.name,
    varietal:   w.varietal,
    vintage:    w.vintage,
    region:     w.region,
    country:    w.country,
    row:        w.row,
    shelf:      w.shelf,
    slot:       w.slot,
    sticker:    w.sticker,
    price:      w.price,
    drink_from: w.drinkFrom,
    drink_to:   w.drinkTo,
    rating:     w.rating,
    notes:      w.notes,
    occasion:   w.occasion,
  };
}

// ═══════════════════════════════════════════════════════════════════════════════
function CellarMap({ wines, mapFilter, setMapFilter, slotMap, setSelected, setSlotPending, setView, setAddStep }) {
    const mapWines = mapFilter === "all" ? wines : mapFilter === "rtd" ? wines.filter(w=>w.sticker==="green"||w.sticker==="yellow"||w.sticker==="blue") : mapFilter === "needs" ? wines.filter(w=>w.sticker==="red") : wines.filter(w=>w.sticker==="black");
    const mapSlots = buildSlotMap(mapWines);

    // Reusable row column renderer
    const renderRow = (rowDef, mapSlots) => {
      const shelves = SHELF_STRUCTURE[rowDef.type];
      const used = wines.filter(w=>w.row===rowDef.id).length;
      const cap  = ROW_CAPACITY[rowDef.id];
      const pct  = Math.round(used/cap*100);
      return (
        <div key={rowDef.id} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:3}}>
          <div style={{fontSize:10,color:C.gold,fontWeight:700,letterSpacing:1,textAlign:"center",marginBottom:2}}>{rowDef.label}</div>
          <div style={{fontSize:9,color:pct>90?"#c0392b":pct>70?"#d4a017":C.muted,marginBottom:4}}>{used}/{cap}</div>
          {shelves.map((capacity,shelfIdx)=>{
            const slots = Array.from({length:capacity},(_,slotIdx)=>{
              const key = slotKey(rowDef.id,shelfIdx,slotIdx);
              const wine = mapSlots[key]||slotMap[key];
              return {key,wine,inFilter:!!mapSlots[key],slotIdx};
            });
            const shelfW = capacity*16+(capacity-1)*2+8;
            return (
              <div key={shelfIdx} style={{display:"flex",gap:2,alignItems:"center",background:"#1a0c04",border:`1px solid ${C.dim}`,borderRadius:4,padding:"3px 4px",width:shelfW,justifyContent:"center",boxSizing:"border-box"}}>
                {slots.map(({key,wine,inFilter,slotIdx})=>{
                  const isEmpty=!wine,dimmed=wine&&!inFilter;
                  return (
                    <div key={key}
                      title={wine?`${wine.name||wine.varietal}\n${wine.winery} ${wine.vintage}\n${STICKER[wine.sticker]?.label}`:`Empty — ${rowDef.id}, Shelf ${shelfIdx+1}, Slot ${slotIdx+1}`}
                      onClick={()=>{ if(wine){setSelected(wine);}else{setSlotPending({row:rowDef.id,shelf:shelfIdx,slot:slotIdx});setView("add");setAddStep(1);} }}
                      style={{width:14,height:14,borderRadius:3,background:isEmpty?C.dim:STICKER[wine.sticker]?.hex,border:isEmpty?`1px dashed ${C.border}`:"none",cursor:"pointer",opacity:dimmed?0.25:1,transition:"transform 0.1s",flexShrink:0}}
                      onMouseEnter={e=>e.currentTarget.style.transform="scale(1.35)"}
                      onMouseLeave={e=>e.currentTarget.style.transform="scale(1)"}
                    />
                  );
                })}
              </div>
            );
          })}
          <div style={{width:"100%",background:C.dim,borderRadius:2,height:3,marginTop:4,overflow:"hidden"}}>
            <div style={{width:`${pct}%`,height:"100%",background:pct>90?"#c0392b":pct>70?"#d4a017":"#1e8449",borderRadius:2}}/>
          </div>
        </div>
      );
    };

    return (
      <div>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:20,flexWrap:"wrap",gap:12}}>
          <div>
            <h2 style={{fontSize:22,color:C.gold,fontWeight:700,margin:0,letterSpacing:1}}>Cellar Map</h2>
            <p style={{color:C.muted,fontSize:13,margin:"4px 0 0"}}>Click any filled slot to inspect · Click any empty slot to assign a bottle</p>
          </div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {[["all","All Bottles"],["rtd","🟢🟡🔵 Ready/Available"],["needs","🔴 Needs Aging"],["past","⚫ Past Peak"]].map(([k,l])=>(
              <button key={k} onClick={()=>setMapFilter(k)} style={chip(mapFilter===k)}>{l}</button>
            ))}
          </div>
        </div>

        {/* Legend */}
        <div style={{display:"flex",gap:16,flexWrap:"wrap",marginBottom:20}}>
          {Object.entries(STICKER).map(([k,v])=>(
            <div key={k} style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:C.muted}}>
              <div style={{width:12,height:12,borderRadius:2,background:v.hex,flexShrink:0}}/>
              {v.label}
            </div>
          ))}
          <div style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:C.muted}}>
            <div style={{width:12,height:12,borderRadius:2,background:C.dim,border:`1px dashed ${C.border}`,flexShrink:0}}/>
            Empty
          </div>
        </div>

        {/* The actual map — scrollable horizontally */}
        <div style={{overflowX:"auto",paddingBottom:8}}>
          <div style={{minWidth:"max-content"}}>

            {/* Single flex row: R1 R2 R3 | [mid group: R4 R5 R6 + magnums inline] | R7 R8 R9 */}
            <div style={{display:"flex",gap:6,alignItems:"flex-start"}}>

              {/* R1, R2, R3 */}
              {ROW_DEFS.filter(r=>["Row 1","Row 2","Row 3"].includes(r.id)).map(rowDef => renderRow(rowDef, mapSlots))}

              {/* Mid group: R4/5/6 stacked beside magnums */}
              <div style={{display:"flex",flexDirection:"column",gap:0,alignItems:"center"}}>
                {/* R4, R5, R6 in a horizontal row */}
                <div style={{display:"flex",gap:6,alignItems:"flex-start"}}>
                  {ROW_DEFS.filter(r=>["Row 4","Row 5","Row 6"].includes(r.id)).map(rowDef => renderRow(rowDef, mapSlots))}
                </div>

                {/* Magnums centered below R4/5/6 */}
                <div style={{display:"flex",gap:6,alignItems:"flex-start",marginTop:8}}>
                  {/* Label above both */}
                  {ROW_DEFS.filter(r=>r.type==="magnum").map(rowDef=>{
                    const shelves = SHELF_STRUCTURE.magnum;
                    const used = wines.filter(w=>w.row===rowDef.id).length;
                    const cap  = ROW_CAPACITY[rowDef.id];
                    const pct  = Math.round(used/cap*100);
                    return (
                      <div key={rowDef.id} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:3}}>
                        <div style={{fontSize:10,color:C.gold,fontWeight:700,letterSpacing:1,textAlign:"center",marginBottom:2}}>{rowDef.label}</div>
                        <div style={{fontSize:9,color:pct>90?"#c0392b":pct>70?"#d4a017":C.muted,marginBottom:4}}>{used}/{cap}</div>
                        {shelves.map((_,shelfIdx)=>{
                          const key  = slotKey(rowDef.id,shelfIdx,0);
                          const wine = mapSlots[key]||slotMap[key];
                          const inFilter = !!mapSlots[key];
                          const isEmpty = !wine, dimmed = wine&&!inFilter;
                          return (
                            <div key={shelfIdx} style={{display:"flex",gap:2,background:"#1a0c04",border:`1px solid ${C.dim}`,borderRadius:4,padding:"3px 4px",boxSizing:"border-box"}}>
                              <div
                                title={wine?`${wine.name||wine.varietal}\n${wine.winery} ${wine.vintage}\nMagnum`:`Empty — ${rowDef.id}, Shelf ${shelfIdx+1}`}
                                onClick={()=>{ if(wine){setSelected(wine);}else{setSlotPending({row:rowDef.id,shelf:shelfIdx,slot:0});setView("add");setAddStep(1);} }}
                                style={{width:26,height:26,borderRadius:4,background:isEmpty?C.dim:STICKER[wine.sticker]?.hex,border:isEmpty?`1px dashed ${C.border}`:"none",cursor:"pointer",opacity:dimmed?0.25:1,transition:"transform 0.1s",flexShrink:0}}
                                onMouseEnter={e=>e.currentTarget.style.transform="scale(1.3)"}
                                onMouseLeave={e=>e.currentTarget.style.transform="scale(1)"}
                              />
                            </div>
                          );
                        })}
                        <div style={{width:"100%",background:C.dim,borderRadius:2,height:3,marginTop:4,overflow:"hidden"}}>
                          <div style={{width:`${pct}%`,height:"100%",background:pct>90?"#c0392b":pct>70?"#d4a017":"#1e8449",borderRadius:2}}/>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* R7, R8, R9 */}
              {ROW_DEFS.filter(r=>["Row 7","Row 8","Row 9"].includes(r.id)).map(rowDef => renderRow(rowDef, mapSlots))}

            </div>
          </div>
        </div>

        {/* Summary row — clickable → search filtered by row */}
        <div style={{marginTop:24}}>
          <div style={{fontSize:11,color:C.muted,textTransform:"uppercase",letterSpacing:1,marginBottom:10}}>Click any row to browse its bottles →</div>
          <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
            {ROW_DEFS.map(r=>{
              const used=wines.filter(w=>w.row===r.id).length;
              const cap=ROW_CAPACITY[r.id];
              const pct=Math.round(used/cap*100);
              const barColor = pct>90?"#a93226":pct>70?"#d4a017":"#1e8449";
              return (
                <div
                  key={r.id}
                  onClick={()=>{ setFilters(f=>({...f,row:[r.id]})); setView("search"); }}
                  title={`Browse all bottles in ${r.id}`}
                  style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:9,padding:"11px 16px",minWidth:86,textAlign:"center",cursor:"pointer",transition:"border-color 0.15s,transform 0.1s"}}
                  onMouseEnter={e=>{e.currentTarget.style.borderColor=C.gold;e.currentTarget.style.transform="translateY(-2px)";}}
                  onMouseLeave={e=>{e.currentTarget.style.borderColor=C.border;e.currentTarget.style.transform="none";}}
                >
                  <div style={{fontSize:12,color:C.gold,fontWeight:700,marginBottom:4}}>{r.id}</div>
                  <div style={{fontSize:20,color:C.text,fontWeight:700}}>{used}</div>
                  <div style={{fontSize:10,color:C.muted,margin:"3px 0 7px"}}>of {cap}</div>
                  <div style={{background:C.dim,borderRadius:3,height:4,overflow:"hidden"}}>
                    <div style={{width:`${pct}%`,height:"100%",background:barColor,borderRadius:3}}/>
                  </div>
                  <div style={{fontSize:10,color:barColor,marginTop:4}}>{pct}%</div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    );
  }

export default function CaveKeeper() {
  const [session,     setSession]     = useState(null);
  const [dbReady,     setDbReady]     = useState(false);
  const [wines,       setWines]       = useState([]);
  const [dbLoading,   setDbLoading]   = useState(true);
  const [view,        setView]        = useState("map");
  const [selected,    setSelected]    = useState(null);   // wine object
  const [slotPending, setSlotPending] = useState(null);   // {row,shelf,slot} for assign
  const [search,      setSearch]      = useState("");
  const [filters,     setFilters]     = useState({ varietal:[], region:[], sticker:[], row:[], occasion:[], rtd:false });
  const [sort,        setSort]        = useState({ key:"winery", dir:"asc" });
  const [addForm,     setAddForm]     = useState({ winery:"", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", row:"Row 1", shelf:0, slot:0, price:"", drinkFrom:2024, drinkTo:2034, rating:"", notes:"", occasion:[] });
  const [addStep,     setAddStep]     = useState(1);
  const [toast,       setToast]       = useState(null);
  const [aiLoading,   setAiLoading]   = useState(false);
  const [aiRec,       setAiRec]       = useState("");
  const [aiOccasion,  setAiOccasion]  = useState("Celebration");
  const [aiGuests,    setAiGuests]    = useState(4);
  const [mapFilter,    setMapFilter]    = useState("all");
  const [notifOpen,    setNotifOpen]    = useState(false);
  const [dismissedIds, setDismissedIds] = useState([]);
  const [editing,      setEditing]      = useState(false);  // wine detail edit mode
  const [editForm,     setEditForm]     = useState(null);   // copy of wine being edited
  const [labelScanning, setLabelScanning] = useState(false);  // scanning label photo
  const [labelPreview,  setLabelPreview]  = useState(null);   // base64 preview
  const [labelResult,   setLabelResult]   = useState(null);   // AI extracted data
  const [fetchingNotes, setFetchingNotes] = useState(false);  // fetching tasting notes
  const fileRef  = useRef();
  const labelRef = useRef();

  // ── Auth ──────────────────────────────────────────────────────────────────
  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session);
    });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_e, s) => {
      setSession(s);
    });
    return () => subscription.unsubscribe();
  }, []);

  // ── Load wines from Supabase ──────────────────────────────────────────────
  useEffect(() => {
    if (!session) return;
    const loadWines = async () => {
      setDbLoading(true);
      const { data, error } = await supabase.from("wines").select("*").order("row").order("shelf").order("slot");
      if (error) {
        console.error("Load error:", error);
        setDbLoading(false);
        return;
      }
      if (data && data.length > 0) {
        setWines(data.map(dbToWine));
        setDbReady(true);
      } else {
        // First run — seed the database from SAMPLE_WINES
        await seedDatabase(SAMPLE_WINES);
        setWines(SAMPLE_WINES);
        setDbReady(true);
      }
      setDbLoading(false);
    };
    loadWines();

    // Real-time subscription — any change on any device updates all devices
    const channel = supabase
      .channel("wines-realtime")
      .on("postgres_changes", { event: "*", schema: "public", table: "wines" }, (payload) => {
        if (payload.eventType === "INSERT") {
          setWines(prev => [...prev, dbToWine(payload.new)]);
        } else if (payload.eventType === "UPDATE") {
          setWines(prev => prev.map(w => w.id === payload.new.id ? dbToWine(payload.new) : w));
        } else if (payload.eventType === "DELETE") {
          setWines(prev => prev.filter(w => w.id !== payload.old.id));
        }
      })
      .subscribe();

    return () => supabase.removeChannel(channel);
  }, [session]);

  // ── Load dismissed alerts from Supabase ───────────────────────────────────
  useEffect(() => {
    if (!session) return;
    supabase.from("dismissed_alerts").select("alert_id").then(({ data }) => {
      if (data) setDismissedIds(data.map(r => r.alert_id));
    });
  }, [session]);

  const slotMap = useMemo(() => buildSlotMap(wines), [wines]);

  const allAlerts = useMemo(() => {
    const alerts = [];
    const nextMonth = currentYear + (new Date().getMonth() === 11 ? 1 : 0);
    const nextMonthNum = (new Date().getMonth() + 1) % 12; // 0-indexed

    wines.forEach(w => {
      const id_ready   = `ready-${w.id}`;
      const id_sticker = `sticker-${w.id}`;
      const id_peak    = `peak-${w.id}`;

      // 1. Becoming ready to drink within ~1 month (drinkFrom === currentYear and month is close)
      if (w.drinkFrom === currentYear && !isRTD(w) === false) {
        // already ready — check sticker mismatch instead
      }
      if (w.drinkFrom === currentYear + 1) {
        alerts.push({
          id: id_ready,
          type: "ready",
          icon: "🟢",
          title: `${w.winery} ${w.vintage} becomes ready next year`,
          detail: `${w.varietal} · ${w.row} Shelf ${w.shelf+1} · Drink window opens ${w.drinkFrom}`,
          action: "Update sticker from Red → " + (w.price >= 100 ? "Blue ($100+)" : w.price >= 50 ? "Yellow ($50–$100)" : "Green (Under $50)"),
          wine: w,
          severity: "info",
        });
      }

      // 2. Sticker mismatch — wine's sticker doesn't match what it should be now
      const correct = autoSticker(w.price, w.drinkFrom, w.drinkTo);
      if (correct !== w.sticker) {
        alerts.push({
          id: id_sticker,
          type: "sticker",
          icon: "🏷️",
          title: `${w.winery} ${w.vintage} — sticker needs updating`,
          detail: `${w.varietal} · ${w.row} Shelf ${w.shelf+1}`,
          action: `Change sticker: ${STICKER[w.sticker]?.label} → ${STICKER[correct]?.label}`,
          wine: w,
          correct,
          severity: correct === "black" ? "critical" : "warning",
        });
      }

      // 3. Approaching peak — within 1 year of drinkTo
      if (w.drinkTo >= currentYear && w.drinkTo <= currentYear + 1 && isRTD(w)) {
        alerts.push({
          id: id_peak,
          type: "peak",
          icon: "⚠️",
          title: `${w.winery} ${w.vintage} — drink soon, approaching peak`,
          detail: `${w.varietal} · ${w.row} Shelf ${w.shelf+1} · Peak: ${w.drinkTo}`,
          action: "Consider opening within the next 12 months",
          wine: w,
          severity: "warning",
        });
      }
    });

    return alerts;
  },[wines]);

  const filteredWines = useMemo(() => {
    // Only show bottles that have a winery name (filter out empty slots)
    let w = wines.filter(x => x.winery && x.winery.trim() !== "");
    if (search) { const q=search.toLowerCase(); w=w.filter(x=>x.winery.toLowerCase().includes(q)||x.varietal.toLowerCase().includes(q)||x.region.toLowerCase().includes(q)||String(x.vintage).includes(q)); }
    if (filters.varietal.length) w=w.filter(x=>filters.varietal.includes(x.varietal));
    if (filters.region.length)   w=w.filter(x=>filters.region.includes(x.region));
    if (filters.sticker.length)  w=w.filter(x=>filters.sticker.includes(x.sticker));
    if (filters.row.length)      w=w.filter(x=>filters.row.includes(x.row));
    if (filters.occasion.length) w=w.filter(x=>x.occasion&&filters.occasion.some(o=>x.occasion.includes(o)));
    if (filters.rtd)             w=w.filter(isRTD);
    w.sort((a,b)=>{ let av=a[sort.key],bv=b[sort.key]; if(typeof av==="string"){av=av.toLowerCase();bv=bv.toLowerCase();} return sort.dir==="asc"?(av>bv?1:-1):(av<bv?1:-1); });
    return w;
  },[wines, search, filters, sort]);

  const varietalData = useMemo(()=>{
    const m={}; wines.forEach(w=>{m[w.varietal]=(m[w.varietal]||0)+1;});
    return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,8).map(([name,value])=>({name,value}));
  },[wines]);
  const stickerData = useMemo(()=>Object.entries(STICKER).map(([k,v])=>({
    name:v.label, value:wines.filter(w=>w.sticker===k).length, color:v.hex
  })).filter(x=>x.value>0),[wines]);
  const vintageBarData = useMemo(()=>{
    const m={}; wines.forEach(w=>{m[w.vintage]=(m[w.vintage]||0)+1;});
    return Object.entries(m).sort((a,b)=>a[0]-b[0]).map(([v,q])=>({vintage:String(v),qty:q}));
  },[wines]);

  // Show auth gate if not logged in
  if (!session) return <AuthGate onAuth={setSession} />;

  // Show loading screen while fetching from Supabase
  if (dbLoading) return (
    <div style={{ fontFamily:"'Palatino Linotype',serif", background:"#0a0604", minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", flexDirection:"column", gap:16 }}>
      <div style={{ fontSize:48 }}>🍷</div>
      <div style={{ color:"#c9a84c", fontSize:18, fontWeight:700, letterSpacing:2 }}>Loading your cellar…</div>
      <div style={{ color:"#7a6848", fontSize:13 }}>Connecting to Supabase</div>
    </div>
  );

  const activeAlerts = allAlerts.filter(a => !dismissedIds.includes(a.id));
  const criticalCount = activeAlerts.filter(a => a.severity === "critical").length;
  const warningCount  = activeAlerts.filter(a => a.severity === "warning").length;
  const infoCount     = activeAlerts.filter(a => a.severity === "info").length;

  const dismissAlert = async (id) => {
    setDismissedIds(prev => [...prev, id]);
    await supabase.from("dismissed_alerts").upsert({ alert_id: id }, { onConflict: "alert_id" });
  };
  const dismissAll = async () => {
    const ids = allAlerts.map(a => a.id);
    setDismissedIds(ids);
    await supabase.from("dismissed_alerts").upsert(ids.map(id => ({ alert_id: id })), { onConflict: "alert_id" });
  };

  const updateSticker = async (wine, correct) => {
    await supabase.from("wines").update({ sticker: correct }).eq("id", String(wine.id));
    dismissAlert(`sticker-${wine.id}`);
    showToast(`Sticker updated to ${STICKER[correct]?.label} for ${wine.winery} ${wine.vintage}`);
  };

  const severityColor = (s) => s === "critical" ? "#a93226" : s === "warning" ? "#d4a017" : "#2471a3";

  // ── Label photo scan ──────────────────────────────────────────────────────
  const handleLabelPhoto = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (ev) => {
      const base64 = ev.target.result; // full data URL
      const base64Data = base64.split(",")[1];
      setLabelPreview(base64);
      setLabelScanning(true);
      setLabelResult(null);
      try {
        const res = await fetch(CLAUDE_URL, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            model: "claude-sonnet-4-20250514",
            max_tokens: 1000,
            messages: [{
              role: "user",
              content: [
                { type: "image", source: { type: "base64", media_type: file.type || "image/jpeg", data: base64Data } },
                { type: "text", text: `You are a master sommelier reading a wine label. Extract ALL information visible and return ONLY a JSON object with these exact keys (use null if not visible):
{
  "winery": "producer/winery name",
  "vintage": 2019,
  "varietal": "grape variety or blend name",
  "region": "appellation or region",
  "country": "country of origin",
  "alcoholPct": 14.5,
  "notes": "any text on label like reserve, estate, vineyard name, winemaker notes",
  "drinkFrom": 2022,
  "drinkTo": 2035
}
Return ONLY the JSON, no other text.` }
              ]
            }]
          })
        });
        const data = await res.json();
        const raw = data.content?.[0]?.text || "{}";
        const clean = raw.replace(/```json|```/g, "").trim();
        const parsed = JSON.parse(clean);
        setLabelResult(parsed);
        // Pre-fill the form with extracted data
        setAddForm(f => ({
          ...f,
          winery:    parsed.winery   || f.winery,
          vintage:   parsed.vintage  || f.vintage,
          varietal:  VARIETALS.find(v => v.toLowerCase().includes((parsed.varietal||"").toLowerCase().split(" ")[0])) || f.varietal,
          region:    REGIONS.find(r => r.toLowerCase().includes((parsed.region||"").toLowerCase().split(" ")[0]))    || f.region,
          drinkFrom: parsed.drinkFrom || f.drinkFrom,
          drinkTo:   parsed.drinkTo   || f.drinkTo,
          notes:     parsed.notes     || f.notes,
        }));
      } catch(err) {
        setLabelResult({ error: "Could not read label. Please fill in manually." });
      }
      setLabelScanning(false);
    };
    reader.readAsDataURL(file);
    e.target.value = "";
  };

  const acceptLabelScan = () => {
    setLabelPreview(null);
    setLabelResult(null);
  };

  const retryLabelScan = () => {
    setLabelPreview(null);
    setLabelResult(null);
    labelRef.current.click();
  };

  // ── Auto tasting notes (for existing wines in detail modal) ───────────────
  const fetchTastingNotes = async (wine) => {
    setFetchingNotes(true);
    try {
      const res = await fetch(CLAUDE_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 200,
          messages: [{
            role: "user",
            content: `You are a master sommelier. Write a 2-sentence tasting note for: ${wine.winery} ${wine.vintage} ${wine.name ? '"'+wine.name+'"' : ""} ${wine.varietal} from ${wine.region}.

Describe only the sensory experience: color, aroma, palate flavors, texture, and finish. Be specific and expert. Start directly with the tasting description — no preamble.`
          }]
        })
      });
      const data = await res.json();
      const note = data.content?.[0]?.text || "No tasting notes found.";
      await supabase.from("wines").update({ notes: note }).eq("id", String(wine.id));
      setSelected(prev => prev ? { ...prev, notes: note } : prev);
      showToast("Tasting notes updated!");
    } catch(e) {
      showToast("Could not fetch tasting notes.", "error");
    }
    setFetchingNotes(false);
  };

  const showToast = (msg, type="success") => {
    setToast({msg,type});
    setTimeout(() => setToast(null), 3200);
  };

  // ── Stats ──
  const totalBottles = wines.filter(w=>w.winery && w.winery.trim()!=="").length;
  const totalValue   = wines.reduce((s,w) => s + w.price, 0);
  const rtdCount     = wines.filter(isRTD).length;

  const toggleFilter = (key, val) => setFilters(f => ({ ...f, [key]: f[key].includes(val) ? f[key].filter(x=>x!==val) : [...f[key],val] }));

  // ── Excel import ──
  const handleImport = e => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = async ev => {
      const wb = XLSX.read(ev.target.result,{type:"binary"});
      const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      const imported = data.map((r,i) => ({
        id: Date.now()+i, winery:r.Winery||"", name:r.Name||r.Winery||"",
        varietal:r.Varietal||"Other", vintage:Number(r.Vintage||2020),
        region:r.Region||"Other", country:r.Country||"",
        row:r.Row||"Row 1", shelf:Number(r.Shelf||0)-1, slot:Number(r.Slot||0)-1,
        sticker:r.Sticker||"green", price:Number(r.Price||0),
        drinkFrom:Number(r.DrinkFrom||currentYear), drinkTo:Number(r.DrinkTo||currentYear+5),
        quantity:1, rating:Number(r.Rating||0), notes:r.Notes||"", occasion:[],
      }));
      const { error } = await supabase.from("wines").upsert(imported.map(wineToDb), { onConflict:"id" });
      if (error) { showToast("Import failed", "error"); return; }
      showToast(`Imported ${imported.length} bottles`);
    };
    reader.readAsBinaryString(file);
    e.target.value="";
  };

  // ── Excel export ──
  const handleExport = () => {
    const data = wines.map(w=>({ Winery:w.winery, Varietal:w.varietal, Vintage:w.vintage, Region:w.region, Row:w.row, Shelf:w.shelf+1, Slot:w.slot+1, Sticker:w.sticker, Price:w.price, DrinkFrom:w.drinkFrom, DrinkTo:w.drinkTo, Rating:w.rating, Notes:w.notes }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"CaveKeeper");
    XLSX.writeFile(wb,"cavekeeper-inventory.xlsx");
    showToast("Exported to Excel");
  };

  // ── Add wine ──
  const nextFreeSlot = (row) => {
    const def = ROW_DEFS.find(r=>r.id===row);
    const shelves = SHELF_STRUCTURE[def.type];
    for (let s=0; s<shelves.length; s++) {
      for (let sl=0; sl<shelves[s]; sl++) {
        if (!slotMap[slotKey(row,s,sl)]) return {shelf:s,slot:sl};
      }
    }
    return null;
  };

  const submitAdd = async () => {
    const sticker = autoSticker(Number(addForm.price), Number(addForm.drinkFrom), Number(addForm.drinkTo));
    // Use user-selected shelf/slot, fall back to next free if not set
    const shelf = Number(addForm.shelf) || 0;
    const slot  = Number(addForm.slot)  || 0;
    // Check if selected slot is already taken
    if (slotMap[slotKey(addForm.row, shelf, slot)]) {
      showToast("That slot is already occupied — pick another", "error"); return;
    }
    const w = { ...addForm, id:Date.now(), shelf, slot, sticker, price:Number(addForm.price), rating:Number(addForm.rating), vintage:Number(addForm.vintage), drinkFrom:Number(addForm.drinkFrom), drinkTo:Number(addForm.drinkTo), quantity:1 };
    const { error } = await supabase.from("wines").insert(wineToDb(w));
    if (error) { showToast(`Error: ${error.message}`, "error"); return; }
    showToast(`${w.name||w.winery} ${w.vintage} added to ${w.row}, shelf ${w.shelf+1}`);
    setAddForm({ winery:"", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", row:"Row 1", shelf:0, slot:0, price:"", drinkFrom:2024, drinkTo:2034, rating:"", notes:"", occasion:[] });
    setAddStep(1); setView("map");
  };

  // Assign pending slot from map click
  const assignToSlot = async () => {
    if (!slotPending) return;
    const sticker = autoSticker(Number(addForm.price), Number(addForm.drinkFrom), Number(addForm.drinkTo));
    const w = { ...addForm, id:Date.now(), row:slotPending.row, shelf:slotPending.shelf, slot:slotPending.slot, sticker, price:Number(addForm.price), rating:Number(addForm.rating), vintage:Number(addForm.vintage), drinkFrom:Number(addForm.drinkFrom), drinkTo:Number(addForm.drinkTo), quantity:1 };
    const { error } = await supabase.from("wines").insert(wineToDb(w));
    if (error) { showToast(`Error: ${error.message}`, "error"); return; }
    showToast(`${w.name||w.winery} ${w.vintage} placed in ${w.row}, shelf ${w.shelf+1}`);
    setSlotPending(null); setAddForm({ winery:"", varietal:"Cabernet Sauvignon", vintage:2022, region:"Napa Valley", row:"Row 1", shelf:0, slot:0, price:"", drinkFrom:2024, drinkTo:2034, rating:"", notes:"", occasion:[] });
  };

  const getAiRec = async () => {
    setAiLoading(true); setAiRec("");
    const rtd = wines.filter(w=>w.sticker==="green"||w.sticker==="yellow"||w.sticker==="blue").slice(0,30).map(w=>
      `${w.name||w.varietal} · ${w.winery} ${w.vintage} · ${w.varietal} · ${w.region} · $${w.price} · LOCATION: ${w.row}, Shelf ${w.shelf+1}, Slot ${w.slot+1}`
    ).join("\n");
    try {
      const res = await fetch(CLAUDE_URL,{
        method:"POST", headers:{"Content-Type":"application/json"},
        body:JSON.stringify({ model:"claude-sonnet-4-20250514", max_tokens:600, messages:[{role:"user",content:`You are a master sommelier. Occasion: ${aiOccasion} for ${aiGuests} guests.\n\nMy available wines (with cellar location):\n${rtd}\n\nRecommend 2-3 specific bottles. For each:\n1. Wine name and vintage\n2. 📍 Cellar location (exactly as listed)\n3. Why it fits this occasion\n4. Serving temperature\n\nBe concise and expert.`}] })
      });
      const d = await res.json();
      setAiRec(d.content?.[0]?.text || "No recommendation.");
    } catch { setAiRec("AI unavailable — check connection."); }
    setAiLoading(false);
  };

  const CHART_COLORS = ["#8b1a1a","#2471a3","#1e8449","#d4a017","#7d6608","#6c3483","#117a65","#784212"];

  // ═══════════════════════════════════════════════════════════════════════════
  // CELLAR MAP COMPONENT (the centrepiece)
  // ═══════════════════════════════════════════════════════════════════════════
;

  // ═══════════════════════════════════════════════════════════════════════════
  return (
    <div style={{fontFamily:"'Palatino Linotype','Book Antiqua',Palatino,serif",background:C.bg,minHeight:"100vh",color:C.text}}>

      {/* ── HEADER ── */}
      <header style={{background:"linear-gradient(135deg,#120600,#1e0d03,#120600)",borderBottom:`1px solid ${C.border}`,padding:"0 20px",display:"flex",alignItems:"center",justifyContent:"space-between",height:60,position:"sticky",top:0,zIndex:100,gap:12}}>
        <div style={{display:"flex",alignItems:"center",gap:10,flexShrink:0}}>
          <span style={{fontSize:24}}>🍷</span>
          <div>
            <div style={{fontSize:17,fontWeight:700,letterSpacing:3,color:C.gold,textTransform:"uppercase",lineHeight:1}}>CaveKeeper</div>
            <div style={{fontSize:9,color:C.goldDim,letterSpacing:2,textTransform:"uppercase"}}>Personal Wine Cellar</div>
          </div>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          {[["Bottles",totalBottles,C.text],["Capacity",`${Math.round(totalBottles/TOTAL_CAPACITY*100)}%`,C.text],["Ready",rtdCount,"#1e8449"],["Value",`$${(totalValue/1000).toFixed(0)}k`,C.gold]].map(([l,v,c])=>(
            <div key={l} style={{background:"#1c0e04",border:`1px solid ${C.border}`,borderRadius:7,padding:"3px 10px",textAlign:"center"}}>
              <div style={{fontSize:15,fontWeight:700,color:c}}>{v}</div>
              <div style={{fontSize:9,color:C.goldDim,textTransform:"uppercase",letterSpacing:1}}>{l}</div>
            </div>
          ))}
          {/* Notification Bell */}
          <div style={{position:"relative"}}>
            <button
              onClick={()=>setNotifOpen(o=>!o)}
              title={`${activeAlerts.length} alert${activeAlerts.length!==1?"s":""}`}
              style={{background: activeAlerts.length ? "#2d0e00" : "#1c0e04", border:`1px solid ${activeAlerts.length ? "#8b1a1a" : C.border}`,borderRadius:7,padding:"3px 12px",cursor:"pointer",fontSize:20,lineHeight:1,display:"flex",alignItems:"center",justifyContent:"center",height:42,position:"relative",transition:"border-color 0.2s"}}
            >
              🔔
              {activeAlerts.length > 0 && (
                <span style={{position:"absolute",top:-5,right:-5,background: criticalCount ? "#a93226" : "#d4a017",color:"#fff",borderRadius:"50%",width:18,height:18,fontSize:10,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"inherit"}}>
                  {activeAlerts.length}
                </span>
              )}
            </button>
          </div>
        </div>
      </header>

      {/* ── NAV ── */}
      <nav style={{background:C.surface,borderBottom:`1px solid ${C.border}`,display:"flex",overflowX:"auto",padding:"0 4px",gap:0}}>
        {NAV.map(n=>(
          <button key={n.id} onClick={()=>setView(n.id)} style={{background:"none",border:"none",color:view===n.id?C.gold:C.muted,padding:"10px 14px",cursor:"pointer",fontSize:11,fontFamily:"inherit",letterSpacing:1,textTransform:"uppercase",borderBottom:view===n.id?`2px solid ${C.gold}`:"2px solid transparent",whiteSpace:"nowrap",display:"flex",flexDirection:"column",alignItems:"center",gap:3,transition:"color 0.15s",flexShrink:0}}>
            <span style={{fontSize:16}}>{n.icon}</span>{n.label}
          </button>
        ))}
        <div style={{marginLeft:"auto",display:"flex",gap:6,alignItems:"center",padding:"0 8px",flexShrink:0}}>
          <button onClick={()=>fileRef.current.click()} style={{background:C.card,border:`1px solid ${C.border}`,color:C.gold,padding:"5px 12px",borderRadius:6,cursor:"pointer",fontSize:11,fontFamily:"inherit",whiteSpace:"nowrap"}}>📥 Import</button>
          <button onClick={handleExport} style={{background:C.card,border:`1px solid ${C.border}`,color:C.gold,padding:"5px 12px",borderRadius:6,cursor:"pointer",fontSize:11,fontFamily:"inherit",whiteSpace:"nowrap"}}>📤 Export</button>
          <button onClick={() => supabase.auth.signOut()} style={{background:C.card,border:`1px solid ${C.border}`,color:C.muted,padding:"5px 12px",borderRadius:6,cursor:"pointer",fontSize:11,fontFamily:"inherit",whiteSpace:"nowrap"}}>Sign Out</button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleImport} style={{display:"none"}}/>
        </div>
      </nav>

      {/* ── TOAST ── */}
      {toast && (
        <div style={{position:"fixed",bottom:24,right:24,background:toast.type==="success"?"#1e8449":"#a93226",color:"#fff",padding:"12px 20px",borderRadius:8,zIndex:999,boxShadow:"0 4px 24px rgba(0,0,0,.6)",fontSize:13,maxWidth:320}}>
          {toast.msg}
        </div>
      )}

      {/* ── NOTIFICATION PANEL ── */}
      {notifOpen && (
        <>
          {/* Backdrop */}
          <div onClick={()=>setNotifOpen(false)} style={{position:"fixed",inset:0,zIndex:149}}/>
          {/* Drawer */}
          <div style={{position:"fixed",top:62,right:16,width:400,maxWidth:"calc(100vw - 32px)",maxHeight:"80vh",background:C.card,border:`1px solid ${C.border}`,borderRadius:12,zIndex:150,boxShadow:"0 12px 48px rgba(0,0,0,.85)",display:"flex",flexDirection:"column",overflow:"hidden"}}>

            {/* Panel header */}
            <div style={{padding:"16px 18px",borderBottom:`1px solid ${C.border}`,display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
              <div>
                <div style={{fontSize:15,color:C.gold,fontWeight:700}}>🔔 Monthly Digest</div>
                <div style={{fontSize:11,color:C.muted,marginTop:2}}>
                  {activeAlerts.length === 0
                    ? "All clear — no action needed"
                    : `${activeAlerts.length} alert${activeAlerts.length!==1?"s":""} · ${criticalCount} critical · ${warningCount} warnings · ${infoCount} info`
                  }
                </div>
              </div>
              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                {activeAlerts.length > 0 && (
                  <button onClick={dismissAll} style={{background:"none",border:`1px solid ${C.border}`,color:C.muted,padding:"4px 10px",borderRadius:6,cursor:"pointer",fontSize:11,fontFamily:"inherit"}}>
                    Dismiss all
                  </button>
                )}
                <button onClick={()=>setNotifOpen(false)} style={{background:"none",border:"none",color:C.muted,fontSize:18,cursor:"pointer",lineHeight:1,padding:4}}>✕</button>
              </div>
            </div>

            {/* Alert list */}
            <div style={{overflowY:"auto",flex:1}}>
              {activeAlerts.length === 0 ? (
                <div style={{padding:40,textAlign:"center"}}>
                  <div style={{fontSize:32,marginBottom:12}}>✅</div>
                  <div style={{color:C.muted,fontSize:14}}>Your cellar is looking great.</div>
                  <div style={{color:C.goldDim,fontSize:12,marginTop:6}}>No sticker changes or drink window alerts this month.</div>
                </div>
              ) : (
                <>
                  {/* Group by severity */}
                  {["critical","warning","info"].map(sev => {
                    const group = activeAlerts.filter(a => a.severity === sev);
                    if (!group.length) return null;
                    const sevLabel = sev === "critical" ? "🚨 Needs Immediate Attention" : sev === "warning" ? "⚠️ Drink Window Alerts" : "ℹ️ Coming Up";
                    return (
                      <div key={sev}>
                        <div style={{padding:"10px 18px 6px",fontSize:10,color:severityColor(sev),textTransform:"uppercase",letterSpacing:1.5,fontWeight:700,background:C.surface,borderBottom:`1px solid ${C.border}`}}>
                          {sevLabel}
                        </div>
                        {group.map(alert => (
                          <div key={alert.id} style={{padding:"14px 18px",borderBottom:`1px solid ${C.surface}`,background:C.card}}>
                            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
                              <div style={{flex:1}}>
                                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
                                  <span style={{fontSize:16}}>{alert.icon}</span>
                                  <span style={{fontSize:13,color:C.text,fontWeight:600,lineHeight:1.3}}>{alert.title}</span>
                                </div>
                                <div style={{fontSize:12,color:C.muted,marginBottom:6,paddingLeft:24}}>{alert.detail}</div>
                                <div style={{fontSize:12,color:severityColor(alert.severity),paddingLeft:24,fontStyle:"italic"}}>{alert.action}</div>
                                {/* One-tap fix for sticker alerts */}
                                {alert.type === "sticker" && (
                                  <button
                                    onClick={()=>updateSticker(alert.wine, alert.correct)}
                                    style={{marginTop:10,marginLeft:24,background:`${severityColor(alert.severity)}22`,border:`1px solid ${severityColor(alert.severity)}`,color:severityColor(alert.severity),padding:"5px 12px",borderRadius:6,cursor:"pointer",fontSize:12,fontFamily:"inherit",fontWeight:600}}
                                  >
                                    ● Fix Sticker Now
                                  </button>
                                )}
                                {/* Navigate to bottle */}
                                <button
                                  onClick={()=>{ setSelected(alert.wine); setNotifOpen(false); }}
                                  style={{marginTop:alert.type==="sticker"?6:10,marginLeft:24,background:"none",border:`1px solid ${C.border}`,color:C.muted,padding:"4px 10px",borderRadius:6,cursor:"pointer",fontSize:11,fontFamily:"inherit"}}
                                >
                                  View Bottle →
                                </button>
                              </div>
                              <button onClick={()=>dismissAlert(alert.id)} style={{background:"none",border:"none",color:C.muted,fontSize:16,cursor:"pointer",flexShrink:0,padding:"0 2px",lineHeight:1}} title="Dismiss">✕</button>
                            </div>
                          </div>
                        ))}
                      </div>
                    );
                  })}
                </>
              )}
            </div>

            {/* Panel footer — digest summary */}
            <div style={{padding:"12px 18px",borderTop:`1px solid ${C.border}`,flexShrink:0,background:C.surface}}>
              <div style={{fontSize:11,color:C.muted,textAlign:"center"}}>
                CaveKeeper checks your cellar monthly · {new Date().toLocaleString("default",{month:"long",year:"numeric"})} digest
              </div>
            </div>
          </div>
        </>
      )}

      <main style={{padding:"20px",maxWidth:1440,margin:"0 auto"}}>

        {/* ════════════════════ CELLAR MAP ════════════════════ */}
        {view==="map" && <CellarMap wines={wines} mapFilter={mapFilter} setMapFilter={setMapFilter} slotMap={slotMap} setSelected={setSelected} setSlotPending={setSlotPending} setView={setView} setAddStep={setAddStep}/>}

        {/* ════════════════════ MY CELLAR ════════════════════ */}
        {view==="cellar" && (
          <div>
            <SectionTitle>My Cellar</SectionTitle>
            <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:16}}>
              <button onClick={()=>setFilters(f=>({...f,rtd:!f.rtd}))} style={chip(filters.rtd,"#1e8449")}>✓ Ready to Drink</button>
              {Object.entries(STICKER).map(([k,v])=>(
                <button key={k} onClick={()=>toggleFilter("sticker",k)} style={{...chip(filters.sticker.includes(k)),background:filters.sticker.includes(k)?`${v.hex}33`:C.card,color:filters.sticker.includes(k)?v.hex:C.muted,border:`1px solid ${filters.sticker.includes(k)?v.hex:C.border}`}}>
                  ● {v.label}
                </button>
              ))}
              {(filters.sticker.length||filters.rtd)&&<button onClick={()=>setFilters({varietal:[],region:[],sticker:[],row:[],occasion:[],rtd:false})} style={{...chip(false),color:"#a93226",borderColor:"#a93226"}}>✕ Clear</button>}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(270px,1fr))",gap:14}}>
              {filteredWines.map(w=><WineCard key={w.id} w={w} onClick={()=>setSelected(w)}/>)}
            </div>
            {!filteredWines.length&&<Empty msg="No wines match filters."/>}
          </div>
        )}

        {/* ════════════════════ ADD WINE ════════════════════ */}
        {view==="add" && (
          <div style={{maxWidth:600,margin:"0 auto"}}>
            <SectionTitle>{slotPending ? `Assign Wine → ${slotPending.row}, Shelf ${slotPending.shelf+1}, Slot ${slotPending.slot+1}` : "Add Wine to Cellar"}</SectionTitle>
            {slotPending && (
              <div style={{background:"#1e2d04",border:`1px solid #3a5a08`,borderRadius:8,padding:12,marginBottom:16,fontSize:13,color:"#7dbb30"}}>
                📍 You selected an empty slot from the Cellar Map. Fill in the details below and the bottle will be placed there precisely.
              </div>
            )}
            <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:24}}>
              {/* Steps */}
              <div style={{display:"flex",gap:8,marginBottom:24}}>
                {["Wine Details","Drink Window","Confirm"].map((s,i)=>(
                  <div key={i} style={{flex:1,textAlign:"center"}}>
                    <div style={{width:26,height:26,borderRadius:"50%",background:addStep>i+1?"#1e8449":addStep===i+1?C.gold:C.dim,color:addStep>=i+1?"#fff":C.muted,display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 5px",fontSize:12,fontWeight:700}}>{addStep>i+1?"✓":i+1}</div>
                    <div style={{fontSize:10,color:addStep===i+1?C.gold:C.muted,textTransform:"uppercase",letterSpacing:1}}>{s}</div>
                  </div>
                ))}
              </div>

              {addStep===1&&(
                <div style={{display:"flex",flexDirection:"column",gap:14}}>

                  {/* ── Label Photo Scan ── */}
                  <input ref={labelRef} type="file" accept="image/*" capture="environment" onChange={handleLabelPhoto} style={{display:"none"}}/>

                  {!labelPreview && (
                    <div
                      onClick={()=>labelRef.current.click()}
                      style={{border:`2px dashed ${C.border}`,borderRadius:10,padding:"20px 16px",textAlign:"center",cursor:"pointer",background:C.surface,transition:"border-color 0.2s"}}
                      onMouseEnter={e=>e.currentTarget.style.borderColor=C.gold}
                      onMouseLeave={e=>e.currentTarget.style.borderColor=C.border}
                    >
                      <div style={{fontSize:32,marginBottom:8}}>📷</div>
                      <div style={{fontSize:14,color:C.gold,fontWeight:700,marginBottom:4}}>Scan Wine Label</div>
                      <div style={{fontSize:12,color:C.muted}}>Take a photo or upload an image — AI will extract all details automatically</div>
                    </div>
                  )}

                  {/* Label preview + scanning state */}
                  {labelPreview && (
                    <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:10,overflow:"hidden"}}>
                      <div style={{display:"flex",gap:12,padding:12,alignItems:"flex-start"}}>
                        <img src={labelPreview} alt="Wine label" style={{width:90,height:120,objectFit:"cover",borderRadius:6,flexShrink:0,border:`1px solid ${C.border}`}}/>
                        <div style={{flex:1}}>
                          {labelScanning && (
                            <div style={{paddingTop:16}}>
                              <div style={{fontSize:13,color:C.gold,fontWeight:700,marginBottom:8}}>🤖 Reading label…</div>
                              <div style={{background:C.dim,borderRadius:4,height:4,overflow:"hidden"}}>
                                <div style={{height:"100%",background:C.gold,borderRadius:4,animation:"scan 1.5s ease-in-out infinite",width:"60%"}}/>
                              </div>
                              <div style={{fontSize:12,color:C.muted,marginTop:8}}>Extracting winery, vintage, varietal, region…</div>
                            </div>
                          )}
                          {labelResult && !labelResult.error && (
                            <div>
                              <div style={{fontSize:12,color:"#1e8449",fontWeight:700,marginBottom:8}}>✓ Label scanned — fields pre-filled below</div>
                              <div style={{fontSize:12,color:C.muted,lineHeight:1.8}}>
                                {labelResult.winery && <div>🏠 <strong style={{color:C.text}}>{labelResult.winery}</strong></div>}
                                {labelResult.vintage && <div>📅 <strong style={{color:C.text}}>{labelResult.vintage}</strong></div>}
                                {labelResult.varietal && <div>🍇 <strong style={{color:C.text}}>{labelResult.varietal}</strong></div>}
                                {labelResult.region && <div>📍 <strong style={{color:C.text}}>{labelResult.region}</strong></div>}
                              </div>
                              <div style={{display:"flex",gap:8,marginTop:10}}>
                                <button onClick={acceptLabelScan} style={{...btn(false,"#1e8449"),fontSize:12,padding:"6px 14px"}}>✓ Looks good</button>
                                <button onClick={retryLabelScan} style={{background:"none",border:`1px solid ${C.border}`,color:C.muted,padding:"6px 12px",borderRadius:6,cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>Retry</button>
                              </div>
                            </div>
                          )}
                          {labelResult?.error && (
                            <div>
                              <div style={{fontSize:13,color:"#a93226",marginBottom:8}}>⚠️ {labelResult.error}</div>
                              <button onClick={retryLabelScan} style={{background:"none",border:`1px solid ${C.border}`,color:C.muted,padding:"6px 12px",borderRadius:6,cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>Try again</button>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  )}

                  <div style={{display:"flex",alignItems:"center",gap:10,color:C.muted,fontSize:12}}>
                    <div style={{flex:1,height:1,background:C.border}}/>
                    or fill in manually
                    <div style={{flex:1,height:1,background:C.border}}/>
                  </div>

                  <FF label="Winery *"><input value={addForm.winery} onChange={e=>setAddForm(f=>({...f,winery:e.target.value}))} placeholder="e.g. Opus One" style={inp}/></FF>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
                    <FF label="Varietal"><select value={addForm.varietal} onChange={e=>setAddForm(f=>({...f,varietal:e.target.value}))} style={inp}>{VARIETALS.map(v=><option key={v}>{v}</option>)}</select></FF>
                    <FF label="Vintage"><input type="number" value={addForm.vintage} onChange={e=>setAddForm(f=>({...f,vintage:e.target.value}))} style={inp}/></FF>
                  </div>
                  <FF label="Region"><select value={addForm.region} onChange={e=>setAddForm(f=>({...f,region:e.target.value}))} style={inp}>{REGIONS.map(r=><option key={r}>{r}</option>)}</select></FF>
                  {!slotPending&&(
                    <FF label="Row">
                      <select value={addForm.row} onChange={e=>setAddForm(f=>({...f,row:e.target.value}))} style={inp}>
                        {ROW_DEFS.map(r=>{
                          const def = ROW_DEFS.find(x=>x.id===r.id);
                          const shelves = SHELF_STRUCTURE[def.type];
                          const totalSlots = shelves.reduce((a,b)=>a+b,0);
                          const usedSlots = wines.filter(w=>w.row===r.id).length;
                          const free = totalSlots - usedSlots;
                          return <option key={r.id} value={r.id}>{r.id} ({free} slots free)</option>;
                        })}
                      </select>
                    </FF>
                  )}
                  {!slotPending && addForm.row && (()=>{
                    const def = ROW_DEFS.find(r=>r.id===addForm.row);
                    const shelves = SHELF_STRUCTURE[def?.type||"full"];
                    return (
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
                        <FF label="Shelf (1-based)">
                          <select value={addForm.shelf} onChange={e=>setAddForm(f=>({...f,shelf:Number(e.target.value),slot:0}))} style={inp}>
                            {shelves.map((cap,si)=>{
                              const takenSlots = Array.from({length:cap},(_,sl)=>slotMap[slotKey(addForm.row,si,sl)]).filter(Boolean).length;
                              const freeSlots = cap - takenSlots;
                              return <option key={si} value={si}>Shelf {si+1} ({freeSlots} free)</option>;
                            })}
                          </select>
                        </FF>
                        <FF label="Slot (Position)">
                          <select value={addForm.slot} onChange={e=>setAddForm(f=>({...f,slot:Number(e.target.value)}))} style={inp}>
                            {Array.from({length:shelves[addForm.shelf]||3},(_,sl)=>{
                              const taken = !!slotMap[slotKey(addForm.row,addForm.shelf,sl)];
                              return <option key={sl} value={sl} disabled={taken}>{sl+1} {taken?"(occupied)":"(free)"}</option>;
                            })}
                          </select>
                        </FF>
                      </div>
                    );
                  })()}
                  <FF label="Occasions">
                    <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                      {OCCASIONS.map(o=><button key={o} onClick={()=>setAddForm(f=>({...f,occasion:f.occasion.includes(o)?f.occasion.filter(x=>x!==o):[...f.occasion,o]}))} style={{...chip(addForm.occasion.includes(o)),fontSize:12}}>{o}</button>)}
                    </div>
                  </FF>
                  <FF label="Notes"><textarea value={addForm.notes} onChange={e=>setAddForm(f=>({...f,notes:e.target.value}))} rows={2} placeholder="Tasting notes, source..." style={{...inp,resize:"vertical"}}/></FF>
                  <button disabled={!addForm.winery} onClick={()=>setAddStep(2)} style={btn(!addForm.winery)}>Next →</button>
                </div>
              )}

              {addStep===2&&(
                <div style={{display:"flex",flexDirection:"column",gap:14}}>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
                    <FF label="Price ($)"><input type="number" value={addForm.price} onChange={e=>setAddForm(f=>({...f,price:e.target.value}))} placeholder="0" style={inp}/></FF>
                    <FF label="Rating (0–100)"><input type="number" min={0} max={100} value={addForm.rating} onChange={e=>setAddForm(f=>({...f,rating:e.target.value}))} placeholder="e.g. 94" style={inp}/></FF>
                    <FF label="Drink From"><input type="number" value={addForm.drinkFrom} onChange={e=>setAddForm(f=>({...f,drinkFrom:e.target.value}))} style={inp}/></FF>
                    <FF label="Drink To"><input type="number" value={addForm.drinkTo} onChange={e=>setAddForm(f=>({...f,drinkTo:e.target.value}))} style={inp}/></FF>
                  </div>
                  <div style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:8,padding:12,fontSize:13,color:C.muted}}>
                    Auto sticker: <span style={{color:STICKER[autoSticker(Number(addForm.price),Number(addForm.drinkFrom),Number(addForm.drinkTo))]?.hex,fontWeight:700}}>● {STICKER[autoSticker(Number(addForm.price),Number(addForm.drinkFrom),Number(addForm.drinkTo))]?.label}</span>
                  </div>
                  <div style={{display:"flex",gap:10}}>
                    <button onClick={()=>setAddStep(1)} style={{...btn(false,C.dim),color:C.muted,flex:1}}>← Back</button>
                    <button onClick={()=>setAddStep(3)} style={{...btn(false),flex:2}}>Preview →</button>
                  </div>
                </div>
              )}

              {addStep===3&&(()=>{
                const sticker = autoSticker(Number(addForm.price),Number(addForm.drinkFrom),Number(addForm.drinkTo));
                const free = slotPending || nextFreeSlot(addForm.row);
                return (
                  <div>
                    <div style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:10,padding:20,marginBottom:16}}>
                      <div style={{fontSize:20,color:C.gold,fontWeight:700}}>{addForm.winery}</div>
                      <div style={{color:"#c4a882",fontSize:15,marginBottom:8}}>{addForm.vintage} · {addForm.varietal} · {addForm.region}</div>
                      <div style={{display:"flex",gap:10,flexWrap:"wrap",fontSize:13}}>
                        <Tag>${addForm.price||0}</Tag>
                        <Tag>🕐 {addForm.drinkFrom}–{addForm.drinkTo}</Tag>
                        {addForm.rating&&<Tag>⭐ {addForm.rating}</Tag>}
                        <span style={{background:STICKER[sticker]?.hex,color:STICKER[sticker]?.text,padding:"3px 10px",borderRadius:12,fontSize:12}}>● {STICKER[sticker]?.label}</span>
                      </div>
                      {free&&<div style={{marginTop:10,fontSize:13,color:"#7dbb30"}}>📍 Will be placed at: {slotPending?.row||addForm.row}, Shelf {(free.shelf+1)}, Slot {(free.slot+1)}</div>}
                      {!free&&<div style={{marginTop:10,fontSize:13,color:"#a93226"}}>⚠️ {addForm.row} is full — please choose another row.</div>}
                    </div>
                    <div style={{display:"flex",gap:10}}>
                      <button onClick={()=>setAddStep(2)} style={{...btn(false,C.dim),color:C.muted,flex:1}}>← Edit</button>
                      <button onClick={slotPending?assignToSlot:submitAdd} style={{...btn(!free,"#1e8449"),flex:2}}>✓ Add to Cellar</button>
                    </div>
                  </div>
                );
              })()}
            </div>
          </div>
        )}

        {/* ════════════════════ SEARCH ════════════════════ */}
        {view==="search" && (
          <div>
            <SectionTitle>Search & Recommendations</SectionTitle>
            {/* AI Sommelier */}
            <div style={{background:"linear-gradient(135deg,#140800,#1e1000)",border:`1px solid #5a3008`,borderRadius:12,padding:22,marginBottom:24}}>
              <div style={{fontSize:15,color:C.gold,fontWeight:700,marginBottom:4}}>🤖 AI Sommelier</div>
              <div style={{fontSize:13,color:C.muted,marginBottom:14}}>Choose an occasion and I'll recommend bottles from your ready-to-drink inventory.</div>
              <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:14}}>
                <div style={{flex:1,minWidth:150}}>
                  <label style={{fontSize:11,color:C.muted,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:1}}>Occasion</label>
                  <select value={aiOccasion} onChange={e=>setAiOccasion(e.target.value)} style={inp}>{OCCASIONS.map(o=><option key={o}>{o}</option>)}</select>
                </div>
                <div style={{width:90}}>
                  <label style={{fontSize:11,color:C.muted,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:1}}>Guests</label>
                  <input type="number" min={1} max={200} value={aiGuests} onChange={e=>setAiGuests(e.target.value)} style={inp}/>
                </div>
                <div style={{display:"flex",alignItems:"flex-end"}}>
                  <button onClick={getAiRec} disabled={aiLoading} style={btn(aiLoading,"#8b1a1a")}>{aiLoading?"Consulting...":"🍷 Recommend"}</button>
                </div>
              </div>
              {aiRec&&<div style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:8,padding:16,whiteSpace:"pre-wrap",fontSize:13,lineHeight:1.8,color:"#c4a882"}}>{aiRec}</div>}
            </div>

            {/* Search + filters */}
            <div style={{position:"relative",marginBottom:16}}>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search winery, varietal, region, vintage…" style={{...inp,paddingLeft:36,fontSize:14}}/>
              <span style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:C.muted}}>🔍</span>
            </div>
            <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:16}}>
              <FG label="Varietal" options={VARIETALS} selected={filters.varietal} onToggle={v=>toggleFilter("varietal",v)}/>
              <FG label="Region"   options={REGIONS}   selected={filters.region}   onToggle={v=>toggleFilter("region",v)}/>
              <FG label="Occasion" options={OCCASIONS}  selected={filters.occasion}  onToggle={v=>toggleFilter("occasion",v)}/>
              <FG label="Row"      options={ROW_DEFS.map(r=>r.id)} selected={filters.row} onToggle={v=>toggleFilter("row",v)}/>
            </div>
            <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:16,alignItems:"center"}}>
              <span style={{fontSize:11,color:C.muted,textTransform:"uppercase",letterSpacing:1}}>Sort:</span>
              {[["winery","Winery"],["vintage","Vintage"],["price","Price"],["rating","Rating"],["drinkFrom","Drink Window"]].map(([k,l])=>(
                <button key={k} onClick={()=>setSort(s=>({key:k,dir:s.key===k&&s.dir==="asc"?"desc":"asc"}))} style={chip(sort.key===k)}>
                  {l}{sort.key===k?(sort.dir==="asc"?" ↑":" ↓"):""}
                </button>
              ))}
              <span style={{marginLeft:"auto",fontSize:12,color:C.muted}}>{filteredWines.length} bottle{filteredWines.length!==1?"s":""}</span>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(270px,1fr))",gap:14}}>
              {filteredWines.map(w=><WineCard key={w.id} w={w} onClick={()=>setSelected(w)}/>)}
            </div>
            {!filteredWines.length&&<Empty msg="No wines match your search."/>}
          </div>
        )}

        {/* ════════════════════ VINTAGE CHART ════════════════════ */}
        {view==="vintage" && (
          <div>
            <SectionTitle>Vintage Chart</SectionTitle>
            <div style={{overflowX:"auto",marginBottom:32}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead>
                  <tr style={{borderBottom:`1px solid ${C.border}`}}>
                    <th style={{textAlign:"left",padding:"10px 12px",color:C.muted,fontWeight:600}}>Region</th>
                    {Object.keys(Object.values(VINTAGE_DATA)[0]).map(yr=>(
                      <th key={yr} style={{padding:"10px 6px",color:C.muted,fontWeight:600,textAlign:"center"}}>{yr}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {Object.entries(VINTAGE_DATA).map(([region,years])=>(
                    <tr key={region} style={{borderBottom:`1px solid ${C.surface}`}}>
                      <td style={{padding:"10px 12px",color:"#c4a882",fontWeight:600,whiteSpace:"nowrap"}}>{region}</td>
                      {Object.entries(years).map(([yr,score])=>(
                        <td key={yr} style={{padding:"6px 4px",textAlign:"center"}}>
                          <div style={{background:scoreColor(score),color:"#fff",borderRadius:5,padding:"4px 2px",fontWeight:700,fontSize:12,minWidth:32}}>{score}</div>
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{display:"flex",gap:16,flexWrap:"wrap",marginBottom:24}}>
              {[["#7b0d1e","97–100 Classic"],["#a93226","94–96 Outstanding"],["#ca6f1e","90–93 Excellent"],["#b7950b","87–89 Very Good"],["#616a6b","<87 Average"]].map(([c,l])=>(
                <div key={l} style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:C.muted}}><div style={{width:18,height:18,borderRadius:3,background:c}}/>{l}</div>
              ))}
            </div>
            <div style={{fontSize:14,color:C.gold,fontWeight:700,marginBottom:14}}>My Cellar Vintages</div>
            <div style={{display:"flex",alignItems:"flex-end",gap:2,height:160,paddingBottom:24,marginTop:8,overflowX:"auto"}}>
              {vintageBarData.map(d=>{
                const max=Math.max(...vintageBarData.map(x=>x.qty),1);
                const h=Math.round(d.qty/max*120);
                return (
                  <div key={d.vintage} style={{flex:1,minWidth:22,display:"flex",flexDirection:"column",alignItems:"center",gap:2}}>
                    <div style={{fontSize:9,color:C.muted}}>{d.qty}</div>
                    <div style={{width:"100%",height:h,background:"#8b1a1a",borderRadius:"3px 3px 0 0",minHeight:3}}/>
                    <div style={{fontSize:8,color:C.muted,writingMode:"vertical-rl",transform:"rotate(180deg)",whiteSpace:"nowrap"}}>{d.vintage}</div>
                  </div>
                );
              })}
            </div>
            </div>
        )}

        {/* ════════════════════ WINE LIST ════════════════════ */}
        {view==="list" && (
          <div>
            <SectionTitle>Wine List</SectionTitle>
            <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:16}}>
              {[["Reds",["Cabernet Sauvignon","Pinot Noir","Merlot","Syrah/Shiraz","Zinfandel","Malbec","Tempranillo","Sangiovese","Grenache","Nebbiolo"]],
                ["Whites",["Chardonnay","Sauvignon Blanc","Pinot Grigio","Riesling","Viognier"]],
                ["Sparkling",["Champagne/Sparkling"]],["Rosé",["Rosé"]],["Other",["Dessert/Port","Other"]]
              ].map(([cat,vars])=>{
                const count=wines.filter(w=>vars.includes(w.varietal)).length;
                return count>0?<button key={cat} onClick={()=>setFilters(f=>({...f,varietal:vars}))} style={chip(filters.varietal.length&&vars.every(v=>filters.varietal.includes(v)))}>{cat} ({count})</button>:null;
              })}
              <button onClick={()=>setFilters(f=>({...f,varietal:[]}))} style={chip(false)}>All</button>
            </div>
            <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:10,overflow:"hidden"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead>
                  <tr style={{background:"#1e0d04"}}>
                    {["","Winery","Vintage","Varietal","Region","Drink Window","Price","Location","Status"].map(h=>(
                      <th key={h} style={{padding:"11px 10px",textAlign:"left",color:C.muted,fontSize:11,fontWeight:600,letterSpacing:1,whiteSpace:"nowrap"}}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filteredWines.map((w,i)=>{
                    const ds = drinkStatus(w);
                    return (
                      <tr key={w.id} onClick={()=>setSelected(w)} style={{borderBottom:`1px solid ${C.surface}`,cursor:"pointer",background:i%2===0?"transparent":C.surface}}>
                        <td style={{padding:"9px 10px"}}><div style={{width:10,height:10,borderRadius:"50%",background:STICKER[w.sticker]?.hex}}/></td>
                        <td style={{padding:"9px 10px",color:C.text,fontWeight:600}}>{w.name||w.winery}</td>
                        <td style={{padding:"9px 10px",color:"#c4a882"}}>{w.vintage}</td>
                        <td style={{padding:"9px 10px",color:C.muted,fontSize:12}}>{w.varietal}</td>
                        <td style={{padding:"9px 10px",color:C.muted,fontSize:12}}>{w.region}</td>
                        <td style={{padding:"9px 10px",fontSize:12,color:ds.color,fontWeight:600}}>{w.drinkFrom}–{w.drinkTo}</td>
                        <td style={{padding:"9px 10px",color:C.gold,fontWeight:700}}>${w.price}</td>
                        <td style={{padding:"9px 10px",color:C.muted,fontSize:12}}>{w.row} · S{w.shelf+1}</td>
                        <td style={{padding:"9px 10px"}}><span style={{background:`${ds.color}22`,color:ds.color,padding:"2px 8px",borderRadius:10,fontSize:11,whiteSpace:"nowrap"}}>{ds.short}</span></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              {!filteredWines.length&&<div style={{padding:40,textAlign:"center",color:C.muted}}>No wines to display.</div>}
            </div>
            {/* Ready to drink report */}
            <div style={{marginTop:24,background:"#0a1a08",border:"1px solid #1e4018",borderRadius:10,padding:20}}>
              <div style={{fontSize:14,color:"#1e8449",fontWeight:700,marginBottom:12}}>✓ Ready to Drink Report — {currentYear}</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:8}}>
                {wines.filter(isRTD).map(w=>(
                  <div key={w.id} style={{display:"flex",justifyContent:"space-between",padding:"7px 12px",background:C.bg,borderRadius:6,fontSize:13}}>
                    <span style={{color:"#c4a882"}}>{w.winery} <span style={{color:C.muted}}>{w.vintage}</span></span>
                    <span style={{color:"#1e8449",fontSize:12}}>{w.row}·S{w.shelf+1}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

                {/* ════════════════════ ANALYTICS ════════════════════ */}
        {view==="charts" && (
          <div>
            <SectionTitle>Cellar Analytics</SectionTitle>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(340px,1fr))",gap:20}}>

              <ChartCard title="Bottles by Varietal">
                <div style={{display:"flex",flexDirection:"column",gap:6,paddingTop:4}}>
                  {varietalData.map((d,i)=>{
                    const max=Math.max(...varietalData.map(x=>x.value));
                    const colors=["#8b1a1a","#2471a3","#1e8449","#d4a017","#7d6608","#6c3483","#117a65","#784212"];
                    return (
                      <div key={d.name} style={{display:"flex",alignItems:"center",gap:8}}>
                        <div style={{width:120,fontSize:11,color:C.muted,textAlign:"right",flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{d.name}</div>
                        <div style={{flex:1,background:C.dim,borderRadius:3,height:18,overflow:"hidden"}}>
                          <div style={{width:`${Math.round(d.value/max*100)}%`,height:"100%",background:colors[i%colors.length],borderRadius:3,display:"flex",alignItems:"center",paddingLeft:6}}>
                            <span style={{fontSize:10,color:"#fff",fontWeight:700}}>{d.value}</span>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </ChartCard>

              <ChartCard title="Cellar by Sticker Color">
                <div style={{display:"flex",flexDirection:"column",gap:8,paddingTop:4}}>
                  {stickerData.map(d=>{
                    const max=Math.max(...stickerData.map(x=>x.value),1);
                    return (
                      <div key={d.name} style={{display:"flex",alignItems:"center",gap:10}}>
                        <div style={{width:10,height:10,borderRadius:"50%",background:d.color,flexShrink:0}}/>
                        <div style={{width:90,fontSize:12,color:C.muted,flexShrink:0}}>{d.name}</div>
                        <div style={{flex:1,background:C.dim,borderRadius:3,height:16,overflow:"hidden"}}>
                          <div style={{width:`${Math.round(d.value/max*100)}%`,height:"100%",background:d.color,borderRadius:3}}/>
                        </div>
                        <div style={{fontSize:12,color:C.text,fontWeight:700,width:28,textAlign:"right"}}>{d.value}</div>
                      </div>
                    );
                  })}
                </div>
              </ChartCard>

              <ChartCard title="Row Occupancy">
                <div style={{display:"flex",flexDirection:"column",gap:6,paddingTop:4}}>
                  {ROW_DEFS.map(r=>{
                    const used=wines.filter(w=>w.row===r.id).length;
                    const cap=ROW_CAPACITY[r.id];
                    const pct=cap>0?Math.round(used/cap*100):0;
                    const color=pct>90?"#a93226":pct>70?"#d4a017":"#1e8449";
                    return (
                      <div key={r.id} style={{display:"flex",alignItems:"center",gap:8}}>
                        <div style={{width:65,fontSize:11,color:C.muted,flexShrink:0}}>{r.id}</div>
                        <div style={{flex:1,background:C.dim,borderRadius:3,height:16,overflow:"hidden"}}>
                          <div style={{width:`${pct}%`,height:"100%",background:color,borderRadius:3}}/>
                        </div>
                        <div style={{fontSize:11,color:C.muted,width:45,textAlign:"right"}}>{used}/{cap}</div>
                      </div>
                    );
                  })}
                </div>
              </ChartCard>

              <ChartCard title="Vintages in Cellar">
                <div style={{display:"flex",alignItems:"flex-end",gap:2,height:160,paddingTop:10,paddingBottom:24,overflowX:"auto"}}>
                  {vintageBarData.map(d=>{
                    const max=Math.max(...vintageBarData.map(x=>x.qty),1);
                    const h=Math.round(d.qty/max*120);
                    return (
                      <div key={d.vintage} style={{flex:1,minWidth:22,display:"flex",flexDirection:"column",alignItems:"center",gap:2}}>
                        <div style={{fontSize:9,color:C.muted}}>{d.qty}</div>
                        <div style={{width:"100%",height:h,background:"#2471a3",borderRadius:"3px 3px 0 0",minHeight:3}}/>
                        <div style={{fontSize:8,color:C.muted,writingMode:"vertical-rl",transform:"rotate(180deg)",whiteSpace:"nowrap"}}>{d.vintage}</div>
                      </div>
                    );
                  })}
                </div>
              </ChartCard>

              <ChartCard title="Cellar Summary">
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                  {[
                    ["Total Bottles",totalBottles,C.text],
                    ["Total Capacity",TOTAL_CAPACITY,C.text],
                    ["Occupancy",`${Math.round(totalBottles/TOTAL_CAPACITY*100)}%`,totalBottles/TOTAL_CAPACITY>.9?"#a93226":"#1e8449"],
                    ["Ready to Drink",rtdCount,"#1e8449"],
                    ["Needs Aging",wines.filter(isNeedsAge).length,"#d4a017"],
                    ["Past Peak",wines.filter(isPastPeak).length,"#a93226"],
                    ["Total Value",`$${totalValue.toLocaleString()}`,C.gold],
                    ["Avg / Bottle",`$${Math.round(totalValue/Math.max(totalBottles,1))}`, "#c4a882"],
                  ].map(([l,v,c])=>(
                    <div key={l} style={{background:C.bg,borderRadius:7,padding:12}}>
                      <div style={{fontSize:10,color:C.dim,textTransform:"uppercase",letterSpacing:1}}>{l}</div>
                      <div style={{fontSize:20,fontWeight:700,color:c,marginTop:4}}>{v}</div>
                    </div>
                  ))}
                </div>
              </ChartCard>

              <ChartCard title="Drink Window Breakdown">
                <div style={{padding:"8px 0"}}>
                  {[
                    {label:"Ready Now",count:rtdCount,color:"#1e8449"},
                    {label:"Needs More Time",count:wines.filter(isNeedsAge).length,color:"#d4a017"},
                    {label:"Past Peak",count:wines.filter(isPastPeak).length,color:"#a93226"},
                  ].map(({label,count,color})=>{
                    const pct=Math.round(count/Math.max(totalBottles,1)*100);
                    return (
                      <div key={label} style={{marginBottom:16}}>
                        <div style={{display:"flex",justifyContent:"space-between",marginBottom:4,fontSize:13}}>
                          <span style={{color:"#c4a882"}}>{label}</span>
                          <span style={{color}}>{count} ({pct}%)</span>
                        </div>
                        <div style={{background:C.dim,borderRadius:3,height:8}}>
                          <div style={{width:`${pct}%`,height:"100%",background:color,borderRadius:3}}/>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </ChartCard>

            </div>
          </div>
        )}

      </main>

      {/* ════════════════════ WINE DETAIL MODAL ════════════════════ */}
      {selected && (
        <div onClick={()=>{setSelected(null);setEditing(false);}} style={{position:"fixed",inset:0,background:"rgba(0,0,0,.88)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
          <div onClick={e=>e.stopPropagation()} style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:14,padding:24,maxWidth:520,width:"100%",maxHeight:"90vh",overflowY:"auto",display:"flex",flexDirection:"column",gap:0}}>

            {/* ── HEADER ── */}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16}}>
              <div>
                <div style={{fontSize:20,color:C.gold,fontWeight:700,lineHeight:1.2}}>{selected.name || selected.varietal}</div>
                <div style={{fontSize:14,color:"#c4a882",marginTop:2}}>{selected.winery}</div>
                <div style={{fontSize:12,color:C.muted,marginTop:1}}>{selected.vintage} · {selected.varietal}</div>
                <div style={{fontSize:11,color:C.muted}}>{selected.region}{selected.country?` · ${selected.country}`:""}</div>
              </div>
              <div style={{display:"flex",gap:8,alignItems:"center",flexShrink:0}}>
                {!editing && (
                  <button
                    onClick={()=>{ setEditForm({...selected}); setEditing(true); }}
                    style={{background:`${C.gold}18`,border:`1px solid ${C.gold}`,color:C.gold,padding:"5px 13px",borderRadius:6,cursor:"pointer",fontSize:12,fontFamily:"inherit",fontWeight:600}}
                  >✏️ Edit</button>
                )}
                <button onClick={()=>{setSelected(null);setEditing(false);}} style={{background:"none",border:"none",color:C.muted,fontSize:20,cursor:"pointer",lineHeight:1,padding:"0 2px"}}>✕</button>
              </div>
            </div>

            {/* ══════════ VIEW MODE ══════════ */}
            {!editing && (<>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
                {[
                  ["💰 Price",`$${selected.price}`],
                  ["📍 Location",`${selected.row} · Shelf ${selected.shelf+1} · Slot ${selected.slot+1}`],
                  ["⭐ Rating",selected.rating?`${selected.rating}/100`:"—"],
                  ["🍾 Drink Window",`${selected.drinkFrom}–${selected.drinkTo}`],
                ].map(([l,v])=>(
                  <div key={l} style={{background:C.bg,borderRadius:7,padding:10}}>
                    <div style={{fontSize:10,color:C.goldDim}}>{l}</div>
                    <div style={{fontSize:13,color:C.text,fontWeight:600,marginTop:2}}>{v}</div>
                  </div>
                ))}
                <div style={{background:C.bg,borderRadius:7,padding:10,display:"flex",alignItems:"center",gap:8}}>
                  <div style={{width:11,height:11,borderRadius:"50%",background:STICKER[selected.sticker]?.hex,flexShrink:0}}/>
                  <div>
                    <div style={{fontSize:10,color:C.goldDim}}>🏷️ Sticker</div>
                    <div style={{fontSize:12,color:C.text,fontWeight:600}}>{STICKER[selected.sticker]?.label}</div>
                  </div>
                </div>
                {/* Rich drink status — full width */}
                {(()=>{ const ds = drinkStatus(selected); return (
                  <div style={{background:C.bg,borderRadius:7,padding:10,gridColumn:"span 2",borderLeft:`3px solid ${ds.color}`}}>
                    <div style={{fontSize:10,color:C.goldDim,marginBottom:3}}>📊 Drink Status</div>
                    <div style={{fontSize:13,color:ds.color,fontWeight:700}}>{ds.icon} {ds.label}</div>
                  </div>
                );})()}
              </div>

              {/* Tasting Notes — fixed height, scrollable */}
              <div style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:7,padding:12,marginBottom:14}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                  <div style={{fontSize:10,color:C.goldDim,textTransform:"uppercase",letterSpacing:1}}>Tasting Notes</div>
                  <button
                    onClick={()=>fetchTastingNotes(selected)}
                    disabled={fetchingNotes}
                    style={{background:fetchingNotes?"none":`${C.gold}18`,border:`1px solid ${fetchingNotes?C.border:C.gold}`,color:fetchingNotes?C.muted:C.gold,padding:"3px 10px",borderRadius:5,cursor:fetchingNotes?"not-allowed":"pointer",fontSize:11,fontFamily:"inherit",fontWeight:600}}
                  >
                    {fetchingNotes ? "⟳ Fetching…" : `🌐 ${selected.notes?"Refresh":"Fetch Notes"}`}
                  </button>
                </div>
                <div style={{maxHeight:72,overflowY:"auto",fontSize:12,color:"#c4a882",fontStyle:"italic",lineHeight:1.7}}>
                  {selected.notes || <span style={{color:C.muted}}>No notes yet — click Fetch Notes to pull tasting notes automatically.</span>}
                </div>
              </div>

              {selected.occasion?.length>0&&(
                <div style={{marginBottom:14}}>
                  <div style={{fontSize:10,color:C.goldDim,marginBottom:6,textTransform:"uppercase",letterSpacing:1}}>Occasions</div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:5}}>{selected.occasion.map(o=><Tag key={o}>{o}</Tag>)}</div>
                </div>
              )}

              <div style={{display:"flex",gap:8}}>
                <button onClick={async ()=>{
                  const { error: delErr } = await supabase.from("wines").delete().eq("id", String(selected.id));
                  if (delErr) { showToast(`Error: ${delErr.message}`, "error"); return; }
                  setSelected(null); setEditing(false);
                  showToast("Bottle removed","error");
                }} style={{flex:1,background:"#2d0a00",border:`1px solid #5a1a00`,color:"#a93226",padding:10,borderRadius:7,cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>🗑 Remove</button>
                <button onClick={async ()=>{
                  // Log consumption
                  await supabase.from("consumption_log").insert({
                    wine_id:    selected.id,
                    wine_name:  selected.name || selected.varietal,
                    winery:     selected.winery,
                    vintage:    selected.vintage,
                    row_loc:    selected.row,
                    shelf_loc:  selected.shelf + 1,
                    slot_loc:   selected.slot + 1,
                    price:      selected.price,
                  });
                  // Delete from cellar
                  const { error: delErr } = await supabase.from("wines").delete().eq("id", String(selected.id));
                  if (delErr) { showToast(`Error: ${delErr.message}`, "error"); return; }
                  // Send Pushover notification
                  try {
                    const { data: { session: s } } = await supabase.auth.getSession();
                    await fetch(NOTIFY_URL, {
                      method: "POST",
                      headers: { "Content-Type":"application/json", "Authorization": `Bearer ${s?.access_token}` },
                      body: JSON.stringify({
                        wineName: selected.name || selected.varietal,
                        winery:   selected.winery,
                        vintage:  selected.vintage,
                        row:      selected.row,
                        shelf:    selected.shelf + 1,
                        slot:     selected.slot + 1,
                        price:    selected.price,
                      }),
                    });
                  } catch(e) { console.warn("Notification failed:", e); }
                  setSelected(null); setEditing(false);
                  showToast("🥂 Cheers! Bottle consumed.");
                }} style={{flex:1,background:"#0a1e08",border:`1px solid #1a4010`,color:"#1e8449",padding:10,borderRadius:7,cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>🥂 Consume</button>
              </div>
            </>)}

            {/* ══════════ EDIT MODE ══════════ */}
            {editing && editForm && (<>
              <div style={{display:"flex",flexDirection:"column",gap:12}}>

                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                  <FF label="Wine Name"><input value={editForm.name||""} onChange={e=>setEditForm(f=>({...f,name:e.target.value}))} style={inp}/></FF>
                  <FF label="Winery"><input value={editForm.winery||""} onChange={e=>setEditForm(f=>({...f,winery:e.target.value}))} style={inp}/></FF>
                  <FF label="Vintage"><input type="number" value={editForm.vintage} onChange={e=>setEditForm(f=>({...f,vintage:Number(e.target.value)}))} style={inp}/></FF>
                  <FF label="Varietal">
                    <select value={editForm.varietal} onChange={e=>setEditForm(f=>({...f,varietal:e.target.value}))} style={inp}>
                      {VARIETALS.map(v=><option key={v}>{v}</option>)}
                    </select>
                  </FF>
                  <FF label="Region"><input value={editForm.region||""} onChange={e=>setEditForm(f=>({...f,region:e.target.value}))} style={inp}/></FF>
                  <FF label="Country"><input value={editForm.country||""} onChange={e=>setEditForm(f=>({...f,country:e.target.value}))} style={inp}/></FF>
                  <FF label="Price ($)"><input type="number" value={editForm.price} onChange={e=>setEditForm(f=>({...f,price:Number(e.target.value)}))} style={inp}/></FF>
                  <FF label="Rating (0–100)"><input type="number" min={0} max={100} value={editForm.rating||""} onChange={e=>setEditForm(f=>({...f,rating:Number(e.target.value)}))} style={inp}/></FF>
                  <FF label="Drink From"><input type="number" value={editForm.drinkFrom} onChange={e=>setEditForm(f=>({...f,drinkFrom:Number(e.target.value)}))} style={inp}/></FF>
                  <FF label="Drink To"><input type="number" value={editForm.drinkTo} onChange={e=>setEditForm(f=>({...f,drinkTo:Number(e.target.value)}))} style={inp}/></FF>
                </div>

                <FF label="Sticker Color">
                  <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                    {Object.entries(STICKER).map(([k,v])=>(
                      <button key={k} onClick={()=>setEditForm(f=>({...f,sticker:k}))}
                        style={{background:editForm.sticker===k?v.hex:`${v.hex}22`,border:`2px solid ${editForm.sticker===k?v.hex:C.border}`,color:editForm.sticker===k?v.text:C.muted,padding:"5px 12px",borderRadius:20,cursor:"pointer",fontSize:12,fontFamily:"inherit",transition:"all 0.15s"}}>
                        ● {v.label}
                      </button>
                    ))}
                  </div>
                </FF>

                <FF label="Row & Location">
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
                    <div>
                      <div style={{fontSize:10,color:C.muted,marginBottom:4}}>Row</div>
                      <select value={editForm.row} onChange={e=>setEditForm(f=>({...f,row:e.target.value}))} style={inp}>
                        {ROW_DEFS.map(r=><option key={r.id} value={r.id}>{r.id}</option>)}
                      </select>
                    </div>
                    <div>
                      <div style={{fontSize:10,color:C.muted,marginBottom:4}}>Shelf (1-based)</div>
                      <input type="number" min={1} value={editForm.shelf+1} onChange={e=>setEditForm(f=>({...f,shelf:Number(e.target.value)-1}))} style={inp}/>
                    </div>
                    <div>
                      <div style={{fontSize:10,color:C.muted,marginBottom:4}}>Slot (A=1,B=2,C=3)</div>
                      <input type="number" min={1} max={3} value={editForm.slot+1} onChange={e=>setEditForm(f=>({...f,slot:Number(e.target.value)-1}))} style={inp}/>
                    </div>
                  </div>
                </FF>

                <FF label="Tasting Notes">
                  <textarea value={editForm.notes||""} onChange={e=>setEditForm(f=>({...f,notes:e.target.value}))} rows={3} style={{...inp,resize:"vertical",maxHeight:90}}/>
                </FF>

                <div style={{display:"flex",gap:8,marginTop:4}}>
                  <button onClick={()=>setEditing(false)} style={{flex:1,background:"none",border:`1px solid ${C.border}`,color:C.muted,padding:10,borderRadius:7,cursor:"pointer",fontSize:13,fontFamily:"inherit"}}>← Cancel</button>
                  <button
                    onClick={async ()=>{
                      const { error } = await supabase.from("wines").update(wineToDb(editForm)).eq("id", String(editForm.id));
                      if (error) { showToast("Failed to save changes", "error"); return; }
                      setSelected({...editForm});
                      setEditing(false);
                      showToast(`${editForm.name||editForm.winery} updated`);
                    }}
                    style={{flex:2,background:"#1e8449",border:"none",color:"#fff",padding:10,borderRadius:7,cursor:"pointer",fontSize:13,fontFamily:"inherit",fontWeight:700}}
                  >✓ Save Changes</button>
                </div>
              </div>
            </>)}

          </div>
        </div>
      )}
    </div>
  );
}

// ─── MICRO COMPONENTS ─────────────────────────────────────────────────────────
function WineCard({w, onClick}) {
  const ds = drinkStatus(w);
  const displayName = w.name || w.varietal;
  return (
    <div onClick={onClick} style={{background:C.card,border:`1px solid ${C.border}`,borderLeft:`3px solid ${STICKER[w.sticker]?.hex}`,borderRadius:9,padding:14,cursor:"pointer",transition:"transform 0.1s"}}
      onMouseEnter={e=>e.currentTarget.style.transform="translateY(-2px)"}
      onMouseLeave={e=>e.currentTarget.style.transform="none"}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:4}}>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:13,fontWeight:700,color:C.text,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{displayName}</div>
          <div style={{fontSize:11,color:C.muted,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{w.winery}</div>
          <div style={{fontSize:12,color:"#c4a882",marginTop:2}}>{w.vintage} · {w.varietal}</div>
        </div>
        <div style={{textAlign:"right",marginLeft:8,flexShrink:0}}>
          <div style={{fontSize:15,fontWeight:700,color:C.gold}}>${w.price}</div>
          {w.rating>0&&<div style={{fontSize:11,color:C.muted}}>{w.rating}pts</div>}
        </div>
      </div>
      <div style={{fontSize:11,color:C.muted,marginBottom:8}}>{w.region}{w.country?` · ${w.country}`:""}</div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <span style={{fontSize:11,color:C.muted}}>{w.row} · S{w.shelf+1} · Sl{w.slot+1}</span>
        <span style={{fontSize:11,color:ds.color,background:`${ds.color}22`,padding:"2px 8px",borderRadius:10,whiteSpace:"nowrap"}}>{ds.short}</span>
      </div>
    </div>
  );
}

function SectionTitle({children}) {
  return <h2 style={{fontSize:20,color:C.gold,fontWeight:700,marginBottom:18,letterSpacing:1,borderBottom:`1px solid ${C.border}`,paddingBottom:10}}>{children}</h2>;
}
function ChartCard({title,children}) {
  return (
    <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:11,padding:18}}>
      <div style={{fontSize:12,color:C.gold,fontWeight:700,marginBottom:14,textTransform:"uppercase",letterSpacing:1}}>{title}</div>
      {children}
    </div>
  );
}
function FF({label,children}) {
  return (
    <div>
      <label style={{fontSize:11,color:C.muted,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:1}}>{label}</label>
      {children}
    </div>
  );
}
function FG({label,options,selected,onToggle}) {
  const [open,setOpen]=useState(false);
  const btnRef = useRef(null);
  const [pos,setPos]=useState({top:0,left:0});

  const handleOpen = () => {
    if (!open && btnRef.current) {
      const r = btnRef.current.getBoundingClientRect();
      setPos({ top: r.bottom + 4, left: r.left });
    }
    setOpen(o=>!o);
  };

  // Close on outside click
  useEffect(()=>{
    if (!open) return;
    const handler = (e) => { if (!btnRef.current?.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  },[open]);

  return (
    <div style={{position:"relative"}}>
      <button ref={btnRef} onClick={handleOpen} style={{...chip(selected.length>0),whiteSpace:"nowrap"}}>
        {label}{selected.length?` (${selected.length})`:""} {open?"▲":"▼"}
      </button>
      {open&&(
        <div style={{position:"fixed",top:pos.top,left:pos.left,background:C.card,border:`1px solid ${C.border}`,borderRadius:8,padding:6,zIndex:9999,minWidth:200,maxHeight:280,overflowY:"auto",boxShadow:"0 12px 40px rgba(0,0,0,.9)"}}>
          {options.map(o=>(
            <div key={o} onClick={()=>onToggle(o)} style={{padding:"7px 12px",cursor:"pointer",borderRadius:5,display:"flex",alignItems:"center",gap:8,fontSize:12,color:selected.includes(o)?C.gold:"#c4a882",background:selected.includes(o)?"#2a1400":"transparent"}}>
              <span style={{width:13,height:13,borderRadius:3,border:`1px solid ${selected.includes(o)?C.gold:C.border}`,background:selected.includes(o)?C.gold:"transparent",display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,color:C.bg,flexShrink:0}}>{selected.includes(o)?"✓":""}</span>
              {o}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
function Tag({children}) {
  return <span style={{background:C.dim,color:C.muted,padding:"3px 9px",borderRadius:5,fontSize:12}}>{children}</span>;
}
function Empty({msg}) {
  return <div style={{textAlign:"center",padding:60,color:C.goldDim,fontSize:15}}>🍷 {msg}</div>;
}
