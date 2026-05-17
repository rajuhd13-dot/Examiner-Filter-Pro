import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import axios from "axios";
import dotenv from "dotenv";

dotenv.config();

const app = express();
const PORT = 3000;

// The Google Apps Script Web App URL provided by the user
const GAS_URL = (process.env.GAS_DEPLOYMENT_URL || "https://script.google.com/macros/s/AKfycbxA0nzMWabxevsaWBoZinNmq7xBJvHcp9JNyQfn4Qs1gVcqlpmSD5yzYQhDofu7xYAl7w/exec").trim();
const S_ID    = process.env.SPREADSHEET_ID;
const SH_NAME = process.env.SHEET_NAME;

console.log(`[INIT] GAS Proxy target: ${GAS_URL.slice(0, 40)}... (Hidden S_ID: ${!!S_ID})`);

app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ limit: "50mb", extended: true }));

/*************** CACHE LAYER ***************/
const CACHE = new Map<string, { data: any, expiry: number }>();
const TTL = {
  OPTIONS: 14400000, // 4 hours
  SYNC: 30000,       // 30 seconds (nearly live check for sheet changes)
  FILTER: 3600000,   // 1 hour
  LOOKUP: 43200000   // 12 hours
};

/**
 * Prime Time Logic: 8:00 AM to 10:00 PM (Local Time)
 * During these hours, we prioritize cache hits for "instant" feel.
 */
function isPrimeTime() {
  const now = new Date();
  const hour = now.getHours();
  // Adjust for intended 8am-10pm window
  return hour >= 8 && hour < 22;
}

function getCache(key: string) {
  const entry = CACHE.get(key);
  if (entry && entry.expiry > Date.now()) {
    if (isPrimeTime()) {
      console.log(`[CACHE][PRIME] Instant hit: ${key}`);
    } else {
      console.log(`[CACHE] Hit: ${key}`);
    }
    return entry.data;
  }
  return null;
}

function setCache(key: string, data: any, ttlMs: number) {
  if (!data || !data.success) return; // Don't cache failures
  console.log(`[CACHE] Set: ${key} (ttl: ${ttlMs}ms)`);
  CACHE.set(key, { data, expiry: Date.now() + ttlMs });
}

/**
 * Proxy helper to call the Google Apps Script Web App
 * Handles redirects automatically.
 */
async function callGAS(action: string, payload: any = {}, method: 'GET' | 'POST' = 'POST') {
  try {
    // We trim the URL in case it has trailing spaces from .env
    const targetUrl = GAS_URL.trim();
    if (!targetUrl.startsWith("https://script.google.com")) {
       return { success: false, error: "Invalid GAS URL. It must start with https://script.google.com" };
    }

    const url = new URL(targetUrl);
    url.searchParams.set("action", action);
    
    // Inject "Hidden" IDs from environment
    if (S_ID) url.searchParams.set("ssId", S_ID);
    if (SH_NAME) url.searchParams.set("sheetName", SH_NAME);

    // For GET requests, we append each payload key as a query parameter
    if (method === 'GET' && payload && typeof payload === 'object') {
      Object.keys(payload).forEach(key => {
        if (payload[key] !== undefined && payload[key] !== null) {
          if (typeof payload[key] === 'object') {
            url.searchParams.set(key, JSON.stringify(payload[key]));
          } else {
            url.searchParams.set(key, String(payload[key]));
          }
        }
      });
    }

    const maskedUrl = `${url.origin}${url.pathname.slice(0, 15)}.../exec?action=${action}`;
    console.log(`[PROXY] Calling GAS [${method}] [${action}]... URL: ${maskedUrl}`);
    
    try {
      const response = await axios({
        method,
        url: url.toString(),
        data: method === 'POST' ? { action, ...payload } : undefined,
        headers: { 
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        },
        maxRedirects: 5,
        validateStatus: () => true,
        timeout: 180000
      });
      
      console.log(`[PROXY] Status: ${response.status} for action: ${action}`);

      // Check if we got HTML (sign of a login redirect or error page)
      if (typeof response.data === 'string' && (response.data.includes('<!DOCTYPE html>') || response.data.includes('<html'))) {
        console.error(`[PROXY] GAS [${action}] returned HTML instead of JSON.`);
        return { 
          success: false, 
          error: "GAS returned a login page. Apps Script MUST be deployed with 'Who has access: Anyone'.",
          advice: "1. Click 'Deploy' > 'New Deployment'. 2. Set 'Who has access' to 'Anyone'. 3. COPY THE NEW URL and update it in AI Studio Secrets."
        };
      }

      if (response.status >= 400) {
         return { 
           success: false, 
           error: `Google returned HTTP ${response.status}`,
           advice: "Check if the script is deleted or the URL is incorrect."
         };
      }

      return response.data;
    } catch (innerError: any) {
      throw innerError;
    }
  } catch (error: any) {
    const errorMsg = error.response 
      ? `Status ${error.response.status}: ${JSON.stringify(error.response.data).slice(0, 500)}` 
      : error.message;
    
    console.error(`[PROXY] GAS Connection Error [${action}]:`, errorMsg);
    
    let advice = "Please check your Google Apps Script deployment URL.";
    if (error.code === 'ECONNABORTED') advice = "The request timed out. The spreadsheet might be too large or GAS is slow.";
    if (error.response?.status === 404) advice = "The deployment URL returned a 404. Is it the correct Exec URL?";
    
    return { 
      success: false, 
      error: `Proxy Error: ${errorMsg}`,
      advice
    };
  }
}

// API Routes - Forwarding to GAS
app.get("/api/ping", async (req, res) => {
  try {
    const data = await callGAS("ping", {}, 'GET');
    res.json(data);
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get("/api/options", async (req, res) => {
  try {
    const cacheKey = "options";
    const cached = getCache(cacheKey);
    if (cached) return res.json(cached);

    const data = await callGAS("options", {}, 'GET');
    setCache(cacheKey, data, TTL.OPTIONS);
    res.json(data);
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post("/api/filter", async (req, res) => {
  try {
    const cacheKey = `filter_${JSON.stringify(req.body)}`;
    const cached = getCache(cacheKey);
    if (cached) return res.json(cached);

    const data = await callGAS("filter", req.body, 'POST');
    setCache(cacheKey, data, TTL.FILTER);
    res.json(data);
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get("/api/sync", async (req, res) => {
  try {
    const cacheKey = "sync";
    const cached = getCache(cacheKey);
    if (cached) return res.json(cached);

    const data = await callGAS("sync", {}, 'GET');
    setCache(cacheKey, data, TTL.SYNC);
    res.json(data);
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get("/api/lookup", async (req, res) => {
  try {
    const query = req.query.query as string;
    if (!query) return res.json({ success: true, found: false });

    const cacheKey = `lookup_${query.toLowerCase()}`;
    const cached = getCache(cacheKey);
    if (cached) return res.json(cached);

    const data = await callGAS("lookup", { query }, 'GET');
    setCache(cacheKey, data, TTL.LOOKUP);
    res.json(data);
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get("/api/clearCache", async (req, res) => {
  try {
    console.log("[CACHE] Clearing all local cache");
    CACHE.clear();
    const data = await callGAS("clearCache", {}, 'GET');
    res.json(data);
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// Vite Middleware for Development / Static serving for Production
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
    console.log(`Server running on http://localhost:${PORT}`);
    console.log(`GAS Proxy active to: ${GAS_URL}`);
  });
}

startServer();
