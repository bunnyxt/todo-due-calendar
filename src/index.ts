import "dotenv/config";
import express from "express";
import ical, { ICalEventBusyStatus } from "ical-generator";
import fetch from "node-fetch";
import { webcrypto } from "crypto";

// Polyfill crypto for ical-generator
if (!globalThis.crypto) {
  globalThis.crypto = webcrypto as any;
}

const PORT = Number(process.env.PORT || 3000);
const ACCESS_TOKEN = process.env.ACCESS_TOKEN;
const REFRESH_TOKEN = process.env.REFRESH_TOKEN;
const CLIENT_ID = process.env.CLIENT_ID;
const ICS_TOKEN = process.env.ICS_TOKEN || "";

// Check if we have the required tokens
if (!ACCESS_TOKEN && !REFRESH_TOKEN) {
  console.error(
    "❌ Missing tokens! Please set either ACCESS_TOKEN or REFRESH_TOKEN in .env"
  );
  console.log("\nTo get tokens, run: pnpm run get-tokens");
  process.exit(1);
}

if (REFRESH_TOKEN && !CLIENT_ID) {
  console.error("❌ REFRESH_TOKEN provided but missing CLIENT_ID");
  console.log("Please set CLIENT_ID in your .env file");
  process.exit(1);
}

// ---------- Token Management ----------
let currentAccessToken: string | null = ACCESS_TOKEN || null;
let tokenExpiryTime: number = 0;

async function getValidAccessToken(): Promise<string> {
  // If we only have a static access token, use it
  if (!REFRESH_TOKEN) {
    if (!ACCESS_TOKEN) {
      throw new Error("No access token available");
    }
    return ACCESS_TOKEN;
  }

  const now = Date.now();

  // If we have a valid token that hasn't expired yet, return it
  if (currentAccessToken && now < tokenExpiryTime) {
    return currentAccessToken;
  }

  // Otherwise, refresh the token
  console.log("🔄 Refreshing access token...");

  const tokenUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
  const tokenData = new URLSearchParams({
    grant_type: "refresh_token",
    refresh_token: REFRESH_TOKEN,
    client_id: CLIENT_ID!,
    scope: "https://graph.microsoft.com/.default",
  });

  const response = await fetch(tokenUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: tokenData,
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Token refresh failed ${response.status}: ${errorText}`);
  }

  const tokenResponse = (await response.json()) as {
    access_token: string;
    expires_in: number;
    refresh_token?: string;
  };

  // Update the stored tokens
  currentAccessToken = tokenResponse.access_token;
  tokenExpiryTime = now + tokenResponse.expires_in * 1000 - 60000; // Refresh 1 minute before expiry

  // Update refresh token if a new one was provided
  if (tokenResponse.refresh_token) {
    console.log(
      "⚠️  New refresh token received. Please update your REFRESH_TOKEN environment variable."
    );
  }

  console.log("✅ Access token refreshed successfully");
  return currentAccessToken;
}

// ---------- Graph API - Fetch To Do Tasks ----------
const GRAPH = "https://graph.microsoft.com/v1.0";

type TodoTask = {
  id: string;
  title: string;
  status?: string;
  lastModifiedDateTime?: string;
  dueDateTime?: { dateTime?: string; timeZone?: string };
  categories?: string[];
};

async function fetchDueTasks(): Promise<TodoTask[]> {
  const accessToken = await getValidAccessToken();

  const headers = {
    Authorization: `Bearer ${accessToken}`,
    "Content-Type": "application/json",
  };

  const url = `${GRAPH}/me/todo/lists`;
  const listsRes = await fetch(url, { headers });
  if (!listsRes.ok) {
    const text = await listsRes.text();
    throw new Error(`List lists error ${listsRes.status}: ${text}`);
  }
  const listsJson = (await listsRes.json()) as {
    value?: { id: string; displayName: string }[];
  };
  const lists: { id: string; displayName: string }[] = listsJson.value ?? [];

  // Fetch tasks from all lists with rate limiting
  const allTaskArrays: TodoTask[][] = [];
  for (const list of lists) {
    const url = `${GRAPH}/me/todo/lists/${encodeURIComponent(list.id)}/tasks`;
    const r = await fetch(url, { headers });
    if (!r.ok) {
      const text = await r.text();
      throw new Error(`List tasks error ${r.status}: ${text}`);
    }
    const j = (await r.json()) as { value?: TodoTask[] };
    allTaskArrays.push(j.value ?? []);

    // Small delay between requests to avoid rate limiting
    await new Promise((resolve) => setTimeout(resolve, 100));
  }

  // Flatten and filter tasks
  const all: TodoTask[] = [];
  for (const tasks of allTaskArrays) {
    for (const t of tasks) {
      const due = t.dueDateTime?.dateTime;
      const done = (t.status || "").toLowerCase() === "completed";
      if (due && !done) all.push(t);
    }
  }
  return all;
}

// ---------- Generate ICS Calendar ----------
function buildCalendar(tasks: TodoTask[]) {
  const cal = ical({
    name: "todo-due-calendar",
    prodId: { company: "bunnyxt", product: "todo-due-ics", language: "zh-CN" },
    scale: "GREGORIAN",
    ttl: 60 * 30,
  });

  for (const t of tasks) {
    const iso = t.dueDateTime?.dateTime;
    if (!iso) continue;

    const [y, m, d] = iso.slice(0, 10).split("-").map(Number);
    const start = new Date(Date.UTC(y, m - 1, d));
    const end = new Date(Date.UTC(y, m - 1, d + 1));

    cal.createEvent({
      id: `todo-${t.id}`,
      start,
      end,
      allDay: true,
      summary: t.title || "(untitled task)",
      description:
        (t.categories?.length ? `#${t.categories.join(" #")}\n` : "") +
        "Source: Microsoft To Do",
      busystatus: ICalEventBusyStatus.FREE,
      lastModified: t.lastModifiedDateTime
        ? new Date(t.lastModifiedDateTime)
        : undefined,
    });
  }
  return cal;
}

// ---------- HTTP Server ----------
const app = express();

app.get("/todo-due.ics", async (req, res) => {
  try {
    if (ICS_TOKEN) {
      const token = String(req.query.token || "");
      if (token !== ICS_TOKEN) return res.status(401).send("Unauthorized");
    }

    const tasks = await fetchDueTasks();
    const cal = buildCalendar(tasks);

    res.setHeader("Content-Type", "text/calendar; charset=utf-8");
    res.setHeader("Cache-Control", "public, max-age=120");
    res.send(cal.toString());
  } catch (e: any) {
    console.error("❌ /todo-due.ics error:", e?.message || e);
    res.status(500).send("Calendar feed error");
  }
});

app.listen(PORT, () => {
  console.log(
    `📅 Todo Due Calendar ICS feed: http://localhost:${PORT}/todo-due.ics${
      ICS_TOKEN ? `?token=${ICS_TOKEN}` : ""
    }`
  );
});
