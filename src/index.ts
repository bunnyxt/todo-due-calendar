import "dotenv/config";
import express from "express";
import ical from "ical-generator";
import fetch from "node-fetch";
import { webcrypto } from "crypto";

// Polyfill crypto for ical-generator
if (!globalThis.crypto) {
  globalThis.crypto = webcrypto as any;
}

const PORT = Number(process.env.PORT || 3000);
const ACCESS_TOKEN = process.env.ACCESS_TOKEN!;
const ICS_TOKEN = process.env.ICS_TOKEN || "";
const OUTLOOK_TZ = process.env.OUTLOOK_TZ || "Pacific Standard Time";

if (!ACCESS_TOKEN) {
  console.error("❌ Missing ACCESS_TOKEN in .env");
  process.exit(1);
}

// ---------- Graph 拉取 To Do ----------
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
  const headers = {
    Authorization: `Bearer ${ACCESS_TOKEN}`,
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

  const all: TodoTask[] = [];
  for (const list of lists) {
    const url = `${GRAPH}/me/todo/lists/${encodeURIComponent(list.id)}/tasks`;
    const r = await fetch(url, { headers });
    if (!r.ok) {
      const text = await r.text();
      throw new Error(`List tasks error ${r.status}: ${text}`);
    }
    const j = (await r.json()) as { value?: TodoTask[] };
    const tasks: TodoTask[] = j.value ?? [];
    for (const t of tasks) {
      const due = t.dueDateTime?.dateTime;
      const done = (t.status || "").toLowerCase() === "completed";
      if (due && !done) all.push(t);
    }
  }
  return all;
}

// ---------- 生成 ICS ----------
function buildCalendar(tasks: TodoTask[]) {
  const cal = ical({
    name: "To Do – Due Dates",
    prodId: { company: "your-name", product: "todo-due-ics", language: "EN" },
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
      lastModified: t.lastModifiedDateTime
        ? new Date(t.lastModifiedDateTime)
        : undefined,
    });
  }
  return cal;
}

// ---------- HTTP 服务 ----------
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
