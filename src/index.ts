import "dotenv/config";
import express from "express";
import {
  PublicClientApplication,
  Configuration,
  DeviceCodeRequest,
  AccountInfo,
} from "@azure/msal-node";
import ical, { ICalCalendar } from "ical-generator";
import fetch from "node-fetch";

const PORT = Number(process.env.PORT || 3000);
const CLIENT_ID = process.env.CLIENT_ID!;
const ICS_TOKEN = process.env.ICS_TOKEN || ""; // 为空则不校验
const OUTLOOK_TZ = process.env.OUTLOOK_TZ || "Pacific Standard Time";

if (!CLIENT_ID) {
  console.error("Missing CLIENT_ID in .env");
  process.exit(1);
}

// ========== 1) MSAL 配置（Device Code + 公共客户端） ==========
const msalConfig: Configuration = {
  auth: {
    clientId: CLIENT_ID,
    // 如果你用个人账号，consumers 最稳；如果你注册时选了“org+personal”，也可用 common
    authority: "https://login.microsoftonline.com/consumers",
  },
};
const msalApp = new PublicClientApplication(msalConfig);
const scopes = ["Tasks.Read", "offline_access", "openid", "profile"];

async function acquireToken(): Promise<{ accessToken: string }> {
  // 先尝试用缓存里的账号静默获取
  const cache = msalApp.getTokenCache();
  const accounts: AccountInfo[] = await cache.getAllAccounts();
  if (accounts.length) {
    const silent = await msalApp.acquireTokenSilent({
      account: accounts[0],
      scopes,
    });
    if (silent?.accessToken) return { accessToken: silent.accessToken };
  }
  // 设备码交互一次（控制台提示你到 https://microsoft.com/devicelogin 输入 code）
  const req: DeviceCodeRequest = {
    deviceCodeCallback: (info) => console.log(info.message),
    scopes,
  };
  const result = await msalApp.acquireTokenByDeviceCode(req);
  if (!result?.accessToken) throw new Error("Failed to get access token");
  return { accessToken: result.accessToken };
}

// ========== 2) Graph：拉取有 due 的未完成任务 ==========
const GRAPH = "https://graph.microsoft.com/v1.0";

type TodoTask = {
  id: string;
  title: string;
  status?: string;
  lastModifiedDateTime?: string;
  dueDateTime?: { dateTime?: string; timeZone?: string };
  categories?: string[];
};

async function fetchDueTasks(accessToken: string): Promise<TodoTask[]> {
  const headers = {
    Authorization: `Bearer ${accessToken}`,
    Prefer: `outlook.timezone="${OUTLOOK_TZ}"`,
    "Content-Type": "application/json",
  };

  // 1) 列出所有 To Do 列表
  const listsRes = await fetch(
    `${GRAPH}/me/todo/lists?$select=id,displayName`,
    { headers }
  );
  if (!listsRes.ok) throw new Error(`lists error: ${listsRes.status}`);
  const listsJson = await listsRes.json();
  const lists: { id: string; displayName: string }[] = listsJson.value ?? [];

  const all: TodoTask[] = [];

  // 2) 遍历每个列表，抓任务
  for (const list of lists) {
    // 只取关键字段；按需分页扩展（$top）
    const url =
      `${GRAPH}/me/todo/lists/${encodeURIComponent(list.id)}/tasks` +
      `?$select=id,title,status,dueDateTime,lastModifiedDateTime,categories&$top=100`;
    const r = await fetch(url, { headers });
    if (!r.ok) throw new Error(`tasks error: ${r.status}`);
    const j = await r.json();
    const tasks: TodoTask[] = j.value ?? [];

    for (const t of tasks) {
      const due = t.dueDateTime?.dateTime;
      const done = (t.status || "").toLowerCase() === "completed";
      if (due && !done) all.push(t);
    }
  }
  return all;
}

// ========== 3) 生成 ICS（全天事件，以“日期”为准） ==========
function buildCalendar(tasks: TodoTask[]): ICalCalendar {
  const cal = ical({
    name: "To Do – Due Dates",
    prodId: { company: "your-name", product: "todo-due-ics", language: "EN" },
    scale: "GREGORIAN",
    ttl: 60 * 30, // 30分钟
    method: "PUBLISH",
  });

  // 建议刷新频率（客户端可能忽略但加上无害）
  cal.createProperty("REFRESH-INTERVAL;VALUE=DURATION", "PT30M");
  cal.createProperty("X-PUBLISHED-TTL", "PT30M");

  for (const t of tasks) {
    const iso = t.dueDateTime!.dateTime!; // 一般是 "YYYY-MM-DDT00:00:00"
    // 直接取日期部分，避免涉及时区换算
    const [y, m, d] = iso.slice(0, 10).split("-").map(Number);
    const start = new Date(Date.UTC(y, m - 1, d)); // VALUE=DATE 用 UTC 日期即可
    const end = new Date(Date.UTC(y, m - 1, d + 1));

    cal.createEvent({
      id: `todo-${t.id}`, // -> VEVENT:UID
      start,
      end,
      allDay: true, // -> DTSTART;VALUE=DATE / DTEND;VALUE=DATE
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

// ========== 4) HTTP 服务 ==========
const app = express();

app.get("/todo-due.ics", async (req, res) => {
  try {
    // 简单 token 保护（可选）
    if (ICS_TOKEN) {
      const token = String(req.query.token || "");
      if (token !== ICS_TOKEN) return res.status(401).send("Unauthorized");
    }

    const { accessToken } = await acquireToken();
    const tasks = await fetchDueTasks(accessToken);
    const cal = buildCalendar(tasks);

    res.setHeader("Content-Type", "text/calendar; charset=utf-8");
    res.setHeader("Cache-Control", "public, max-age=120");
    res.send(cal.toString());
  } catch (err) {
    console.error(err);
    res.status(500).send("Calendar feed error");
  }
});

app.listen(PORT, () => {
  console.log(
    `ICS feed on http://localhost:${PORT}/todo-due.ics${
      ICS_TOKEN ? `?token=${ICS_TOKEN}` : ""
    }`
  );
});
