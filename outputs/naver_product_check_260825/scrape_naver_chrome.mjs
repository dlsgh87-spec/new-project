import { spawn } from "node:child_process";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const chromePath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe";
const productUrl = "https://brand.naver.com/dongwonpet/products/13136436594";
const port = 9331;
const profileDir = mkdtempSync(join(tmpdir(), "codex-chrome-profile-"));

const chrome = spawn(chromePath, [
  "--headless=new",
  "--disable-gpu",
  "--disable-extensions",
  "--no-first-run",
  "--remote-allow-origins=*",
  `--remote-debugging-port=${port}`,
  `--user-data-dir=${profileDir}`,
  "--lang=ko-KR",
  productUrl,
], { stdio: "ignore" });

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function getJson(url, opts = {}) {
  const res = await fetch(url, opts);
  if (!res.ok) throw new Error(`${res.status} ${res.statusText}: ${url}`);
  return res.json();
}

async function send(ws, method, params = {}) {
  const id = send.nextId++;
  ws.send(JSON.stringify({ id, method, params }));
  return new Promise((resolve, reject) => {
    const onMessage = (event) => {
      const data = JSON.parse(event.data);
      if (data.id !== id) return;
      ws.removeEventListener("message", onMessage);
      if (data.error) reject(new Error(JSON.stringify(data.error)));
      else resolve(data.result);
    };
    ws.addEventListener("message", onMessage);
  });
}
send.nextId = 1;

async function main() {
  let tabs;
  for (let i = 0; i < 40; i++) {
    try {
      tabs = await getJson(`http://127.0.0.1:${port}/json`);
      if (tabs?.length) break;
    } catch {}
    await sleep(250);
  }

  if (!tabs?.length) throw new Error("Chrome DevTools tab not found.");
  const tab = tabs.find((item) => item.url.includes("13136436594")) || tabs[0];
  const ws = new WebSocket(tab.webSocketDebuggerUrl);

  await new Promise((resolve, reject) => {
    ws.addEventListener("open", resolve, { once: true });
    ws.addEventListener("error", reject, { once: true });
  });

  await send(ws, "Page.enable");
  await send(ws, "Runtime.enable");
  await send(ws, "Page.navigate", { url: productUrl });
  await sleep(12000);

  const result = await send(ws, "Runtime.evaluate", {
    expression: `(() => {
      const text = document.body ? document.body.innerText : "";
      const title = document.title || "";
      const html = document.documentElement ? document.documentElement.outerHTML : "";
      const matches = Array.from(new Set((text + "\\n" + html).match(/.{0,40}(?:85g|160g|170g|400g|중량|토핑|뉴트리플랜|레드미트|캔).{0,80}/g) || []));
      return { title, url: location.href, textLength: text.length, text: text.slice(0, 12000), matches: matches.slice(0, 80) };
    })()`,
    returnByValue: true,
    awaitPromise: true,
  });

  console.log(JSON.stringify(result.result.value, null, 2));
  ws.close();
}

main()
  .catch((error) => {
    console.error(error.stack || error.message);
    process.exitCode = 1;
  })
  .finally(() => {
    chrome.kill();
  });
