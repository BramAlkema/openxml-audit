/**
 * Drive one dirty editor session in the bundled EuroOffice example app.
 *
 * The Python transport runs this twice: first inserting a unique marker,
 * then removing it. Each pass must make EuroOffice's own document-state
 * callback mark the page title dirty. Browser exit closes the editing
 * session; the example app's callback then persists the editor-produced
 * DOCX.
 */

import { chromium } from "playwright";

const baseUrl = process.env.EUROOFFICE_EXAMPLE_URL;
const fileName = process.env.EUROOFFICE_FILE_NAME;
const action = process.env.EUROOFFICE_PROBE_ACTION;
const marker = process.env.EUROOFFICE_PROBE_MARKER;

if (!baseUrl || !fileName || !action || !marker) {
  throw new Error(
    "EUROOFFICE_EXAMPLE_URL, EUROOFFICE_FILE_NAME, " +
      "EUROOFFICE_PROBE_ACTION, and EUROOFFICE_PROBE_MARKER are required",
  );
}
if (!new Set(["insert", "remove"]).has(action)) {
  throw new Error(`Unsupported EUROOFFICE_PROBE_ACTION: ${action}`);
}

const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
const diagnostics = [];
page.on("console", (message) => diagnostics.push(`console:${message.type()}:${message.text()}`));
page.on("pageerror", (error) => diagnostics.push(`pageerror:${error.message}`));
page.on("requestfailed", (request) => {
  diagnostics.push(`requestfailed:${request.url()}:${request.failure()?.errorText ?? "unknown"}`);
});

try {
  const editorUrl = new URL("editor", baseUrl);
  editorUrl.searchParams.set("mode", "edit");
  editorUrl.searchParams.set("fileName", fileName);
  editorUrl.searchParams.set("userid", "uid-0");
  editorUrl.searchParams.set("lang", "en");

  await page.goto(editorUrl.toString(), {
    waitUntil: "domcontentloaded",
    timeout: 60_000,
  });
  const editorFrame = await page.waitForSelector("iframe", { timeout: 60_000 });
  const frame = await editorFrame.contentFrame();
  if (!frame) {
    throw new Error("EuroOffice editor iframe did not attach");
  }

  await frame.waitForLoadState("domcontentloaded", { timeout: 60_000 });
  await page.waitForTimeout(12_000);
  const editorInput = frame.locator("#area_id");
  if (await editorInput.count()) {
    await editorInput.focus();
  } else {
    await frame.locator("body").click({ position: { x: 700, y: 500 } });
  }
  await page.keyboard.press("Control+End");
  if (action === "insert") {
    await page.keyboard.type(marker);
  } else {
    for (let index = 0; index < marker.length; index += 1) {
      await page.keyboard.press("Backspace");
    }
  }
  await page.waitForFunction(() => document.title.startsWith("*"), null, {
    timeout: 10_000,
  });
  await page.waitForTimeout(2_000);

  process.stdout.write(
    JSON.stringify({
      opened: true,
      action,
      dirty: true,
      disconnect: "browser_exit",
      pageUrl: page.url(),
      frameUrl: frame.url(),
      diagnostics,
    }),
  );
} finally {
  await browser.close();
}
