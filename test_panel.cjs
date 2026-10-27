// Run with a local server: node test_panel.cjs http://127.0.0.1:8781
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { chromium } = require("playwright");

async function run() {
  const browser = await chromium.launch({ channel: process.env.PLAYWRIGHT_CHANNEL || "chrome", headless: true });
  const context = await browser.newContext({ viewport: { width: 1280, height: 720 }, reducedMotion: "reduce" });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "live-ads-ui-"));
  const csv = {
    name: "界面回归.csv", mimeType: "text/csv",
    buffer: Buffer.from("加热主播,加热开始时间,名称,总消耗,直播间曝光人数,直播间观看人数,总成交金额,总成交订单数\n测试主播,2026-10-27 10:00:00,成交ROI,100,1000,250,200,2\n"),
  };
  await context.addInitScript(() => {
    window.copiedText = "";
    Object.defineProperty(navigator, "clipboard", { configurable: true, value: {
      writeText: async (text) => { window.copiedText = text; },
      read: async () => { throw new DOMException("Read permission denied", "NotAllowedError"); },
    } });
  });

  async function checkViewport(width, height) {
    await page.setViewportSize({ width, height });
    const geometry = await page.evaluate(() => {
      const button = document.querySelector("#processButton").getBoundingClientRect();
      return { bottom: button.bottom, top: button.top, right: button.right,
        width: document.documentElement.scrollWidth, viewport: innerWidth, height: innerHeight };
    });
    assert.ok(geometry.top >= 0 && geometry.bottom <= geometry.height, "Generate remains visible");
    assert.ok(geometry.right <= geometry.viewport && geometry.width <= geometry.viewport, "No page overflow");
  }

  try {
    await page.goto(process.argv[2] || "http://127.0.0.1:8781");
    assert.equal(await page.locator("h1").count(), 0);
    assert.equal(await page.locator("#longPlanPanel").isVisible(), false);
    assert.equal(await page.locator("#settingsPanel").isVisible(), false);
    for (const [w, h] of [[1280, 720], [1080, 720], [1440, 900], [768, 1024], [390, 844], [320, 640]]) {
      await checkViewport(w, h);
    }
    await page.screenshot({ path: path.join(outputDir, "mobile-empty.png"), fullPage: true });
    await checkViewport(1280, 720);
    await page.screenshot({ path: path.join(outputDir, "desktop-empty.png"), fullPage: true });

    const imagePage = await context.newPage();
    await imagePage.setViewportSize({ width: 640, height: 360 });
    await imagePage.setContent('<body style="font:24px sans-serif;padding:32px;background:white"><h2>测试消耗截图</h2><p>消耗金额　100.00</p><p>当场成交金额　200.00</p><p>当场成交订单数　2</p></body>');
    const image = { name: '截图<&>.png', mimeType: "image/png", buffer: await imagePage.screenshot() };
    await imagePage.close();
    await page.locator("#screenshotInput").setInputFiles([image, image, image]);
    assert.equal(await page.locator("#screenshotZone .thumb-card").count(), 3);
    const contained = await page.evaluate(() => {
      const zone = document.querySelector("#screenshotZone").getBoundingClientRect();
      const thumbs = document.querySelector("#screenshotThumbs").getBoundingClientRect();
      return thumbs.top >= zone.top && thumbs.bottom <= zone.bottom && thumbs.right <= zone.right;
    });
    assert.ok(contained, "Thumbnails remain inside the upload box");
    await page.locator("[data-preview-screenshot='0']").click();
    assert.equal(await page.locator("#imageLightbox").isVisible(), true);
    await page.keyboard.press("Escape");
    await page.locator("[data-remove-screenshot='0']").click();
    assert.equal(await page.locator(".thumb-card").count(), 2);
    await page.screenshot({ path: path.join(outputDir, "desktop-images.png"), fullPage: true });
    await page.locator("#pasteScreenshotButton").click();
    await page.waitForFunction(() => document.querySelector("#message").textContent.includes("无法读取剪切板"));
    await page.locator("#pasteTarget").evaluate((element, base64) => {
      const bytes = Uint8Array.from(atob(base64), (char) => char.charCodeAt(0));
      const data = new DataTransfer();
      data.items.add(new File([bytes], "粘贴截图.png", { type: "image/png" }));
      element.dispatchEvent(new ClipboardEvent("paste", { bubbles: true, cancelable: true, clipboardData: data }));
    }, image.buffer.toString("base64"));
    assert.equal(await page.locator(".thumb-card").count(), 3, "Paste is handled once, not twice");
    assert.equal(await page.locator("#pastePanel").isVisible(), false);
    for (const [w, h] of [[768, 800], [390, 844], [320, 640]]) await checkViewport(w, h);
    await page.screenshot({ path: path.join(outputDir, "mobile-images.png"), fullPage: true });
    await checkViewport(1280, 720);

    await page.locator("#toggleLongPlan").click();
    await page.locator("[data-long-plan='消耗总金额']").fill("25");
    await page.locator("#toggleSettings").click();
    await checkViewport(1280, 720);
    await page.screenshot({ path: path.join(outputDir, "desktop-expanded.png"), fullPage: true });
    await page.locator("#resetWorkspaceButton").click();
    assert.equal(await page.locator(".thumb-card").count(), 0);
    assert.equal(await page.locator("[data-long-plan='消耗总金额']").inputValue(), "");
    assert.equal(await page.locator("#processButton").isDisabled(), true);
    assert.equal(await page.locator("#toggleLongPlan").getAttribute("aria-expanded"), "false");

    // Exercise the real export endpoint before mocking delayed responses.
    await page.locator("#fileInput").setInputFiles(csv);
    await page.locator("#processButton").click();
    await page.locator("#settlementPreview").waitFor({ state: "visible" });
    const settlementTop = (await page.locator("#settlementPreview").boundingBox()).y;
    assert.ok(settlementTop >= 80 && settlementTop < 110, "Generation scrolls below the fixed toolbar");
    assert.ok(await page.locator("#sheet3Table").evaluate((table) => Boolean(table.compareDocumentPosition(document.querySelector("#sheet2Table")) & Node.DOCUMENT_POSITION_FOLLOWING)));
    const download = await context.request.get(new URL(await page.locator("#downloadLink").getAttribute("href"), page.url()).href);
    assert.equal(download.status(), 200);
    assert.equal((await download.body()).subarray(0, 2).toString(), "PK");
    for (const key of ["sheet3", "sheet2"]) {
      await page.locator(`[data-copy='${key}-body']`).click();
      const expected = await page.locator(`#${key}Table`).evaluate((table) => [...table.querySelectorAll("tbody tr")]
        .map((row) => [...row.children].slice(1).map((cell) => cell.textContent).join("\t")).join("\n"));
      assert.equal(await page.evaluate(() => window.copiedText), expected, "Copy excludes header and first column, preserving percent formats");
      await page.locator(`[data-copy='${key}-all']`).click();
      assert.ok((await page.evaluate(() => window.copiedText)).includes("\t"));
    }
    await page.locator("#sheet3SelectButton").click();
    await page.locator("#sheet3Table tbody td").nth(1).click();
    await page.locator("#sheet3CopySelectedButton").click();
    assert.equal(await page.evaluate(() => window.copiedText), await page.locator("#sheet3Table tbody td").nth(1).textContent());
    await page.locator("#settlementPreview").scrollIntoViewIfNeeded();
    await page.screenshot({ path: path.join(outputDir, "desktop-results.png") });

    for (const endpoint of ["process", "ocr-long-plan"]) {
      await page.locator("#resetWorkspaceButton").click();
      let complete;
      const waiting = new Promise((resolve) => { complete = resolve; });
      let received;
      const started = new Promise((resolve) => { received = resolve; });
      const handler = async (route) => {
        received();
        await waiting;
        await route.fulfill({ json: { ok: true, fields: { 消耗总金额: 999 }, previews: { sheet3: [["旧结果"], [999]] }, fileName: "旧结果.xlsx" } }).catch(() => {});
      };
      await page.route(`**/api/${endpoint}`, handler);
      if (endpoint === "process") {
        await page.locator("#fileInput").setInputFiles(csv);
        await page.locator("#processButton").click();
      } else {
        await page.locator("#toggleLongPlan").click();
        await page.locator("#longPlanOcrInput").setInputFiles(image);
      }
      await started;
      await page.locator("#resetWorkspaceButton").click();
      complete();
      await page.unrouteAll({ behavior: "wait" });
      assert.equal(await page.locator("#previewSection").isVisible(), false);
      assert.equal(await page.locator("[data-long-plan='消耗总金额']").inputValue(), "");
      assert.equal(await page.locator("#ocrPreviewPanel").isVisible(), false);
      assert.equal(await page.locator("#processButton").isDisabled(), true);
      assert.equal(await page.locator("#ocrLongPlanButton").isDisabled(), false);
    }
    assert.deepEqual(errors, []);
    console.log(`PASS: responsive layout, in-box previews, paste, real export, copy ranges, selection, scroll, reset and cancellation. Screenshots: ${outputDir}`);
  } finally {
    await browser.close();
  }
}

run().catch((error) => { console.error(error); process.exitCode = 1; });
