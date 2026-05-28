const fs = require('node:fs/promises');
const path = require('node:path');
const { createRequire } = require('node:module');

const bundledRequire = createRequire(
  'C:/Users/SeanPegoraro/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright/package.json'
);

const { chromium } = bundledRequire('playwright');

async function main() {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1440, height: 920 }, deviceScaleFactor: 1 });
  await page.goto('http://127.0.0.1:4173', { waitUntil: 'networkidle' });
  const firstHeading = await page.locator('h1').first().textContent();
  const visibleSave = await page.locator('button:has-text("Salva")').count();
  const screenshot = await page.screenshot({ fullPage: true });
  const outputPath = path.resolve('outputs', 'renderer-preview.png');
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await fs.writeFile(outputPath, screenshot);
  await browser.close();

  console.log(JSON.stringify({ firstHeading, visibleSave, outputPath }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
