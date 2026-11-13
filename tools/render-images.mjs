// Renders resume.html and cover-letter.html as HD PNG images (3x) of the paper card.
// Usage: node tools/render-images.mjs  (needs puppeteer-core + installed Chrome)
import { createRequire } from 'module';
import path from 'path';
import { fileURLToPath, pathToFileURL } from 'url';

const require = createRequire(import.meta.url);
let puppeteer;
try { puppeteer = require('puppeteer-core'); } catch { puppeteer = require('C:/salita/node_modules/puppeteer-core'); }

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = process.argv[2] || root;
const CHROME = 'C:/Program Files/Google/Chrome/Application/chrome.exe';

const browser = await puppeteer.launch({ executablePath: CHROME, headless: true });
const page = await browser.newPage();
await page.setViewport({ width: 908, height: 1000, deviceScaleFactor: 3 });
for (const [file, name] of [['resume.html', 'Earl-De-Castro-Resume'], ['cover-letter.html', 'Earl-De-Castro-Cover-Letter']]) {
  for (const scheme of ['dark', 'light']) {
    await page.emulateMediaFeatures([{ name: 'prefers-color-scheme', value: scheme }]);
    await page.goto(pathToFileURL(path.join(root, file)).href, { waitUntil: 'networkidle0' });
    await page.evaluate(() => document.fonts.ready);
    // Image = just the paper, with a little page background around it.
    await page.addStyleTag({ content: '.nav,.doc-bar{display:none!important} body{padding:24px;margin:0} .paper{margin:0 auto!important;max-width:860px}' });
    const file2 = path.join(out, `${name}${scheme === 'light' ? '-light' : ''}.png`);
    await page.screenshot({ path: file2, fullPage: true });
    console.log('saved', file2);
  }
}
await browser.close();
