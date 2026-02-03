import fs from 'node:fs';
import path from 'node:path';
import { chromium } from 'playwright';

function normalizeNum(value) {
  if (value == null) return null;
  const v = String(value).trim();
  if (!v) return null;
  const cleaned = v.replace(/\./g, '').replace(',', '.');
  const n = Number(cleaned);
  return Number.isNaN(n) ? v : n;
}

function parseArgs() {
  const args = process.argv.slice(2);
  const out = {
    username: null,
    password: process.env.MFP_DIARY_PASSWORD || null,
    date: null,
    out: null,
    saveHtml: null,
    headful: true,
    browser: process.env.MFP_BROWSER || 'chromium',
    locale: process.env.MFP_LOCALE || 'pt',
    profileDir: process.env.MFP_PROFILE_DIR || path.join('scripts', '.mfp_profile'),
  };
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg === '--username') out.username = args[++i];
    else if (arg === '--password') out.password = args[++i];
    else if (arg === '--date') out.date = args[++i];
    else if (arg === '--out') out.out = args[++i];
    else if (arg === '--save-html') out.saveHtml = args[++i];
    else if (arg === '--headless') out.headful = false;
    else if (arg === '--browser') out.browser = args[++i];
    else if (arg === '--locale') out.locale = args[++i];
    else if (arg === '--profile-dir') out.profileDir = args[++i];
  }
  return out;
}

function waitForEnter() {
  return new Promise((resolve) => {
    process.stdin.resume();
    process.stdin.once('data', () => {
      process.stdin.pause();
      resolve();
    });
  });
}

async function extractData(page) {
  return page.evaluate(() => {
    const text = (el) => (el ? el.textContent.trim() : '');
    const lower = (s) => (s || '').toLowerCase();

    const result = { meals: [] };

    const dateEl = document.querySelector('button#date, #date, .date, .diary-date');
    if (dateEl) result.dateText = text(dateEl);

    let table = document.querySelector('#diary-table');
    if (!table) table = document.querySelector('table');
    if (!table) return result;

    const rows = Array.from(table.querySelectorAll('tr'));
    let currentMeal = null;
    const totals = {};

    for (const row of rows) {
      const rowText = lower(text(row));
      if (!rowText) continue;

      if (row.classList.contains('meal_header') || rowText.includes('cafe da manha') || rowText.includes('almo?o') || rowText.includes('almoco') || rowText.includes('jantar') || rowText.includes('lanches')) {
        currentMeal = text(row).replace(/\s+/g, ' ').trim();
        if (!currentMeal) currentMeal = 'Meal';
        result.meals.push({ name: currentMeal, items: [] });
        continue;
      }

      if (rowText.includes('totals') || rowText.includes('total') || rowText.includes('totais')) {
        const cells = Array.from(row.querySelectorAll('td,th')).map(c => text(c));
        if (cells.length >= 6) {
          totals.calories = cells[cells.length - 6];
          totals.carbs = cells[cells.length - 5];
          totals.fat = cells[cells.length - 4];
          totals.protein = cells[cells.length - 3];
          totals.sodium = cells[cells.length - 2];
          totals.sugar = cells[cells.length - 1];
        }
        continue;
      }

      const cells = Array.from(row.querySelectorAll('td')).map(c => text(c));
      if (cells.length >= 6) {
        const name = cells[0];
        if (!name || lower(name).includes('adicionar alimento')) continue;
        if (!currentMeal) {
          currentMeal = 'Unknown';
          result.meals.push({ name: currentMeal, items: [] });
        }
        const item = {
          name,
          calories: cells[cells.length - 6],
          carbs: cells[cells.length - 5],
          fat: cells[cells.length - 4],
          protein: cells[cells.length - 3],
          sodium: cells[cells.length - 2],
          sugar: cells[cells.length - 1],
        };
        result.meals[result.meals.length - 1].items.push(item);
      }
    }

    result.totals = totals;
    return result;
  });
}

async function run() {
  const args = parseArgs();
  if (!args.username) throw new Error('Missing --username');
  if (!args.password) throw new Error('Missing password. Set MFP_DIARY_PASSWORD env var or use --password.');

  const baseUrl = `https://www.myfitnesspal.com/${args.locale}/food/diary/${args.username}`;
  const url = args.date ? `${baseUrl}?date=${args.date}` : baseUrl;

  const launchOptions = {
    headless: !args.headful,
    args: ['--disable-blink-features=AutomationControlled'],
  };

  if (args.browser === 'msedge') {
    launchOptions.channel = 'msedge';
  }

  fs.mkdirSync(args.profileDir, { recursive: true });

  const context = await chromium.launchPersistentContext(args.profileDir, {
    ...launchOptions,
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    locale: 'pt-BR',
  });

  const page = await context.newPage();
  await page.goto(url, { waitUntil: 'domcontentloaded' });

  // Cloudflare challenge handling
  try {
    await page.waitForFunction(() => {
      const t = document.title.toLowerCase();
      return !(t.includes('moment') || t.includes('aguarde'));
    }, { timeout: 60000 });
  } catch {
    console.log('Cloudflare challenge still active. Please solve it in the opened browser, then press ENTER here.');
    await waitForEnter();
  }

  // Password gate
  const passwordInput = await page.$('input[type="password"], input[name="password"], input[name="diary_password"]');
  if (passwordInput) {
    await passwordInput.fill(args.password);
    await passwordInput.press('Enter');
    await page.waitForLoadState('domcontentloaded');
    await page.waitForTimeout(2000);
  }

  const data = await extractData(page);

  if (args.saveHtml) {
    fs.writeFileSync(args.saveHtml, await page.content(), 'utf-8');
  }

  await context.close();

  if (data.totals) {
    data.totals = Object.fromEntries(Object.entries(data.totals).map(([k, v]) => [k, normalizeNum(v)]));
  }
  if (data.meals) {
    for (const meal of data.meals) {
      for (const item of meal.items || []) {
        for (const k of ['calories', 'carbs', 'fat', 'protein', 'sodium', 'sugar']) {
          item[k] = normalizeNum(item[k]);
        }
      }
    }
  }

  if (!data.dateText) data.dateText = args.date || new Date().toISOString().slice(0, 10);

  let out = args.out;
  if (!out) {
    const slug = String(data.dateText).replace(/[^0-9\-]/g, '') || 'date';
    out = path.join('Relatorios_Intervals', `mfp_diary_${slug}.json`);
  }

  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, JSON.stringify(data, null, 2), 'utf-8');
  console.log(`Saved: ${out}`);
}

run().catch((err) => {
  console.error(err.message || err);
  process.exit(1);
});
