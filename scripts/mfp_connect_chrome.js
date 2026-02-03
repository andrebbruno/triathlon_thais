/**
 * MyFitnessPal Scraper - Connects to Chrome already open in debug mode
 *
 * BEFORE RUNNING:
 *   1. Run iniciar_chrome_logado.bat (for exercise data) or iniciar_chrome_debug.bat
 *   2. Login if needed
 *   3. When the diary is visible, run this script
 *
 * Usage:
 *   node mfp_connect_chrome.js                # last 7 days
 *   node mfp_connect_chrome.js --days 30      # last 30 days
 *   node mfp_connect_chrome.js --close        # close Chrome when done
 */

const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
const fs = require('fs');
const path = require('path');

puppeteer.use(StealthPlugin());

const CONFIG = {
  username: process.env.MFP_USERNAME || 'REPLACE_ME',
  diaryPassword: process.env.MFP_DIARY_PASSWORD || 'REPLACE_ME', // only used for public diary access
  defaultDays: 7,
  debugPort: 9223,
  extractExercise: true
};

function formatDate(date) {
  return date.toISOString().split('T')[0];
}

function subtractDays(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() - days);
  return result;
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function main() {
  const args = process.argv.slice(2);

  let days = CONFIG.defaultDays;
  let startDate = null;
  let endDate = new Date();
  let closeBrowser = false;

  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--days' || args[i] === '-d') days = parseInt(args[i + 1], 10) || 7;
    if (args[i] === '--start') startDate = new Date(args[i + 1]);
    if (args[i] === '--end') endDate = new Date(args[i + 1]);
    if (args[i] === '--close' || args[i] === '--close-browser' || args[i] === '--close-chrome') closeBrowser = true;
  }

  if (!startDate) startDate = subtractDays(endDate, days - 1);

  console.log('='.repeat(50));
  console.log('MyFitnessPal Scraper (Conectando ao Chrome)');
  console.log('='.repeat(50));
  console.log(`Periodo: ${formatDate(startDate)} a ${formatDate(endDate)}`);
  console.log('');

  if (CONFIG.username === 'REPLACE_ME') {
    console.log('ERRO: defina MFP_USERNAME no ambiente antes de rodar.');
    console.log('Exemplo: $env:MFP_USERNAME="seu_usuario"');
    process.exit(1);
  }

  console.log(`Conectando ao Chrome na porta ${CONFIG.debugPort}...`);

  let browser;
  try {
    browser = await puppeteer.connect({
      browserURL: `http://127.0.0.1:${CONFIG.debugPort}`,
      defaultViewport: null
    });
    console.log('Conectado!\n');
  } catch (err) {
    console.log('ERRO: Nao foi possivel conectar ao Chrome.');
    console.log('');
    console.log('Certifique-se de que:');
    console.log('  1. Voce executou iniciar_chrome_debug.bat (ou iniciar_chrome_logado.bat)');
    console.log('  2. O Chrome esta aberto com a pagina do MFP');
    console.log('');
    console.log(`Detalhes: ${err.message}`);
    process.exit(1);
  }

  const pages = await browser.pages();
  let page = pages.find(p => p.url().includes('myfitnesspal')) || pages[0];

  if (!page) {
    page = await browser.newPage();
  }

  console.log(`Pagina atual: ${page.url()}`);

  async function ensureDiaryAccess() {
    const pwdField = await page.$('input[type="password"]');
    if (pwdField) {
      await pwdField.type(CONFIG.diaryPassword, { delay: 50 });
      const btn = await page.$('button[type="submit"], input[type="submit"]');
      if (btn) await btn.click();
      await page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 60000 }).catch(() => {});
    }
  }

  async function gotoWithRetry(url, retries = 1) {
    for (let attempt = 0; attempt <= retries; attempt++) {
      try {
        await page.goto(url, { waitUntil: 'networkidle2', timeout: 60000 });
        return;
      } catch (err) {
        const msg = err.message || '';
        if ((msg.includes('detached Frame') || msg.includes('Execution context')) && attempt < retries) {
          page = await browser.newPage();
          continue;
        }
        throw err;
      }
    }
  }

  let content = await page.content();
  if (content.includes('Just a moment') || content.includes('cf_chl')) {
    console.log('\n*** CAPTCHA DETECTADO ***');
    console.log('Resolva o captcha no navegador e aguarde...\n');

    for (let i = 0; i < 60; i++) {
      await sleep(2000);
      content = await page.content();
      if (!content.includes('Just a moment') && !content.includes('cf_chl')) {
        console.log('Captcha resolvido! Continuando...\n');
        break;
      }
    }
  } else {
    console.log('Acesso direto (sem captcha)!\n');
  }

  await ensureDiaryAccess();

  const allData = [];
  let currentDate = new Date(startDate);

  console.log('Extraindo dados do diario:');
  console.log('-'.repeat(50));

  while (currentDate <= endDate) {
    const dateStr = formatDate(currentDate);
    process.stdout.write(`  ${dateStr}... `);

    try {
      await gotoWithRetry(`https://www.myfitnesspal.com/food/diary/${CONFIG.username}?date=${dateStr}`, 1);

      await sleep(2000);

      content = await page.content();
      if (content.includes('Just a moment') || content.includes('cf_chl')) {
        console.log('Cloudflare detectado - resolva manualmente e aguarde');
        await sleep(30000);
      }

      await ensureDiaryAccess();

      const data = await page.evaluate((extractExercise) => {
        const normalizeText = (text) => {
          if (!text) return '';
          try {
            return text.normalize('NFC');
          } catch (e) {
            return text;
          }
        };

        const stripDiacritics = (text) => {
          if (!text) return '';
          try {
            return text.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
          } catch (e) {
            return text;
          }
        };

        const result = {
          meals: { breakfast: [], lunch: [], dinner: [], snacks: [] },
          exercise: [],
          totals: { calories: 0, carbs: 0, fat: 0, protein: 0, sodium: 0, sugar: 0, caloriesBurned: 0 }
        };

        const parseNum = s => {
          if (!s) return 0;
          const lines = s.trim().split('\n').map(l => l.trim()).filter(l => l);
          if (lines.length === 0) return 0;
          const firstLine = lines[0];
          const cleaned = firstLine.replace(/,/g, '').replace(/[^\d.-]/g, '');
          const n = parseFloat(cleaned);
          return isNaN(n) ? 0 : n;
        };

        const table = document.querySelector('#diary-table') ||
                      document.querySelector('table.table0') ||
                      document.querySelector('.diary-table') ||
                      document.querySelector('table');

        if (!table) return result;

        const rows = table.querySelectorAll('tr');
        let currentMeal = 'snacks';

        rows.forEach(row => {
          const rowTextRaw = (row.textContent || '').toLowerCase();
          const rowText = stripDiacritics(rowTextRaw);
          const className = row.className || '';

          if (className.includes('meal_header') || className.includes('bottom') || row.querySelector('.meal_header')) {
            if (rowText.includes('breakfast') || rowText.includes('cafe')) currentMeal = 'breakfast';
            else if (rowText.includes('lunch') || rowText.includes('almoco')) currentMeal = 'lunch';
            else if (rowText.includes('dinner') || rowText.includes('jantar')) currentMeal = 'dinner';
            else if (rowText.includes('snack') || rowText.includes('lanche')) currentMeal = 'snacks';
            return;
          }

          if (rowText.includes('breakfast') && !rowText.includes('total')) currentMeal = 'breakfast';
          else if (rowText.includes('lunch') && !rowText.includes('total')) currentMeal = 'lunch';
          else if (rowText.includes('dinner') && !rowText.includes('total')) currentMeal = 'dinner';
          else if (rowText.includes('snack') && !rowText.includes('total')) currentMeal = 'snacks';

          const cells = row.querySelectorAll('td');
          if (cells.length < 2) return;

          const name = normalizeText(cells[0]?.textContent?.trim());

          if (!name || name.length < 2 ||
              rowText.includes('total') ||
              rowText.includes('add food') ||
              rowText.includes('goal') ||
              rowText.includes('remaining') ||
              rowText.includes('quick tools') ||
              rowText.includes('your ip') ||
              rowText.includes('ray id')) return;

          const item = {
            name: name,
            calories: parseNum(cells[1]?.textContent),
            carbs: parseNum(cells[2]?.textContent),
            fat: parseNum(cells[3]?.textContent),
            protein: parseNum(cells[4]?.textContent),
            sodium: parseNum(cells[5]?.textContent),
            sugar: parseNum(cells[6]?.textContent)
          };

          if (item.calories > 0 || item.protein > 0 || item.carbs > 0) {
            result.meals[currentMeal].push(item);
          }
        });

        Object.values(result.meals).flat().forEach(item => {
          result.totals.calories += item.calories || 0;
          result.totals.carbs += item.carbs || 0;
          result.totals.fat += item.fat || 0;
          result.totals.protein += item.protein || 0;
          result.totals.sodium += item.sodium || 0;
          result.totals.sugar += item.sugar || 0;
        });

        if (extractExercise) {
          const exerciseTable = document.querySelector('#diary-exercise-table') ||
                                document.querySelector('table.table1') ||
                                document.querySelector('[id*="exercise"]');

          if (exerciseTable) {
            const exRows = exerciseTable.querySelectorAll('tr');
            exRows.forEach(row => {
              const cells = row.querySelectorAll('td');
              if (cells.length >= 2) {
                const name = normalizeText(cells[0]?.textContent?.trim());
                const cal = parseNum(cells[1]?.textContent);
                if (name && cal > 0 &&
                    !name.toLowerCase().includes('total') &&
                    !name.toLowerCase().includes('add exercise') &&
                    !name.toLowerCase().includes('cardiovascular')) {
                  result.exercise.push({ name, caloriesBurned: cal });
                  result.totals.caloriesBurned += cal;
                }
              }
            });
          }

          if (result.totals.caloriesBurned === 0) {
            const pageText = document.body.innerText || '';
            const match = pageText.match(/exercise[:\s]*[-]?(\d+)/i) ||
                          pageText.match(/earned[:\s]*(\d+)/i);
            if (match) {
              result.totals.caloriesBurned = parseInt(match[1], 10) || 0;
            }
          }
        }

        return result;
      }, CONFIG.extractExercise);

      const totalItems = Object.values(data.meals).flat().length;
      const burned = data.totals.caloriesBurned || 0;
      const burnedStr = burned > 0 ? `, -${burned} exercicio` : '';
      console.log(`OK - ${totalItems} itens, ${Math.round(data.totals.calories)} kcal${burnedStr}`);

      allData.push({ date: dateStr, ...data });
    } catch (err) {
      console.log(`ERRO: ${err.message}`);
      allData.push({
        date: dateStr,
        error: err.message,
        meals: { breakfast: [], lunch: [], dinner: [], snacks: [] },
        totals: { calories: 0, carbs: 0, fat: 0, protein: 0 }
      });
    }

    currentDate.setDate(currentDate.getDate() + 1);
    await sleep(1500);
  }

  console.log('\nExtracao concluida! (navegador mantido aberto)');

  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const bom = '\uFEFF';

  const reportsDir = path.resolve(__dirname, '..', 'Relatorios_MFP');
  fs.mkdirSync(reportsDir, { recursive: true });
  console.log(`\nRelatorios MFP: ${reportsDir}`);

  const jsonFile = path.join(reportsDir, `mfp_diary_${timestamp}.json`);
  fs.writeFileSync(jsonFile, bom + JSON.stringify(allData, null, 2), 'utf8');
  console.log(`\nDados salvos em: ${jsonFile}`);

  const csvFile = path.join(reportsDir, `mfp_diary_${timestamp}.csv`);
  const csvRows = ['date,meal,name,calories,carbs,fat,protein,sodium,sugar'];

  allData.forEach(day => {
    if (day.meals) {
      Object.entries(day.meals).forEach(([meal, items]) => {
        items.forEach(item => {
          const row = [
            day.date,
            meal,
            `"${(item.name || '').replace(/"/g, '""')}"`,
            item.calories || 0,
            item.carbs || 0,
            item.fat || 0,
            item.protein || 0,
            item.sodium || 0,
            item.sugar || 0
          ];
          csvRows.push(row.join(','));
        });
      });
    }
  });

  fs.writeFileSync(csvFile, bom + csvRows.join('\n'), 'utf8');
  console.log(`Dados salvos em: ${csvFile}`);

  console.log('\n' + '='.repeat(50));
  console.log('RESUMO');
  console.log('='.repeat(50));

  const daysWithData = allData.filter(d => d.totals && d.totals.calories > 0);

  if (daysWithData.length > 0) {
    const totals = daysWithData.reduce((acc, d) => ({
      cal: acc.cal + d.totals.calories,
      burned: acc.burned + (d.totals.caloriesBurned || 0),
      prot: acc.prot + d.totals.protein,
      carb: acc.carb + d.totals.carbs,
      fat: acc.fat + d.totals.fat
    }), { cal: 0, burned: 0, prot: 0, carb: 0, fat: 0 });

    const avg = {
      cal: Math.round(totals.cal / daysWithData.length),
      burned: Math.round(totals.burned / daysWithData.length),
      prot: Math.round(totals.prot / daysWithData.length),
      carb: Math.round(totals.carb / daysWithData.length),
      fat: Math.round(totals.fat / daysWithData.length)
    };

    console.log(`Dias com dados: ${daysWithData.length}/${allData.length}`);
    console.log('\nMedia diaria:');
    console.log(`  Calorias consumidas: ${avg.cal} kcal`);
    if (avg.burned > 0) {
      console.log(`  Calorias exercicio:  -${avg.burned} kcal`);
      console.log(`  Calorias liquidas:   ${avg.cal - avg.burned} kcal`);
    }
    console.log(`  Proteina:            ${avg.prot}g`);
    console.log(`  Carboidratos:        ${avg.carb}g`);
    console.log(`  Gordura:             ${avg.fat}g`);
  } else {
    console.log('Nenhum dado extraido.');
  }

  console.log('\nConcluido!');

  if (closeBrowser) {
    await browser.close();
    console.log('Chrome encerrado.');
  } else {
    browser.disconnect();
  }
}

main().catch(err => {
  console.error('Erro fatal:', err.message);
  process.exit(1);
});
