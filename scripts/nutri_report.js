/* Nutrition weekly report: merges Intervals + MFP data into a separate report.
 *
 * Usage:
 *   node scripts/nutri_report.js
 *   node scripts/nutri_report.js --intervals-current <path> --intervals-next <path> --mfp <path>
 */

const fs = require('fs');
const path = require('path');

const REPO_ROOT = path.resolve(__dirname, '..');
const INTERVALS_DIR = path.join(REPO_ROOT, 'Relatorios_Intervals');
const MFP_DIR = path.join(REPO_ROOT, 'Relatorios_MFP');
const OUT_DIR = path.join(REPO_ROOT, 'Relatorios_Nutri');

function readJson(filePath) {
  const raw = fs.readFileSync(filePath, 'utf8');
  const text = raw.replace(/^\uFEFF/, '');
  return JSON.parse(text);
}

function listFiles(dir, prefix, suffix) {
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir)
    .filter((f) => f.startsWith(prefix) && f.endsWith(suffix))
    .sort();
}

function parseIntervalsRange(name) {
  const match = name.match(/^report_(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})\.json$/);
  if (!match) return null;
  return { start: match[1], end: match[2], name };
}

function pickLatestIntervals() {
  const files = listFiles(INTERVALS_DIR, 'report_', '.json');
  if (files.length === 0) return null;
  return path.join(INTERVALS_DIR, files[files.length - 1]);
}

function pickNextIntervals(currentPath) {
  const files = listFiles(INTERVALS_DIR, 'report_', '.json')
    .map((f) => parseIntervalsRange(f))
    .filter(Boolean)
    .sort((a, b) => a.start.localeCompare(b.start));

  const currentName = path.basename(currentPath);
  const current = parseIntervalsRange(currentName);
  if (!current) return null;

  const next = files.find((f) => f.start > current.end);
  if (!next) return null;
  return path.join(INTERVALS_DIR, next.name);
}

function pickLatestMfp() {
  const files = listFiles(MFP_DIR, 'mfp_diary_', '.json');
  if (files.length === 0) return null;
  return path.join(MFP_DIR, files[files.length - 1]);
}

function sum(values) {
  return values.reduce((acc, v) => acc + (v || 0), 0);
}

function round(value, digits = 0) {
  const pow = Math.pow(10, digits);
  return Math.round((value || 0) * pow) / pow;
}

function formatDateRange(start, end) {
  return `${start} a ${end}`;
}

function calcMfpSummary(mfpData) {
  const days = Array.isArray(mfpData) ? mfpData : [];
  const daysWithData = days.filter((d) => d.totals && (d.totals.calories || 0) > 0);

  const totalCalories = sum(daysWithData.map((d) => d.totals.calories || 0));
  const totalProtein = sum(daysWithData.map((d) => d.totals.protein || 0));
  const totalCarbs = sum(daysWithData.map((d) => d.totals.carbs || 0));
  const totalFat = sum(daysWithData.map((d) => d.totals.fat || 0));
  const totalBurned = sum(daysWithData.map((d) => d.totals.caloriesBurned || 0));

  const count = daysWithData.length || 0;
  const avg = {
    calories: count ? round(totalCalories / count, 0) : 0,
    protein: count ? round(totalProtein / count, 0) : 0,
    carbs: count ? round(totalCarbs / count, 0) : 0,
    fat: count ? round(totalFat / count, 0) : 0,
    burned: count ? round(totalBurned / count, 0) : 0
  };

  return {
    daysTotal: days.length,
    daysWithData: count,
    totals: {
      calories: round(totalCalories, 0),
      protein: round(totalProtein, 0),
      carbs: round(totalCarbs, 0),
      fat: round(totalFat, 0),
      burned: round(totalBurned, 0)
    },
    averages: avg,
    netAverage: avg.calories - avg.burned
  };
}

function calcIntervalsSummary(report) {
  const activities = report.atividades || [];
  const summary = {
    totalHours: report.semana?.tempo_total_horas || 0,
    totalTss: report.semana?.carga_total_tss || 0,
    totalDistanceKm: report.semana?.distancia_total_km || 0,
    byTypeMinutes: {},
    countByType: {}
  };

  activities.forEach((a) => {
    const type = a.type || 'Unknown';
    const minutes = a.moving_time_min || 0;
    summary.byTypeMinutes[type] = round((summary.byTypeMinutes[type] || 0) + minutes, 1);
    summary.countByType[type] = (summary.countByType[type] || 0) + 1;
  });

  return summary;
}

function calcPlannedSummary(report, weightKg) {
  if (!report || !report.treinos_planejados || !weightKg || weightKg <= 0) return null;
  const planned = report.treinos_planejados || [];
  const byTypeMinutes = {};
  planned.forEach((p) => {
    const type = p.type || 'Unknown';
    const minutes = p.moving_time_min || 0;
    byTypeMinutes[type] = round((byTypeMinutes[type] || 0) + minutes, 1);
  });

  const kcalByType = {};
  const metMap = {
    Ride: 8,
    Run: 9,
    Swim: 8,
    Workout: 3.5,
    WeightTraining: 3.5
  };

  let totalKcal = 0;
  Object.entries(byTypeMinutes).forEach(([type, minutes]) => {
    const met = metMap[type] || 0;
    const hours = minutes / 60;
    const kcal = met > 0 ? met * weightKg * hours : 0;
    if (kcal > 0) {
      kcalByType[type] = round(kcal, 0);
      totalKcal += kcal;
    }
  });

  return {
    plannedCount: planned.length,
    byTypeMinutes,
    estimatedTrainingKcal: round(totalKcal, 0),
    estimatedDailyKcal: planned.length ? round(totalKcal / 7, 0) : 0,
    kcalByType
  };
}

function writeReport(outPath, data, mdPath) {
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, JSON.stringify(data, null, 2), 'utf8');

  const md = [];
  md.push(`# Relatorio Nutricional (${formatDateRange(data.period.start, data.period.end)})`);
  md.push('');
  md.push('## Resumo da Semana Executada');
  md.push(`- Treino: ${data.training.totalHours} h, ${data.training.totalTss} TSS`);
  md.push(`- Distancia total: ${data.training.totalDistanceKm} km`);
  md.push(`- MFP: ${data.nutrition.daysWithData}/${data.nutrition.daysTotal} dias com dados`);
  md.push(`- Media calorias: ${data.nutrition.averages.calories} kcal`);
  md.push(`- Proteina media: ${data.nutrition.averages.protein} g`);
  md.push(`- Carboidratos medio: ${data.nutrition.averages.carbs} g`);
  md.push(`- Gordura media: ${data.nutrition.averages.fat} g`);
  if (data.nutrition.averages.burned > 0) {
    md.push(`- Calorias exercicio (MFP): -${data.nutrition.averages.burned} kcal`);
    md.push(`- Calorias liquidas: ${data.nutrition.netAverage} kcal`);
  }
  md.push('');
  md.push('## Critica do Executado');
  md.push('- (A preencher pela agente Nutri)');
  md.push('');
  md.push('## Timing de Suplementos');
  md.push('- (A preencher pela agente Nutri)');
  md.push('');
  md.push('## Proxima Semana (Planejada)');
  if (data.nextWeek) {
    md.push(`- Treinos planejados: ${data.nextWeek.plannedCount}`);
    md.push(`- Estimativa gasto treino: ${data.nextWeek.estimatedTrainingKcal} kcal (media ${data.nextWeek.estimatedDailyKcal} kcal/dia)`);
  } else {
    md.push('- Sem dados de planejamento para a proxima semana.');
  }
  md.push('');
  md.push('## Observacoes');
  md.push('- Estimativas caloricas baseadas em MET (valor aproximado).');

  fs.writeFileSync(mdPath, md.join('\n'), 'utf8');
}

function parseArgs() {
  const args = process.argv.slice(2);
  const out = {};
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--intervals-current') out.current = args[i + 1];
    if (args[i] === '--intervals-next') out.next = args[i + 1];
    if (args[i] === '--mfp') out.mfp = args[i + 1];
  }
  return out;
}

function main() {
  const args = parseArgs();
  const currentPath = args.current || pickLatestIntervals();
  const nextPath = args.next || (currentPath ? pickNextIntervals(currentPath) : null);
  const mfpPath = args.mfp || pickLatestMfp();

  if (!currentPath || !mfpPath) {
    console.error('Missing inputs: current Intervals report or MFP report not found.');
    process.exit(1);
  }

  const current = readJson(currentPath);
  const next = nextPath ? readJson(nextPath) : null;
  const mfp = readJson(mfpPath);

  const weightKg = current.metricas?.peso_atual;
  const trainingSummary = calcIntervalsSummary(current);
  const nutritionSummary = calcMfpSummary(mfp);
  const nextWeekSummary = next ? calcPlannedSummary(next, weightKg) : null;
  const critique = [];

  const outFileName = `nutri_report_${current.semana.inicio}_${current.semana.fim}.json`;
  const outMdName = `nutri_report_${current.semana.inicio}_${current.semana.fim}.md`;
  const outPath = path.join(OUT_DIR, outFileName);
  const mdPath = path.join(OUT_DIR, outMdName);

  const payload = {
    period: { start: current.semana.inicio, end: current.semana.fim },
    training: trainingSummary,
    nutrition: nutritionSummary,
    nextWeek: nextWeekSummary,
    critique,
    critique_manual: true,
    critique_note: 'Critica deve ser preenchida pela agente Nutri.'
  };

  writeReport(outPath, payload, mdPath);
  console.log(`Nutri report saved: ${outPath}`);
  console.log(`Nutri report saved: ${mdPath}`);
}

main();
