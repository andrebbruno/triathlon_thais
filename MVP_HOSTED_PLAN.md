# MVP Hosted (Free) - Plan

Goal: keep cost near zero and automate weekly export + web view.

## Stack (free)
- GitHub Actions: weekly cron to run export.
- GitHub Pages: static site with reports.
- Intervals.icu API: data source.

## Files added
- `.github/workflows/weekly-export.yml`
- `scripts/build_site.ps1`

## Setup steps
1. Create a GitHub repo and push this folder.
2. In GitHub > Settings > Secrets and variables > Actions:
   - Add `INTERVALS_API_KEY` with your key.
3. Enable GitHub Pages:
   - Source: `Deploy from a branch`
   - Branch: `main`
   - Folder: `/docs`
4. In Actions, run **weekly-export** once (manual).
5. Check `https://<user>.github.io/<repo>/` for reports list.

## What the workflow does
1. Validates `INTERVALS_API_KEY` secret.
2. Runs `export-intervals-week-com-notas.ps1` and `intervals-longterm-coach-edition.ps1`.
3. Builds static site from `Relatorios_Intervals`.
4. Commits `Relatorios_Intervals` and `docs`.

## Notes
- The export already saves JSON and planned MD in `Relatorios_Intervals`.
- You can change the cron time in the workflow file.

## Roadmap (next improvements)
### Fluxo
1. Checklist de execucao no report (export/longterm/analysis/upload/build).
2. Carimbo de confiabilidade (dados completos/parciais/estimados).
3. Notas do atleta como voto (opinioes influenciam a decisao, nao substituem).
4. Alertas de tendencia curta (ex: TSB baixo + sono ruim).
5. Run log/status salvo em `Relatorios_Intervals`.
6. Email HTML: template moderno + resumo da semana.
7. Link direto para o ultimo report (alem do index).

### Dados e analise semanal
1. Mapeamento de objetivo do treino (definir a intencao: VO2, base, recovery, tecnica).
2. Planejado vs executado por bloco (aderencia por segmento, nao so total).
3. Correlacao com wellness (sono/HRV/FC repouso no dia do treino).

### Nutricao (relatorio separado)
1. Relatorio Nutri separado (nao misturar com report do Coach).
2. MFP semanal (7 dias) + Intervals executado: base para critica do que foi feito.
3. Intervals planejado (semana seguinte): estimar necessidades caloricas.
4. Critica nao automatica: feita pela agente Nutri.
5. Consolidar arquivos em Relatorios_Nutri.

### Longo prazo (tendencias)
1. Curvas CTL/ATL/TSB com faixa alvo para a prova A.
2. Historico de peso e HRV com media movel e correlacao com carga.
3. Progressao por modalidade (volume e intensidade por esporte).
4. Previsao de performance ate a prova A (pace/FTP).
5. Identificar blocos automaticamente (build/deload/peak).

### Relatorio (experiencia)
1. Resumo executivo no topo (3 bullets: bom/ruim/mudar).
2. Comparacao com semana anterior (setas + %).
3. Recomendacoes taticas por modalidade (1 acao imediata).
4. Glossario inteligente com tooltip nos cards.

### Regras iniciais: objetivo do treino (mapeamento)
- Sweet Spot, Limiar, Threshold, Tempo -> objetivo: Threshold/Tempo
- VO2, VO2max, Intervalos curtos, 30s/1m -> objetivo: VO2/Anaerobio
- Z2, Endurance, Base, Longao -> objetivo: Base/Endurance
- Recovery, Regenerativo, Solto -> objetivo: Recuperacao
- Tecnica, Drills, Educativo -> objetivo: Tecnica
- Forca, Strength, Gym -> objetivo: Forca
- Brick, Transicao, T2 -> objetivo: Especifico Triathlon

## API key (local vs GitHub)
### Local MVP
The scripts now look for the key in this order:
1. `INTERVALS_API_KEY` environment variable
2. `api_key.txt` in the project root
3. `%USERPROFILE%\.intervals\api_key.txt` (or `$HOME/.intervals/api_key.txt`)

Recommended: create the local file outside the repo:
`C:\Users\<USER>\.intervals\api_key.txt`

### GitHub Actions
Use `INTERVALS_API_KEY` in GitHub Secrets. The workflow reads the secret directly and does not write any key file.
