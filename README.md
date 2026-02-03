# Triathlon Thais

Relatórios e planejamento semanal da atleta Thais Lourenço, com exportação automática do Intervals.icu e site estático em `docs/`.

## Como rodar localmente

### 1) Exportar semana (Intervals)
```powershell
$env:INTERVALS_API_KEY = "<sua_api_key>"
$env:INTERVALS_ATHLETE_ID = "i446982"
.\export-intervals-week-com-notas.ps1
```

### 2) Gerar relatório semanal (HTML)
```powershell
.\scripts\build_site.ps1
```

Abrir: `docs\index.html`

## GitHub Actions

O workflow `weekly-export` roda todo domingo (14h BRT / 17h UTC) e:
- Exporta Intervals
- Gera o site em `docs/`
- Faz commit automático

### Secrets necessários
Configure em **Settings → Secrets and variables → Actions**:
- `INTERVALS_API_KEY` (Intervals.icu)
- `INTERVALS_ATHLETE_ID` (ex.: `i446982`)

Opcional (se quiser e-mail):
- `SMTP_SERVER`, `SMTP_PORT`, `SMTP_USERNAME`, `SMTP_PASSWORD`, `SMTP_TO`, `SMTP_FROM`

## GitHub Pages

O site é publicado via GitHub Actions (workflow `deploy-pages`).
Link esperado:
```
https://andrebbruno.github.io/triathlon_thais/
```

## Estrutura
- `Relatorios_Intervals/`: JSONs semanais + planejados
- `docs/`: site estático (página principal e relatórios)
- `scripts/`: build do site e utilitários

---
Se quiser ajustar layout, KPIs ou regras de treino, edite `COACHING_MEMORY.md` e rode `build_site.ps1`.
