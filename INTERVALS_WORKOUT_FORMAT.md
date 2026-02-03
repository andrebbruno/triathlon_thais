# Intervals.icu Workout Builder - Formato de Descricao

Este documento descreve o formato correto para criar workouts via API que sejam corretamente interpretados pelo Workout Builder do Intervals.icu.

**Ultima atualizacao:** 2026-01-24

---

## API Bulk - Regras de Envio

1. **O endpoint `/events/bulk` exige JSON em formato de array**, mesmo para 1 treino. Exemplo correto:

```json
[
  {
    "external_id": "2026-01-28-run-teste-tiros-30s",
    "category": "WORKOUT",
    "start_date_local": "2026-01-28T00:00:00",
    "type": "Run",
    "name": "Run - Teste Tiros 30s",
    "description": "Warmup\n- 10m 06:45/km pace\n\nTiros 8x\n- 30s Z4 pace\n- 1m20s Z1 pace"
  }
]
```

Se enviar um objeto sozinho (sem `[]`), o servidor retorna **JSON parse error**.

---

## NATACAO (Swim) - REGRAS ESPECIAIS

### Regras Obrigatorias

1. **Sempre usar `meters`** (200meters, NAO 200m) - "m" sozinho e interpretado como minutos!
2. **Sempre terminar com `pace`** para cada step de nado
3. **Usar `Rest Xs` ou `Rest Xm`** para descanso entre series
4. **Usar titulos de secao**: Warmup, Serie Principal Repeat Nx, Cooldown

### Formato Correto

```
Warmup
- 200meters Z1 pace

Serie Principal Repeat 4x
- 100meters Z2 pace
- Rest 10s

Cooldown
- 50meters Z1 pace
```

**Resultado esperado:** Duracao 23m, Distancia 850m, Carga 16, Intensidade 64%

### Formato ERRADO (nao funciona!)

```
- Aquecimento 10m Z1 pace      <- 10m = 10 MINUTOS!
- Series 4x
  - 100m Z2 pace
  - 50m Z1 pace
- Volta calma 5m Z1 pace
```

**Resultado errado:** Duracao 2h45m, Distancia 6.3km (absurdo!)

### Repeticoes

Para series com repeticao:
```
Serie Principal Repeat 4x
- 100meters Z2 pace
- Rest 10s
- 50meters Z1 pace
```

Ou formato alternativo:
```
Serie Principal Repeat 4x
- 100meters Z2 pace
- Rest 10s
```

### Materiais/Educativos (para uso futuro)

Para treinos com materiais (pullbuoy, palmar, prancha, etc):
```
Warmup
- 200meters Z1 pace

Educativo com Pullbuoy Repeat 4x
- 50meters Z2 pace
- Rest 15s

Serie Principal Repeat 6x
- 100meters Z3 pace
- Rest 20s

Cooldown
- 100meters Z1 pace
```

---

## CICLISMO (Ride)

### Formato com Zonas

```
- Aquecimento
  15m Z1

- Endurance
  30m Z2

- Volta a calma
  10m Z1
```

### Formato com Watts

```
- Aquecimento
  15m 100-120w

- Sweet Spot
  20m 88-94% FTP

- Volta a calma
  10m <100w
```

### Intervalos

```
- Aquecimento 15m Z1
- VO2max 5x
  - 3m 105-120% FTP
  - 3m Z1
- Volta a calma 10m Z1
```

---

## CORRIDA (Run)

### Formato com Zonas

```
- Aquecimento 10m Z1 pace
- Base aerobia 25m Z2 pace
- Volta a calma 5m Z1 pace
```

### Formato com Pace

```
- Aquecimento 10m 6:30-7:00/km pace
- Tempo 20m 5:00-5:15/km pace
- Volta a calma 5m Z1 pace
```

### Intervalos/Tiros

```
- Aquecimento 10m Z1 pace
- Tiros 8x
  - 400m Z4 pace
  - 200m Z1 pace
- Volta a calma 5m Z1 pace
```

---

## MUSCULACAO (WeightTraining)

Formato simples:
```
- Aquecimento 10m
- Treino principal 40m
- Alongamento 10m
```

---

## Tabela de Referencia

### Unidades

| Modalidade | Distancia | Tempo |
|------------|-----------|-------|
| Natacao | `100meters`, `200meters` | NAO usar tempo! |
| Ciclismo | `10km` | `15m`, `1h`, `1h30m` |
| Corrida | `400m`, `1km` | `10m`, `30s` |

### Zonas

| Zona | Descricao | FC aproximada |
|------|-----------|---------------|
| Z1 | Recovery | <60% FCmax |
| Z2 | Endurance | 60-70% FCmax |
| Z3 | Tempo | 70-80% FCmax |
| Z4 | Threshold | 80-90% FCmax |
| Z5 | VO2max | 90-95% FCmax |
| Z6 | Anaerobio | >95% FCmax |

---

## Exemplos JSON para Script

### Natacao com Repeticao

```json
{
  "date": "2026-01-28",
  "sport": "swim",
  "name": "Natacao - Aerobio com Series",
  "steps": [
    { "duration": "200", "zone": "Z1", "type": "warmup", "description": "Warmup" },
    { "duration": "4x100", "zone": "Z2", "rest": "15s", "description": "Serie Principal" },
    { "duration": "50", "zone": "Z1", "type": "cooldown", "description": "Cooldown" }
  ]
}
```

### Bike Sweet Spot

```json
{
  "date": "2026-01-29",
  "sport": "ride",
  "name": "Bike - Sweet Spot",
  "steps": [
    { "duration": "15m", "zone": "Z1", "description": "Aquecimento" },
    { "duration": "20m", "zone": "88-94% FTP", "description": "Sweet Spot" },
    { "duration": "5m", "zone": "Z1", "description": "Recuperacao" },
    { "duration": "20m", "zone": "88-94% FTP", "description": "Sweet Spot" },
    { "duration": "10m", "zone": "Z1", "description": "Volta a calma" }
  ]
}
```

### Corrida Base

```json
{
  "date": "2026-01-30",
  "sport": "run",
  "name": "Run - Base Z2",
  "steps": [
    { "duration": "10m", "zone": "Z1", "description": "Aquecimento" },
    { "duration": "30m", "zone": "Z2", "description": "Base aerobia" },
    { "duration": "5m", "zone": "Z1", "description": "Volta a calma" }
  ]
}
```

---

## Fonte

- Testes praticos com Intervals.icu API + Workout Builder
- Screenshots de comparacao (natacaoERRADO.png vs natacaoCERTO.png)
- Validado em 2026-01-24
