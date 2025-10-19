
# Finanz‑Agent (Einnahmen & Ausgaben)

Ziel: Ein-/Ausgaben verarbeiten, automatisch kategorisieren und monatlich auswerten.
- Auto‑Kategorisierung über `rules.csv`
- Dashboard/Export in `Budget.xlsx`
- Monats‑Reports in `reports/`
- GitHub Actions Workflow: `.github/workflows/run-agent.yml`

## Struktur
Siehe Repo-Baum. Lege Bank‑CSV‑Exporte in `data/inbox/` ab.

## Start (lokal)
rm -f Budget.xlsx

python finanzen_agent.py

python notify.py daily --send
