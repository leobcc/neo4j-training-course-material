# Neo4j Training Course — Materiale Didattico

Corso di 5 sessioni su database a grafo con Neo4j, rivolto a consulenti IT.

## Contenuto

| Sessione | Argomento | Slide | Notebook |
|----------|-----------|-------|----------|
| 1 | Jump Start: Grafi & Primo Database | ✅ | ✅ |
| 2 | Cypher: Da Basi a Query Reali | 🚧 | 🚧 |
| 3 | Data Modeling & Import | 🚧 | 🚧 |
| 4 | Graph Data Science | 🚧 | 🚧 |
| 5 | AI & GraphRAG | 🚧 | 🚧 |

## Setup rapido

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

Copia `.env.example` in `.env` e inserisci le credenziali AuraDB.

## Struttura

```
slides/            — Slide deck (PPTX)
notebooks/         — Jupyter Notebook
data/              — CSV anonimizzati
project/           — Guida e setup del progetto
scripts/           — Utility Python
```

## Prerequisiti

- Python 3.10+
- Conto Neo4j AuraDB gratuito (console.neo4j.io)
- Git
