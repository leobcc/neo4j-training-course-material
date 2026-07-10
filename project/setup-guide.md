# Neo4j Setup Guide

Choose one of the following options to get Neo4j running.

## Option A: Neo4j AuraDB Free (Recommended)

1. Go to https://console.neo4j.io/ and sign in
2. Create a new instance — select **AuraDB Free** tier
3. Once created, copy the **connection URI** (looks like `neo4j+s://<id>.databases.neo4j.io`)
4. Save the **password** you set during creation

**Limits:** 200K nodes, 400K relationships — more than enough for our course (~300 nodes, ~7K relationships).

> ⚠️ AuraDB Free pauses after 72h of inactivity and is deleted after 30 days paused. Reactivate it from the console before each session.

## Option B: Docker (Local)

```bash
docker run \
  --name neo4j-course \
  -p 7474:7474 -p 7687:7687 \
  -e NEO4J_AUTH=neo4j/password123 \
  -e NEO4J_PLUGINS='["apoc", "graph-data-science"]' \
  neo4j:2025-latest
```

Open http://localhost:7474. Connection: `neo4j://localhost:7687` / `neo4j` / `password123`

## Configure Connection

Copy the env template and fill in your AuraDB details:

```bash
copy .env.example .env
```

Then edit `.env` with your actual URI and password. The notebook loads these automatically.

## Data Source

Single CSV: `data/org_graph.csv` (279 rows, 68 columns). Generated from the original Excel by:

```bash
python scripts/transform_data.py
```

## Load the Org Graph Data

Open and run the notebook:

```bash
.venv\Scripts\jupyter notebook notebooks\session-01\session-01.ipynb
```

The notebook loads the CSV into Neo4j via Python driver and runs your first queries.
