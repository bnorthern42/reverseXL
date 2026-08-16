#!/bin/bash
set -e

# ==============================================================================
# Excel Schema Extraction PDF Report Generator (Tabular Summary)
# ==============================================================================

TEMP_MD="temp_report.md"
trap 'rm -f "$TEMP_MD"' EXIT INT TERM

# 1. Dependency Checks
if ! command -v jq &> /dev/null; then
    echo "[ERROR] 'jq' is not installed. Please install jq (e.g., 'sudo apt install jq' or 'sudo pacman -S jq')." >&2
    exit 1
fi

if ! command -v pandoc &> /dev/null; then
    echo "[ERROR] 'pandoc' is not installed. Please install pandoc (e.g., 'sudo apt install pandoc' or 'sudo pacman -S pandoc')." >&2
    exit 1
fi

# 2. File Check
JSON_INPUT="schema-report.json"
PDF_OUTPUT="schema-report.pdf"

if [ ! -f "$JSON_INPUT" ]; then
    echo "[ERROR] Input file '$JSON_INPUT' not found in current directory." >&2
    echo "Please run the extractor first: java -jar target/excel-schema-extractor-1.0.0-SNAPSHOT.jar <workbook.xlsx>" >&2
    exit 1
fi

echo "==> Generating summary PDF report from '$JSON_INPUT'..."

# 3. Aggregate Counts
TOTAL_NODES=$(jq '. | length' "$JSON_INPUT")
INPUT_NODES=$(jq '[.[] | select(.role == "USER_INPUT")] | length' "$JSON_INPUT")
OUTPUT_NODES=$(jq '[.[] | select(.role == "OUTPUT")] | length' "$JSON_INPUT")
INTERMEDIATE_NODES=$(jq '[.[] | select(.role == "INTERMEDIATE")] | length' "$JSON_INPUT")
GEN_DATE=$(date "+%Y-%m-%d %H:%M:%S")

# 4. Construct Markdown Document
cat <<EOF > "$TEMP_MD"
# Excel Schema Extraction Report

**Generated On:** $GEN_DATE  
**Source Artifact:** \`$JSON_INPUT\`  

---

## Executive Summary

- **Total Classified Nodes:** $TOTAL_NODES
- **Identified User Inputs (\`USER_INPUT\`):** $INPUT_NODES
- **Identified Outputs (\`OUTPUT\`):** $OUTPUT_NODES
- **Intermediate Calculation Nodes (\`INTERMEDIATE\`):** $INTERMEDIATE_NODES

> **Note:** Internal formula DAG calculations ($INTERMEDIATE_NODES intermediate nodes) are omitted from detailed tables below to highlight key boundary parameters and primary outputs.

---

## Identified User Inputs (\`USER_INPUT\`)

| Sheet | Cell | Confidence | Evidence |
|---|---|---|---|
EOF

# Append USER_INPUT rows
jq -r '.[] | select(.role == "USER_INPUT") | "| \(.sheetName) | \(.cellReference) | \(.confidence)% | \(.evidence | join("; ")) |"' "$JSON_INPUT" >> "$TEMP_MD"

cat <<EOF >> "$TEMP_MD"

---

## Identified Key Outputs (\`OUTPUT\`)

| Sheet | Cell | Confidence | Evidence |
|---|---|---|---|
EOF

# Append OUTPUT rows
jq -r '.[] | select(.role == "OUTPUT") | "| \(.sheetName) | \(.cellReference) | \(.confidence)% | \(.evidence | join("; ")) |"' "$JSON_INPUT" >> "$TEMP_MD"

# 5. Compile Markdown to PDF using Pandoc
PANDOC_ARGS=("$TEMP_MD" -f markdown-yaml_metadata_block -o "$PDF_OUTPUT" -V geometry:margin=1in)

if command -v xelatex &> /dev/null; then
    PANDOC_ARGS+=(--pdf-engine=xelatex -V mainfont="DejaVu Sans")
fi

pandoc "${PANDOC_ARGS[@]}"

echo "[SUCCESS] PDF report successfully compiled: $PDF_OUTPUT"
