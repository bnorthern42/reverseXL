#!/bin/bash
set -e

# ==============================================================================
# Master Excel Schema Extraction & Reporting Pipeline
# ==============================================================================

TARGET_FILE="${1:-engineering_model.xlsx}"

if [ ! -f "$TARGET_FILE" ]; then
    echo "[ERROR] Target Excel workbook not found: '$TARGET_FILE'" >&2
    exit 1
fi

JAR_FILE="target/excel-schema-extractor-1.0.0-SNAPSHOT.jar"

# 1. Maven Incremental Package
echo "==> Building / updating JAR package with Maven..."
mvn package -DskipTests

# 2. Step 1: Java Heuristic Engine & Schema Extraction
echo "[1/4] Running Java Heuristic Engine..."
java -jar "$JAR_FILE" "$TARGET_FILE"

if [ ! -f "schema-report.json" ]; then
    echo "[ERROR] Extraction failed: 'schema-report.json' was not created." >&2
    exit 1
fi

# 3. Ensure Helper Scripts are Executable
chmod +x generate_pdf_report.sh generate_call_graph.py

# 4. Step 2: Tabular Summary PDF Generation
echo "[2/4] Generating Tabular PDF Report..."
./generate_pdf_report.sh

# 5. Step 3: Visual Graphviz Call Graph Generation
echo "[3/4] Generating Visual Call Graph..."
./generate_call_graph.py

# 6. Step 4: Completion Confirmation
echo "[4/4] Pipeline Complete! Outputs: schema-report.pdf, call_graph.pdf"
