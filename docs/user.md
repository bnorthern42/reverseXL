# ReverseXL User Guide

Welcome to the **ReverseXL** user guide. This document explains how to run ReverseXL, reverse-engineer implicit schemas and dependency graphs from Excel workbooks, and interpret the generated reports and visualization artifacts.

---

## 📋 Table of Contents
1. [Overview](#-overview)
2. [Prerequisites & Installation](#-prerequisites--installation)
3. [One-Click Automated Extraction](#-one-click-automated-extraction)
4. [Manual Pipeline Execution](#-manual-pipeline-execution)
5. [Understanding Output Artifacts](#-understanding-output-artifacts)
   - [Machine-Readable Schema (`schema-report.json`)](#1-machine-readable-schema-schema-reportjson)
   - [Executive Tabular PDF (`schema-report.pdf`)](#2-executive-tabular-pdf-schema-reportpdf)
   - [Visual Computational DAG (`call_graph.pdf`)](#3-visual-computational-dag-call_graphpdf)
6. [Interpreting Cell Roles & Heuristics](#-interpreting-cell-roles--heuristics)
7. [Troubleshooting & FAQs](#-troubleshooting--faqs)

---

## 🌟 Overview

Engineering and financial spreadsheets often function as critical computational software systems, yet they lack explicit schemas, data dictionaries, and architectural diagrams. Important input parameters, internal multi-stage calculation formulas, and final business outputs are mixed together in complex 2D grids.

**ReverseXL** automatically analyzes the workbook structure to identify:
- **`USER_INPUT`**: Configurable boundary parameters that drive the model (e.g., constants, validated cells, unlocked parameters).
- **`INTERMEDIATE`**: Internal calculation DAG nodes (formulas with both upstream inputs and downstream consumers).
- **`OUTPUT`**: Key calculation endpoints, leaf formulas, and exported summary metrics.

---

## 🛠️ Prerequisites & Installation

ReverseXL requires Java 17 and lightweight system utilities for generating PDF and visual graph artifacts.

### Debian / Ubuntu
```bash
sudo apt update
sudo apt install -y openjdk-17-jdk maven jq pandoc texlive-xetex fonts-dejavu graphviz
```

### Arch Linux / Manjaro
```bash
sudo pacman -S jdk17-openjdk maven jq pandoc texlive-core ttf-dejavu graphviz
```

### macOS (Homebrew)
```bash
brew install openjdk@17 maven jq pandoc graphviz
```

---

## ⚡ One-Click Automated Extraction

The recommended way to run ReverseXL is using the master pipeline script [`extract.sh`](file:///home/bradn44/Projects/current/excelExtractor/extract.sh):

```bash
# Analyze default engineering workbook (engineering_model.xlsx)
./extract.sh

# Analyze your custom workbook (.xlsx, .xlsm, .xls)
./extract.sh /path/to/my_model.xlsx
```

### What `extract.sh` Does:
1. **Idempotent Build**: Verifies that the shaded JAR exists in `target/` and automatically compiles with `mvn clean package -DskipTests` if needed.
2. **Heuristic Engine Execution**: Runs the Java POI analysis engine on the target workbook and outputs `schema-report.json`.
3. **Tabular Summary Generation**: Executes [`generate_pdf_report.sh`](file:///home/bradn44/Projects/current/excelExtractor/generate_pdf_report.sh) to produce `schema-report.pdf`.
4. **Visual Graph Compilation**: Executes [`generate_call_graph.py`](file:///home/bradn44/Projects/current/excelExtractor/generate_call_graph.py) to produce `call_graph.pdf`.

---

## 🔬 Manual Pipeline Execution

If you prefer to run specific steps individually:

### 1. Build Executable JAR
```bash
mvn clean package -DskipTests
```

### 2. Extract Schema to JSON
```bash
java -jar target/excel-schema-extractor-1.0.0-SNAPSHOT.jar path/to/workbook.xlsx
```

### 3. Generate Tabular PDF
```bash
./generate_pdf_report.sh
```

### 4. Generate Visual Call Graph PDF
```bash
./generate_call_graph.py
```

---

## 📄 Understanding Output Artifacts

Running the pipeline produces three primary artifacts in your working directory:

### 1. Machine-Readable Schema (`schema-report.json`)
A complete JSON array of all classified cells with confidence percentages, detailed evidence lists, and direct upstream references (`dependsOn`):

```json
[
  {
    "sheetName": "Model",
    "cellReference": "D2",
    "role": "USER_INPUT",
    "confidence": 50,
    "evidence": [
      "Constant value referenced by formulas",
      "Has adjacent label: 'σ'"
    ],
    "dependsOn": []
  },
  {
    "sheetName": "Model",
    "cellReference": "F2",
    "role": "INTERMEDIATE",
    "confidence": 90,
    "evidence": [
      "Formula with both upstream and downstream dependencies"
    ],
    "dependsOn": [
      "Model!D2",
      "Model!E2"
    ]
  },
  {
    "sheetName": "Model",
    "cellReference": "R3",
    "role": "OUTPUT",
    "confidence": 60,
    "evidence": [
      "Has adjacent label: 'Efficiency'",
      "Formula with zero dependents (Leaf node)"
    ],
    "dependsOn": [
      "Model!Q3"
    ]
  }
]
```

### 2. Executive Tabular PDF (`schema-report.pdf`)
A condensed, 2-to-4 page document formatted with Markdown and Pandoc:
- **Executive Summary**: Total cell counts by role.
- **`USER_INPUT` Table**: Lists every identified input parameter, its sheet coordinate, confidence score, and adjacent label context.
- **`OUTPUT` Table**: Lists every identified output calculation, unit indicator, and classification evidence.
- **Noise Reduction**: Internal `INTERMEDIATE` calculations are counted in the summary but omitted from detailed tables for readability.

### 3. Visual Computational DAG (`call_graph.pdf`)
A vector PDF diagram generated via Graphviz `dot`:
- **Green Rectangles (`USER_INPUT`)**: Starting input parameters (root nodes).
- **Grey Ellipses (`INTERMEDIATE`)**: Multi-tiered formula calculations.
- **Blue Double Octagons (`OUTPUT`)**: Final result leaf nodes and KPI outputs.
- **Directed Arrows**: Trace data flow from inputs $\to$ intermediate steps $\to$ outputs.

---

## 🎯 Interpreting Cell Roles & Heuristics

| Role | Meaning | Typical Heuristic Signals |
| :--- | :--- | :--- |
| **`USER_INPUT`** | Independent variables and user-configurable parameters | • Constant cell referenced by $\ge 1$ formula<br>• Unlocked cell in a protected worksheet<br>• Cell with Data Validation drop-downs / bounds<br>• Cell reset by VBA `Range().ClearContents` macros<br>• Adjacent descriptive labels to the left or above |
| **`INTERMEDIATE`** | Step-by-step formula logic in the computational DAG | • Formula cell with both upstream dependencies and downstream consumers (90% confidence) |
| **`OUTPUT`** | End-of-line model results and business conclusions | • Formula cell with zero downstream dependents (Leaf node)<br>• Formula on `Summary` / `Result` worksheets<br>• Adjacent engineering unit strings (`kg`, `m`, `USD`, `%`)<br>• Cells referenced by VBA export routines or chart series |

---

## ❓ Troubleshooting & FAQs

### Q: Why did the PDF generation fail with a font warning?
**A:** If your workbook contains mathematical or Greek symbols (`σ`, `µ`, `ε`, `φ`, `ω`, `γ`, `β`), ensure `texlive-xetex` and `fonts-dejavu` (or `ttf-dejavu`) are installed on your machine.

### Q: Can ReverseXL analyze `.xlsm` macro workbooks without executing VBA code?
**A:** Yes. ReverseXL uses POI's `VBAMacroReader` to perform safe static text analysis on the VBA bytecode stream, detecting `.ClearContents` routines without executing untrusted code.

### Q: How can I integrate `schema-report.json` into a CI/CD pipeline?
**A:** You can parse `schema-report.json` using `jq` or Python to enforce schema validation rules, prevent breaking changes in workbook calculations, or automatically generate API schemas for spreadsheet-backed microservices.
