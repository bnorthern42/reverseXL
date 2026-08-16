# ReverseXL: Excel Schema & Dependency Extractor

[![Java](https://img.shields.io/badge/Java-17%2B-blue.svg)](https://openjdk.org/)
[![Apache POI](https://img.shields.io/badge/Apache%20POI-5.2.4-orange.svg)](https://poi.apache.org/)
[![License](https://img.shields.io/badge/License-Apache%202.0-green.svg)](LICENSE)

A high-assurance, dependency-minimized Java engine and reporting pipeline designed to **reverse-engineer implicit data schemas and computational DAGs** from complex Excel workbooks (`.xlsx`, `.xlsm`, `.xls`). 

By fusing structural cell metadata, physical spatial layout cues, dynamic style clustering, regex-based formula DAG reconstruction, and static VBA macro analysis, ReverseXL automatically classifies cells into functional roles: **`USER_INPUT`**, **`OUTPUT`**, and **`INTERMEDIATE`** calculations.

---

## 📚 Documentation Index

- **[User Guide (`docs/user.md`)](file:///home/bradn44/Projects/current/excelExtractor/docs/user.md)**: End-to-end user workflows, CLI options, interpretation of tabular PDF summaries and visual DAG graphs, and troubleshooting.
- **[Developer Guide (`docs/dev.md`)](file:///home/bradn44/Projects/current/excelExtractor/docs/dev.md)**: Deep dive into the heuristic scoring matrix, two-pass analysis engine, POI adapters, DAG algorithms, and extending rules.

---

## ⚡ Quickstart: Automated End-to-End Pipeline

ReverseXL includes a master orchestration script [`extract.sh`](file:///home/bradn44/Projects/current/excelExtractor/extract.sh) that manages idempotent compilation, heuristic extraction, tabular PDF summarization, and visual Graphviz DAG rendering in a single command:

```bash
# Run against default model (engineering_model.xlsx)
./extract.sh

# Run against a custom workbook
./extract.sh /path/to/my_financial_model.xlsm
```

### Generated Output Artifacts

| Artifact | Format | Description |
| :--- | :--- | :--- |
| **`schema-report.json`** | JSON | Machine-readable schema array containing cell addresses, roles, confidence scores, evidence trails, and directed dependency edges (`dependsOn`). |
| **`schema-report.pdf`** | PDF | High-density tabular executive report summarizing key inputs and outputs while omitting intermediate DAG calculation noise. |
| **`call_graph.pdf`** | PDF | Color-coded Graphviz vector DAG tracing data flow from green user input parameters through intermediate calculation layers to blue output targets. |

---

## 🔒 Security & Supply Chain Posture

Modern spreadsheet parsing utilities written in Node.js/Python frequently rely on deep, fragmented dependency trees with dozens of transitive packages—introducing substantial risk of software supply chain poisoning, prototype pollution, and side-channel vulnerabilities.

**ReverseXL enforces a hardened, monolithic dependency posture:**
- **Zero Deep Dependency Sprawl:** Built exclusively on Java 17 Standard Edition, the battle-tested **Apache POI 5.2.4** monolith, and Google **Gson 2.10.1**.
- **Hermetic Supply Chain:** Direct, pinned dependencies eliminate transitive dependency injection and upstream drift.
- **Static Macro Inspection:** Embedded VBA macros are analyzed strictly as static string tokens via POI's `VBAMacroReader` without executing bytecode or script runtimes.
- **Framework-Independent TDD:** Built-in lightweight test runner ([`StringTddRunner`](file:///home/bradn44/Projects/current/excelExtractor/src/test/java/org/reversexl/StringTddRunner.java)) ensures hermetic test verification without JUnit/TestNG dependency surface.

---

## 🧠 Evaluated Heuristic Engine Signals

ReverseXL scores cells across multiple physical and logical dimensions:

| Heuristic Dimension | Evaluated Signals | Role Influence |
| :--- | :--- | :--- |
| **Spatial Layout** | • Adjacent text labels (1–2 cells left or 1 cell above)<br>• Adjacent engineering unit strings (1 cell right: `kg`, `Hz`, `USD`, `m/s`, etc.) | `USER_INPUT` (+2)<br>`OUTPUT` (+2) |
| **Dependency Graph** | • Downstream dependents count (constants referenced by downstream formulas)<br>• Upstream dependencies count (references inside formula)<br>• Formulas with both upstream and downstream links | `USER_INPUT` (+3 to +4)<br>`INTERMEDIATE` (90% Conf)<br>`OUTPUT` Leaf (+4) |
| **Structural Protection** | • Unlocked cells inside password-protected sheets<br>• Cell protection flags (`getLocked()`) | `USER_INPUT` (+5) |
| **Data Validation** | • Intersecting sheet `DataValidation` bounding boxes | `USER_INPUT` (+4) |
| **Named Ranges** | • Explicit named ranges mapped to cell addresses | `USER_INPUT` (+2)<br>`OUTPUT` (+2) |
| **Style Clustering** | • First-pass frequency analysis of constant vs formula cell styles<br>• Dynamic dominant input/output style detection with default style (index 0) and collision suppression | `USER_INPUT` (+3)<br>`OUTPUT` (+3) |
| **VBA Macro Sweeping** | • Static analysis of `Range("...").ClearContents` routines via `VBAMacroReader` | `USER_INPUT` (+4) |

---

## 🏗️ Architecture Overview

```
org.reversexl
├── engine
│   ├── CellContext.java          # Core interface abstracting cell attributes and heuristic signals
│   ├── SimpleCellContext.java    # Fluent POJO implementation & builder for unit tests
│   ├── PoiCellContext.java       # Apache POI adapter bridging real Cell objects to CellContext
│   └── CellScorer.java           # Point-based heuristic classification engine
├── parser
│   └── WorkbookParser.java       # High-level POI analyzer & graph extractor
├── Main.java                     # CLI entry point, 2-pass analyzer, and JSON report generator
└── [test]
    ├── StringTddRunner.java      # Lightweight string-based assertion runner
    ├── CellScorerTest.java       # Heuristic unit test suite
    ├── PoiCellContextTest.java   # POI context adapter test suite
    └── WorkbookParserTest.java   # End-to-end integration test suite
```

---

## 🚀 Manual Build & Execution

### Prerequisites
- **Java Development Kit (JDK) 17+**
- **Apache Maven 3.8+**
- **`jq`**, **`pandoc`**, **`graphviz`** *(for PDF report and call graph scripts)*

### 1. Compile & Package
Build the self-contained executable shaded JAR:
```bash
mvn clean package -DskipTests
```

### 2. Run Java Extractor CLI Directly
```bash
# Analyze target workbook and generate schema-report.json
java -jar target/excel-schema-extractor-1.0.0-SNAPSHOT.jar path/to/model.xlsx
```

### 3. Generate Reports Independently
```bash
# Generate concise tabular PDF summary
./generate_pdf_report.sh

# Generate visual Graphviz DAG call graph
./generate_call_graph.py
```

### 4. Run Custom TDD Test Suites
Execute the embedded zero-dependency test runner suites:
```bash
mvn clean test-compile
java -cp target/excel-schema-extractor-1.0.0-SNAPSHOT.jar:target/test-classes org.reversexl.CellScorerTest
java -cp target/excel-schema-extractor-1.0.0-SNAPSHOT.jar:target/test-classes org.reversexl.PoiCellContextTest
java -cp target/excel-schema-extractor-1.0.0-SNAPSHOT.jar:target/test-classes org.reversexl.WorkbookParserTest
```
