#!/usr/bin/env python3
"""
Excel Schema Extractor - Visual DAG Call Graph Generator
Groups repeated formula series by Sheet, Column, and normalized formula signature
to display the complete logical calculation flow without repetitive row clutter.
"""

import json
import os
import re
import shutil
import subprocess
import sys

def main():
    # 1. Dependency Check
    if not shutil.which("dot"):
        sys.stderr.write("[ERROR] 'dot' (Graphviz) is not installed.\n")
        sys.stderr.write("Please install Graphviz (e.g., 'sudo apt install graphviz' or 'sudo pacman -S graphviz').\n")
        sys.exit(1)

    json_file = "schema-report.json"
    dot_file = "call_graph.dot"
    pdf_file = "call_graph.pdf"

    # 2. Input File Check
    if not os.path.isfile(json_file):
        sys.stderr.write(f"[ERROR] '{json_file}' not found in the current directory.\n")
        sys.stderr.write("Run the schema extractor first: java -jar target/excel-schema-extractor-1.0.0-SNAPSHOT.jar <workbook.xlsx>\n")
        sys.exit(1)

    print(f"==> Loading '{json_file}' and collapsing formula series by column...")

    with open(json_file, "r", encoding="utf-8") as f:
        nodes_data = json.load(f)

    # 3. Group Nodes by (Sheet, Column, Normalized Formula)
    role_priority = {
        "OUTPUT": 3,
        "USER_INPUT": 2,
        "INTERMEDIATE": 1,
        "UNKNOWN": 0
    }

    node_to_group = {}
    grouped_nodes = {}

    for cell in nodes_data:
        sheet = cell.get("sheetName", "")
        cell_ref = cell.get("cellReference", "")
        role = cell.get("role", "UNKNOWN")
        raw_formula = cell.get("formula", "")
        node_id = f"{sheet}!{cell_ref}"

        # Extract column letters
        col_match = re.match(r"^([A-Z]+)\d+$", cell_ref)
        col = col_match.group(1) if col_match else "Col"

        # Create row-agnostic normalized formula signature (e.g. SUM(A1:A10) -> SUM(A:A))
        norm_formula = re.sub(r"\d+", "", raw_formula)

        if raw_formula:
            group_key = f"{sheet}!Col_{col}|{norm_formula}"
        else:
            # Constants / User inputs remain individual nodes
            group_key = f"{sheet}!{cell_ref}"

        node_to_group[node_id] = group_key

        if group_key not in grouped_nodes:
            grouped_nodes[group_key] = {
                "sheet": sheet,
                "col": col,
                "cell_ref": cell_ref,
                "role": role,
                "norm_formula": norm_formula,
                "is_formula": bool(raw_formula),
                "count": 1
            }
        else:
            grouped_nodes[group_key]["count"] += 1
            # Promote role if higher priority (e.g. OUTPUT > INTERMEDIATE)
            current_role = grouped_nodes[group_key]["role"]
            if role_priority.get(role, 0) > role_priority.get(current_role, 0):
                grouped_nodes[group_key]["role"] = role

    # 4. Map Dependency Edges Between Logical Groups
    grouped_edges = set()

    for cell in nodes_data:
        sheet = cell.get("sheetName", "")
        cell_ref = cell.get("cellReference", "")
        node_id = f"{sheet}!{cell_ref}"
        start_group = node_to_group.get(node_id)

        if not start_group:
            continue

        for dependency in cell.get("dependsOn", []):
            end_group = node_to_group.get(dependency)
            # Data flows from source dependency TO the calculation group
            if end_group and end_group != start_group:
                grouped_edges.add((end_group, start_group))

    print(f"==> Collapsed {len(nodes_data)} raw cells into {len(grouped_nodes)} logical series nodes with {len(grouped_edges)} distinct logical connections")

    # 5. Build Graphviz DOT Definition
    dot_lines = [
        "digraph ExcelGroupedCallGraph {",
        '    layout="dot";',
        '    rankdir="LR";',
        '    nodesep="0.4";',
        '    ranksep="1.0";',
        '    splines="polyline";',
        '    node [fontname="Helvetica", fontsize=10];',
        '    edge [fontname="Helvetica", fontsize=8, color="#555555"];',
        ""
    ]

    role_styles = {
        "USER_INPUT": 'shape="box", style="filled", fillcolor="lightgreen"',
        "OUTPUT": 'shape="doubleoctagon", style="filled", fillcolor="lightblue"',
        "INTERMEDIATE": 'shape="ellipse", style="filled", fillcolor="lightgrey"'
    }

    # Pass 1: Render Grouped Nodes
    for group_key, meta in sorted(grouped_nodes.items()):
        role = meta["role"]
        sheet = meta["sheet"]
        col = meta["col"]
        count = meta["count"]
        norm_formula = meta["norm_formula"]
        is_formula = meta["is_formula"]
        style_attrs = role_styles.get(role, 'shape="ellipse", style="filled", fillcolor="white"')

        if is_formula:
            truncated_formula = (norm_formula[:20] + "...") if len(norm_formula) > 20 else norm_formula
            # Escape quotes/newlines for Graphviz
            clean_formula = truncated_formula.replace('"', '\\"')
            if count > 1:
                label = f"{sheet}!Col {col} ({count} rows)\\n{clean_formula}\\n({role})"
            else:
                label = f"{sheet}!{meta['cell_ref']}\\n{clean_formula}\\n({role})"
        else:
            label = f"{sheet}!{meta['cell_ref']}\\n({role})"

        dot_lines.append(f'    "{group_key}" [label="{label}", {style_attrs}];')

    dot_lines.append("")

    # Pass 2: Render Logical Directed Edges
    for source_group, target_group in sorted(grouped_edges):
        dot_lines.append(f'    "{source_group}" -> "{target_group}";')

    dot_lines.append("}")
    dot_content = "\n".join(dot_lines)

    # 6. Write DOT file, Render PDF, and Clean Up
    try:
        with open(dot_file, "w", encoding="utf-8") as f:
            f.write(dot_content)

        print(f"==> Compiling '{dot_file}' to '{pdf_file}' using Graphviz 'dot'...")
        subprocess.run(["dot", "-Tpdf", dot_file, "-o", pdf_file], check=True)
        print(f"[SUCCESS] Logical DAG call graph successfully generated: {pdf_file}")
    except subprocess.CalledProcessError as e:
        sys.stderr.write(f"[ERROR] Graphviz rendering failed with code {e.returncode}.\n")
        sys.exit(1)
    finally:
        if os.path.exists(dot_file):
            os.remove(dot_file)

if __name__ == "__main__":
    main()
