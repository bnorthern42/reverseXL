package org.reversexl;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import org.apache.poi.poifs.macros.VBAMacroReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.reversexl.engine.CellScorer;
import org.reversexl.engine.CellScorer.ClassificationResult;
import org.reversexl.engine.CellScorer.Role;
import org.reversexl.engine.PoiCellContext;

import java.io.File;
import java.io.FileWriter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Main entry point for the ReverseXL CLI extraction engine.
 * Executes a two-pass workbook scan, evaluates heuristic signals, builds formula dependency edges,
 * and serializes the resulting schema to a structured JSON report.
 */
public class Main {

    private static final Pattern CELL_REF_PATTERN = Pattern.compile("\\b([A-Z]{1,3}[0-9]{1,7})\\b");
    private static final Pattern CLEAR_CONTENTS_PATTERN = Pattern.compile("(?i)Range\\(\"([A-Z]{1,3}[0-9]{1,7})(:[A-Z]{1,3}[0-9]{1,7})?\"\\)\\.ClearContents");

    /**
     * Private constructor to prevent instantiation of utility CLI class.
     */
    private Main() {}

    /**
     * Data transfer record representing an individual classified cell in the output JSON report.
     *
     * @param sheetName worksheet name
     * @param cellReference cell alphanumeric reference (e.g., C2)
     * @param role string name of the classified {@link Role}
     * @param confidence confidence level percentage (0 to 100)
     * @param evidence rationale and heuristic triggering signals
     * @param dependsOn list of upstream cell references that this cell's formula depends on
     * @param formula raw Excel formula string, or empty string if not a formula
     */
    public record ClassifiedCell(
            String sheetName,
            String cellReference,
            String role,
            int confidence,
            List<String> evidence,
            List<String> dependsOn,
            String formula
    ) {}

    /**
     * CLI entry point.
     *
     * @param args command-line arguments specifying workbook path
     */
    public static void main(String[] args) {
        String filePath = "engineering_model.xlsx";

        if (args.length > 0) {
            if ("--analyze".equalsIgnoreCase(args[0]) && args.length > 1) {
                filePath = args[1];
            } else if (!args[0].startsWith("-")) {
                filePath = args[0];
            }
        }

        File targetFile = new File(filePath);
        if (!targetFile.exists() || !targetFile.isFile()) {
            System.err.println("Error: File not found or not a valid file: " + filePath);
            System.exit(1);
        }

        System.out.println("Processing workbook: " + targetFile.getName());

        // 1. VBA Macro Pre-processing (ClearContents extraction)
        Set<String> clearedCells = new HashSet<>();
        try (VBAMacroReader macroReader = new VBAMacroReader(targetFile)) {
            Map<String, String> macros = macroReader.readMacros();
            for (String macroCode : macros.values()) {
                if (macroCode != null) {
                    Matcher m = CLEAR_CONTENTS_PATTERN.matcher(macroCode);
                    while (m.find()) {
                        String cellRef = m.group(1);
                        if (cellRef != null) {
                            clearedCells.add(cellRef.toUpperCase());
                        }
                    }
                }
            }
        } catch (Exception ignored) {
            // Graceful degradation for non-macro workbooks (.xlsx)
        }

        try (Workbook workbook = WorkbookFactory.create(targetFile)) {
            // 2. First Pass: Style Clustering Analysis & Lightweight Dependency Graph
            Map<Short, Integer> constantStyleCounts = new HashMap<>();
            Map<Short, Integer> formulaStyleCounts = new HashMap<>();
            Map<String, Integer> downstreamCounts = new HashMap<>();

            for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
                Sheet sheet = workbook.getSheetAt(s);
                if (sheet == null) {
                    continue;
                }

                String sheetName = sheet.getSheetName();

                for (Row row : sheet) {
                    if (row == null) {
                        continue;
                    }

                    for (Cell cell : row) {
                        if (cell == null) {
                            continue;
                        }

                        CellType type = cell.getCellType();
                        CellStyle style = cell.getCellStyle();

                        // Style clustering: ignore default style index 0
                        if (style != null && style.getIndex() != 0) {
                            short styleIdx = style.getIndex();
                            if (type == CellType.FORMULA) {
                                formulaStyleCounts.merge(styleIdx, 1, Integer::sum);
                            } else if (type != CellType.BLANK) {
                                if (type == CellType.STRING) {
                                    String str = cell.getStringCellValue();
                                    if (str == null || str.trim().isEmpty()) {
                                        // Skip empty strings
                                    } else {
                                        constantStyleCounts.merge(styleIdx, 1, Integer::sum);
                                    }
                                } else {
                                    constantStyleCounts.merge(styleIdx, 1, Integer::sum);
                                }
                            }
                        }

                        // Lightweight Dependency Graph pre-processing (downstream counts)
                        if (type == CellType.FORMULA) {
                            try {
                                String formula = cell.getCellFormula();
                                if (formula != null) {
                                    Matcher matcher = CELL_REF_PATTERN.matcher(formula);
                                    while (matcher.find()) {
                                        String ref = matcher.group(1);
                                        String key = sheetName + "!" + ref;
                                        downstreamCounts.merge(key, 1, Integer::sum);
                                    }
                                }
                            } catch (Exception ignored) {
                                // Ignore POI evaluation / parse exceptions safely
                            }
                        }
                    }
                }
            }

            short dominantInputStyle = getDominantStyleIndex(constantStyleCounts);
            short dominantOutputStyle = getDominantStyleIndex(formulaStyleCounts);

            // Discard overlapping style indices on collision
            if (dominantInputStyle != -1 && dominantInputStyle == dominantOutputStyle) {
                dominantInputStyle = -1;
                dominantOutputStyle = -1;
            }

            // 3. Pre-process Named Ranges at workbook level to avoid N+1 queries
            Map<String, String> cellRefToNameMap = new HashMap<>();
            try {
                for (Name name : workbook.getAllNames()) {
                    if (name == null || name.isDeleted() || name.isFunctionName()) {
                        continue;
                    }
                    try {
                        String formula = name.getRefersToFormula();
                        if (formula != null && !formula.trim().isEmpty()) {
                            String normalized = formula.trim().replace("$", "").replace("'", "").toUpperCase();
                            cellRefToNameMap.put(normalized, name.getNameName());
                        }
                    } catch (Exception ignored) {
                        // Handle complex range formulas safely
                    }
                }
            } catch (Exception ignored) {
            }

            // 4. Second Pass: Classification Loop with Upstream & Downstream Dependencies
            List<ClassifiedCell> classifiedCells = new ArrayList<>();

            for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
                Sheet sheet = workbook.getSheetAt(s);
                if (sheet == null) {
                    continue;
                }

                String sheetName = sheet.getSheetName();
                String sheetKeyPrefix = sheetName.replace("'", "").toUpperCase() + "!";

                List<? extends DataValidation> validations = Collections.emptyList();
                try {
                    validations = sheet.getDataValidations();
                } catch (Exception ignored) {
                }

                for (Row row : sheet) {
                    if (row == null) {
                        continue;
                    }

                    for (Cell cell : row) {
                        if (cell == null) {
                            continue;
                        }

                        String cellAddress = cell.getAddress().formatAsString();
                        String cellKey = sheetKeyPrefix + cellAddress.toUpperCase();
                        String namedRange = cellRefToNameMap.get(cellKey);
                        if (namedRange == null) {
                            namedRange = cellRefToNameMap.get(cellAddress.toUpperCase());
                        }

                        String graphKey = sheetName + "!" + cellAddress;
                        int downstreamCount = downstreamCounts.getOrDefault(graphKey, 0);

                        // Upstream Dependency & Formula Extraction for formulas
                        List<String> dependsOn = new ArrayList<>();
                        String formulaStr = "";
                        if (cell.getCellType() == CellType.FORMULA) {
                            try {
                                String formula = cell.getCellFormula();
                                if (formula != null) {
                                    formulaStr = formula;
                                    Matcher matcher = CELL_REF_PATTERN.matcher(formula);
                                    while (matcher.find()) {
                                        String ref = matcher.group(1);
                                        dependsOn.add(sheetName + "!" + ref);
                                    }
                                }
                            } catch (Exception ignored) {
                            }
                        }

                        int upstreamCount = dependsOn.size();

                        PoiCellContext ctx = new PoiCellContext(
                                cell, validations, namedRange, dominantInputStyle, dominantOutputStyle,
                                downstreamCount, upstreamCount, clearedCells
                        );
                        ClassificationResult result = CellScorer.classify(ctx);

                        if (result.role != Role.UNKNOWN) {
                            classifiedCells.add(new ClassifiedCell(
                                    sheetName,
                                    cellAddress,
                                    result.role.name(),
                                    result.confidencePercentage,
                                    result.evidence,
                                    dependsOn,
                                    formulaStr
                            ));

                            if (result.role == Role.USER_INPUT || result.role == Role.OUTPUT) {
                                System.out.printf("[%s] %s!%s -> Role: %s, Confidence: %d%%%n",
                                        result.role,
                                        sheetName,
                                        cellAddress,
                                        result.role,
                                        result.confidencePercentage
                                );
                                if (result.evidence != null && !result.evidence.isEmpty()) {
                                    System.out.println("  Evidence: " + String.join("; ", result.evidence));
                                }
                            }
                        }
                    }
                }
            }

            // 5. Output JSON Schema Report
            File reportFile = new File("schema-report.json");
            Gson gson = new GsonBuilder().setPrettyPrinting().create();
            try (FileWriter writer = new FileWriter(reportFile)) {
                gson.toJson(classifiedCells, writer);
            }

            System.out.printf("%n[SUCCESS] Schema report saved to: %s%n", reportFile.getAbsolutePath());
            System.out.printf("Total classified nodes: %d%n", classifiedCells.size());

        } catch (Exception e) {
            System.err.println("Error processing workbook: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }

    /**
     * Determines the dominant style index from frequency counts.
     *
     * @param counts map of style index to occurrence count
     * @return dominant style index, or -1 if empty
     */
    private static short getDominantStyleIndex(Map<Short, Integer> counts) {
        short dominant = -1;
        int max = 0;
        for (Map.Entry<Short, Integer> entry : counts.entrySet()) {
            if (entry.getValue() > max) {
                max = entry.getValue();
                dominant = entry.getKey();
            }
        }
        return dominant;
    }
}
