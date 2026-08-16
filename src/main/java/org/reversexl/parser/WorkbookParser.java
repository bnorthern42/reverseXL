package org.reversexl.parser;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.reversexl.engine.CellContext;
import org.reversexl.engine.CellScorer;
import org.reversexl.engine.SimpleCellContext;

import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.EnumMap;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * High-level parser that analyzes an Excel workbook using Apache POI, extracting
 * style clusters, spatial layout signals, dependency graphs, and running the heuristic classification engine.
 */
public class WorkbookParser {

    private static final Pattern UNIT_PATTERN = Pattern.compile(
            "^(?i)(kg|g|mg|lbs?|oz|m|cm|mm|km|ft|in|yd|mi|s|sec|min|hr|hrs|d|days?|usd|eur|gbp|jpy|aud|cad|chf|cny|\\$|€|£|¥|%|pa|kpa|mpa|psi|bar|w|kw|mw|hp|v|mv|kv|a|ma|hz|khz|mhz|ghz|deg|°c|°f|k|rpm|m/s|km/h|mph|l|ml|gal|m3|cm3|sqm|sqft|items?|units?|qty|pcs|users?|months?|years?|q[1-4])$"
    );

    /**
     * Constructs a new WorkbookParser instance.
     */
    public WorkbookParser() {}

    /**
     * Represents a single analyzed cell entry with its classification result.
     *
     * @param sheetName name of the worksheet
     * @param cellRef cell reference address (e.g., A1)
     * @param cellType POI cell type
     * @param rawValue raw formatted value
     * @param classification heuristic classification result
     */
    public record AnalysisEntry(
            String sheetName,
            String cellRef,
            CellType cellType,
            String rawValue,
            CellScorer.ClassificationResult classification
    ) {}

    /**
     * Encapsulates the complete analysis report for a workbook.
     *
     * @param filename workbook file name
     * @param totalSheets total count of sheets analyzed
     * @param totalCells total count of cells scanned
     * @param entries detailed cell analysis entries
     * @param roleCounts summary distribution of classified roles
     */
    public record WorkbookAnalysisReport(
            String filename,
            int totalSheets,
            int totalCells,
            List<AnalysisEntry> entries,
            Map<CellScorer.Role, Long> roleCounts
    ) {}

    /**
     * Parses and analyzes a workbook from a local file.
     *
     * @param file target Excel file
     * @return analysis report
     * @throws Exception if POI parsing fails
     */
    public WorkbookAnalysisReport analyze(File file) throws Exception {
        try (Workbook workbook = WorkbookFactory.create(file)) {
            return analyze(workbook, file.getName());
        }
    }

    /**
     * Parses and analyzes a workbook from an input stream.
     *
     * @param is input stream of workbook bytes
     * @param filename display file name
     * @return analysis report
     * @throws Exception if POI parsing fails
     */
    public WorkbookAnalysisReport analyze(InputStream is, String filename) throws Exception {
        try (Workbook workbook = WorkbookFactory.create(is)) {
            return analyze(workbook, filename);
        }
    }

    /**
     * Performs end-to-end analysis on an open POI {@link Workbook} instance.
     *
     * @param workbook open workbook instance
     * @param filename display file name
     * @return analysis report
     */
    public WorkbookAnalysisReport analyze(Workbook workbook, String filename) {
        List<AnalysisEntry> entries = new ArrayList<>();
        Map<CellScorer.Role, Long> roleCounts = new EnumMap<>(CellScorer.Role.class);
        for (CellScorer.Role r : CellScorer.Role.values()) {
            roleCounts.put(r, 0L);
        }

        // 1. Build Named Ranges map: "SheetName!CellRef" -> Name
        Map<String, String> namedRanges = new HashMap<>();
        for (Name name : workbook.getAllNames()) {
            if (name != null && !name.isDeleted() && name.getRefersToFormula() != null) {
                String formula = name.getRefersToFormula();
                namedRanges.put(formula.replace("$", "").toUpperCase(), name.getNameName());
            }
        }

        // 2. Build dependency graph across all formula cells
        Map<String, Set<String>> downstreamGraph = new HashMap<>(); // Source cell -> cells that depend on it
        Map<String, Set<String>> upstreamGraph = new HashMap<>();   // Formula cell -> source cells it references

        int totalSheets = workbook.getNumberOfSheets();
        int totalCells = 0;

        for (int s = 0; s < totalSheets; s++) {
            Sheet sheet = workbook.getSheetAt(s);
            String sheetName = sheet.getSheetName();

            for (Row row : sheet) {
                for (Cell cell : row) {
                    totalCells++;
                    if (cell.getCellType() == CellType.FORMULA) {
                        String currentKey = makeKey(sheetName, cell.getColumnIndex(), cell.getRowIndex());
                        upstreamGraph.putIfAbsent(currentKey, new HashSet<>());

                        try {
                            String formula = cell.getCellFormula();
                            Set<String> referencedKeys = extractReferencedKeys(sheetName, formula);
                            for (String refKey : referencedKeys) {
                                upstreamGraph.get(currentKey).add(refKey);
                                downstreamGraph.computeIfAbsent(refKey, k -> new HashSet<>()).add(currentKey);
                            }
                        } catch (Exception ignored) {
                            // POI formula parse fallback
                        }
                    }
                }
            }
        }

        // 3. Identify dominant style clusters for input (constants with downstream dependents) and output (formulas with 0 dependents)
        Map<Short, Integer> inputStyleCounts = new HashMap<>();
        Map<Short, Integer> outputStyleCounts = new HashMap<>();

        for (int s = 0; s < totalSheets; s++) {
            Sheet sheet = workbook.getSheetAt(s);
            String sheetName = sheet.getSheetName();

            for (Row row : sheet) {
                for (Cell cell : row) {
                    String key = makeKey(sheetName, cell.getColumnIndex(), cell.getRowIndex());
                    short styleIndex = cell.getCellStyle() != null ? cell.getCellStyle().getIndex() : -1;
                    if (styleIndex >= 0) {
                        int downstream = downstreamGraph.getOrDefault(key, Collections.emptySet()).size();
                        if (cell.getCellType() != CellType.FORMULA && downstream > 0) {
                            inputStyleCounts.merge(styleIndex, 1, Integer::sum);
                        } else if (cell.getCellType() == CellType.FORMULA && downstream == 0) {
                            outputStyleCounts.merge(styleIndex, 1, Integer::sum);
                        }
                    }
                }
            }
        }

        short dominantInputStyle = getDominantStyle(inputStyleCounts);
        short dominantOutputStyle = getDominantStyle(outputStyleCounts);

        // 4. Classify each cell
        for (int s = 0; s < totalSheets; s++) {
            Sheet sheet = workbook.getSheetAt(s);
            String sheetName = sheet.getSheetName();
            boolean isSheetProtected = sheet.getProtect();

            // Data validation ranges in this sheet
            List<CellRangeAddress> validationRanges = new ArrayList<>();
            for (DataValidation dv : sheet.getDataValidations()) {
                if (dv.getRegions() != null) {
                    validationRanges.addAll(Arrays.asList(dv.getRegions().getCellRangeAddresses()));
                }
            }

            for (Row row : sheet) {
                for (Cell cell : row) {
                    int colIdx = cell.getColumnIndex();
                    int rowIdx = cell.getRowIndex();
                    String cellRef = new CellReference(rowIdx, colIdx).formatAsString();
                    String cellKey = makeKey(sheetName, colIdx, rowIdx);

                    CellStyle style = cell.getCellStyle();
                    boolean isLocked = style == null || style.getLocked();
                    boolean isFormula = cell.getCellType() == CellType.FORMULA;
                    short styleIdx = style != null ? style.getIndex() : -1;

                    boolean hasValidation = validationRanges.stream().anyMatch(r -> r.isInRange(rowIdx, colIdx));
                    String namedRange = namedRanges.get(sheetName.toUpperCase() + "!" + cellRef.replace("$", "").toUpperCase());
                    if (namedRange == null) {
                        namedRange = namedRanges.get(cellRef.replace("$", "").toUpperCase());
                    }

                    // Spatial layout: adjacent labels (left or above)
                    String adjacentLabel = findAdjacentLabel(sheet, rowIdx, colIdx);
                    boolean hasAdjacentUnit = checkAdjacentUnit(sheet, rowIdx, colIdx);

                    int downstreamCount = downstreamGraph.getOrDefault(cellKey, Collections.emptySet()).size();
                    int upstreamCount = upstreamGraph.getOrDefault(cellKey, Collections.emptySet()).size();

                    CellContext ctx = SimpleCellContext.builder()
                            .sheetName(sheetName)
                            .locked(isLocked)
                            .sheetProtected(isSheetProtected)
                            .dataValidation(hasValidation)
                            .formula(isFormula)
                            .downstreamDependentsCount(downstreamCount)
                            .upstreamDependenciesCount(upstreamCount)
                            .matchesInputStyleCluster(styleIdx != -1 && styleIdx == dominantInputStyle)
                            .matchesOutputStyleCluster(styleIdx != -1 && styleIdx == dominantOutputStyle)
                            .namedRange(namedRange)
                            .adjacentLabel(adjacentLabel)
                            .adjacentUnit(hasAdjacentUnit)
                            .hidden(sheet.isColumnHidden(colIdx) || row.getZeroHeight())
                            .build();

                    CellScorer.ClassificationResult result = CellScorer.classify(ctx);
                    roleCounts.put(result.role, roleCounts.get(result.role) + 1);

                    String rawVal = formatCellValue(cell);
                    entries.add(new AnalysisEntry(sheetName, cellRef, cell.getCellType(), rawVal, result));
                }
            }
        }

        return new WorkbookAnalysisReport(filename, totalSheets, totalCells, entries, roleCounts);
    }

    /**
     * Generates a normalized lookup key for a cell reference (e.g., SHEET1!A1).
     *
     * @param sheetName sheet name
     * @param col 0-based column index
     * @param row 0-based row index
     * @return normalized uppercase key
     */
    private static String makeKey(String sheetName, int col, int row) {
        return sheetName.toUpperCase() + "!" + new CellReference(row, col).formatAsString();
    }

    /**
     * Determines the dominant style index from frequency counts with a minimum threshold of 2.
     *
     * @param counts map of style indices to occurrence counts
     * @return dominant style index, or -1 if no cluster found
     */
    private static short getDominantStyle(Map<Short, Integer> counts) {
        return counts.entrySet().stream()
                .filter(e -> e.getValue() >= 2)
                .max(Map.Entry.comparingByValue())
                .map(Map.Entry::getKey)
                .orElse((short) -1);
    }

    /**
     * Extracts cell references referenced in a formula string.
     *
     * @param currentSheet default current sheet name
     * @param formula formula text
     * @return set of normalized referenced cell keys
     */
    private static Set<String> extractReferencedKeys(String currentSheet, String formula) {
        Set<String> refs = new HashSet<>();
        Pattern cellPattern = Pattern.compile("(?i)(?:'([^']+)'|([A-Za-z0-9_]+))?!?\\$?([A-Za-z]{1,3})\\$?([0-9]{1,7})");
        Matcher matcher = cellPattern.matcher(formula);
        while (matcher.find()) {
            String sheet = matcher.group(1);
            if (sheet == null) sheet = matcher.group(2);
            if (sheet == null) sheet = currentSheet;

            String colStr = matcher.group(3);
            String rowStr = matcher.group(4);
            try {
                int row = Integer.parseInt(rowStr) - 1;
                int col = CellReference.convertColStringToIndex(colStr);
                refs.add(sheet.toUpperCase() + "!" + new CellReference(row, col).formatAsString());
            } catch (Exception ignored) {}
        }
        return refs;
    }

    /**
     * Finds an adjacent label string 1-2 cells to the left or 1 cell above.
     *
     * @param sheet sheet instance
     * @param rowIdx row index
     * @param colIdx column index
     * @return found label text, or null if none
     */
    private static String findAdjacentLabel(Sheet sheet, int rowIdx, int colIdx) {
        // 1-2 cells to the left
        for (int c = colIdx - 1; c >= Math.max(0, colIdx - 2); c--) {
            Row r = sheet.getRow(rowIdx);
            if (r != null) {
                Cell cell = r.getCell(c);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String str = cell.getStringCellValue().trim();
                    if (!str.isEmpty() && !UNIT_PATTERN.matcher(str).matches()) {
                        return str;
                    }
                }
            }
        }
        // 1 cell above
        if (rowIdx > 0) {
            Row r = sheet.getRow(rowIdx - 1);
            if (r != null) {
                Cell cell = r.getCell(colIdx);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String str = cell.getStringCellValue().trim();
                    if (!str.isEmpty() && !UNIT_PATTERN.matcher(str).matches()) {
                        return str;
                    }
                }
            }
        }
        return null;
    }

    /**
     * Checks if an engineering unit is present immediately to the right.
     *
     * @param sheet sheet instance
     * @param rowIdx row index
     * @param colIdx column index
     * @return true if adjacent unit string is present
     */
    private static boolean checkAdjacentUnit(Sheet sheet, int rowIdx, int colIdx) {
        Row r = sheet.getRow(rowIdx);
        if (r != null) {
            Cell cell = r.getCell(colIdx + 1);
            if (cell != null && cell.getCellType() == CellType.STRING) {
                String str = cell.getStringCellValue().trim();
                return UNIT_PATTERN.matcher(str).matches();
            }
        }
        return false;
    }

    /**
     * Formats a cell's raw value as a string representation.
     *
     * @param cell POI cell
     * @return string representation
     */
    private static String formatCellValue(Cell cell) {
        if (cell == null) return "";
        try {
            switch (cell.getCellType()) {
                case STRING: return cell.getStringCellValue();
                case NUMERIC: return String.valueOf(cell.getNumericCellValue());
                case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
                case FORMULA: return "=" + cell.getCellFormula();
                case BLANK: return "";
                default: return "";
            }
        } catch (Exception e) {
            return "";
        }
    }
}
