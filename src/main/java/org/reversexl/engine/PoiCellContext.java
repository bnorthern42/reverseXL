package org.reversexl.engine;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.Collections;
import java.util.List;
import java.util.Set;

/**
 * Adapter that maps a real Apache POI {@link Cell} object to the {@link CellContext} interface,
 * evaluating spatial layouts, style clusters, validation regions, and macro signals.
 */
public class PoiCellContext implements CellContext {

    private final Cell cell;
    private final String adjacentLabel;
    private final boolean adjacentUnit;
    private final String namedRange;
    private final boolean dataValidation;
    private final short dominantInputStyle;
    private final short dominantOutputStyle;
    private final int downstreamDependentsCount;
    private final int upstreamDependenciesCount;
    private final Set<String> clearedCells;

    /**
     * Constructs a basic PoiCellContext for a given cell with default fallback values.
     *
     * @param cell Apache POI cell instance
     */
    public PoiCellContext(Cell cell) {
        this(cell, Collections.emptyList(), null, (short) -1, (short) -1, 0, 0, Collections.emptySet());
    }

    /**
     * Constructs a PoiCellContext with data validation and named range context.
     *
     * @param cell Apache POI cell instance
     * @param sheetValidations active data validation regions on the sheet
     * @param namedRange named range bound to this cell, or null
     */
    public PoiCellContext(Cell cell, List<? extends DataValidation> sheetValidations, String namedRange) {
        this(cell, sheetValidations, namedRange, (short) -1, (short) -1, 0, 0, Collections.emptySet());
    }

    /**
     * Constructs a PoiCellContext with validation, named range, and style cluster indices.
     *
     * @param cell Apache POI cell instance
     * @param sheetValidations active data validation regions on the sheet
     * @param namedRange named range bound to this cell, or null
     * @param dominantInputStyle workbook dominant input style index
     * @param dominantOutputStyle workbook dominant output style index
     */
    public PoiCellContext(Cell cell, List<? extends DataValidation> sheetValidations, String namedRange,
                          short dominantInputStyle, short dominantOutputStyle) {
        this(cell, sheetValidations, namedRange, dominantInputStyle, dominantOutputStyle, 0, 0, Collections.emptySet());
    }

    /**
     * Constructs a PoiCellContext with style clusters and downstream dependent counts.
     *
     * @param cell Apache POI cell instance
     * @param sheetValidations active data validation regions on the sheet
     * @param namedRange named range bound to this cell, or null
     * @param dominantInputStyle workbook dominant input style index
     * @param dominantOutputStyle workbook dominant output style index
     * @param downstreamDependentsCount number of downstream formula cells referencing this cell
     */
    public PoiCellContext(Cell cell, List<? extends DataValidation> sheetValidations, String namedRange,
                          short dominantInputStyle, short dominantOutputStyle, int downstreamDependentsCount) {
        this(cell, sheetValidations, namedRange, dominantInputStyle, dominantOutputStyle, downstreamDependentsCount, 0, Collections.emptySet());
    }

    /**
     * Constructs a PoiCellContext with upstream and downstream dependency counts.
     *
     * @param cell Apache POI cell instance
     * @param sheetValidations active data validation regions on the sheet
     * @param namedRange named range bound to this cell, or null
     * @param dominantInputStyle workbook dominant input style index
     * @param dominantOutputStyle workbook dominant output style index
     * @param downstreamDependentsCount number of downstream formula cells referencing this cell
     * @param upstreamDependenciesCount number of upstream cell references in this cell's formula
     */
    public PoiCellContext(Cell cell, List<? extends DataValidation> sheetValidations, String namedRange,
                          short dominantInputStyle, short dominantOutputStyle,
                          int downstreamDependentsCount, int upstreamDependenciesCount) {
        this(cell, sheetValidations, namedRange, dominantInputStyle, dominantOutputStyle, downstreamDependentsCount, upstreamDependenciesCount, Collections.emptySet());
    }

    /**
     * Full constructor for PoiCellContext.
     *
     * @param cell Apache POI cell instance
     * @param sheetValidations active data validation regions on the sheet
     * @param namedRange named range bound to this cell, or null
     * @param dominantInputStyle workbook dominant input style index
     * @param dominantOutputStyle workbook dominant output style index
     * @param downstreamDependentsCount number of downstream formula cells referencing this cell
     * @param upstreamDependenciesCount number of upstream cell references in this cell's formula
     * @param clearedCells set of uppercase cell references targeted by VBA clear routines
     */
    public PoiCellContext(Cell cell, List<? extends DataValidation> sheetValidations, String namedRange,
                          short dominantInputStyle, short dominantOutputStyle,
                          int downstreamDependentsCount, int upstreamDependenciesCount,
                          Set<String> clearedCells) {
        this.cell = cell;
        this.adjacentLabel = resolveAdjacentLabel(cell);
        this.adjacentUnit = resolveAdjacentUnit(cell);
        this.namedRange = (namedRange != null && !namedRange.trim().isEmpty()) ? namedRange.trim() : null;
        this.dataValidation = resolveDataValidation(cell, sheetValidations);
        this.dominantInputStyle = dominantInputStyle;
        this.dominantOutputStyle = dominantOutputStyle;
        this.downstreamDependentsCount = downstreamDependentsCount;
        this.upstreamDependenciesCount = upstreamDependenciesCount;
        this.clearedCells = clearedCells != null ? clearedCells : Collections.emptySet();
    }

    /**
     * Evaluates whether a cell falls inside any active Data Validation regions on the sheet.
     *
     * @param cell target cell
     * @param sheetValidations list of sheet validations
     * @return true if cell intersects a data validation bounding box
     */
    private static boolean resolveDataValidation(Cell cell, List<? extends DataValidation> sheetValidations) {
        if (cell == null || sheetValidations == null || sheetValidations.isEmpty()) {
            return false;
        }

        int rowIdx = cell.getRowIndex();
        int colIdx = cell.getColumnIndex();

        for (DataValidation dv : sheetValidations) {
            if (dv == null) {
                continue;
            }
            CellRangeAddressList regions = dv.getRegions();
            if (regions == null) {
                continue;
            }
            for (CellRangeAddress range : regions.getCellRangeAddresses()) {
                if (range != null && range.isInRange(rowIdx, colIdx)) {
                    return true;
                }
            }
        }

        return false;
    }

    /**
     * Resolves adjacent label text immediately to the left or above the target cell.
     *
     * @param cell target cell
     * @return adjacent string label, or null if absent
     */
    private static String resolveAdjacentLabel(Cell cell) {
        if (cell == null || cell.getSheet() == null) {
            return null;
        }

        Sheet sheet = cell.getSheet();
        int rowIndex = cell.getRowIndex();
        int colIndex = cell.getColumnIndex();

        // 1. Safely inspect cell immediately to the left (colIndex - 1)
        if (colIndex > 0) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell leftCell = row.getCell(colIndex - 1);
                if (leftCell != null && leftCell.getCellType() == CellType.STRING) {
                    String val = leftCell.getStringCellValue();
                    if (val != null && !val.trim().isEmpty()) {
                        return val.trim();
                    }
                }
            }
        }

        // 2. If null or blank, safely inspect cell immediately above (rowIndex - 1)
        if (rowIndex > 0) {
            Row aboveRow = sheet.getRow(rowIndex - 1);
            if (aboveRow != null) {
                Cell aboveCell = aboveRow.getCell(colIndex);
                if (aboveCell != null && aboveCell.getCellType() == CellType.STRING) {
                    String val = aboveCell.getStringCellValue();
                    if (val != null && !val.trim().isEmpty()) {
                        return val.trim();
                    }
                }
            }
        }

        return null;
    }

    /**
     * Resolves whether an engineering unit string exists immediately to the right of the cell.
     *
     * @param cell target cell
     * @return true if adjacent unit string is present
     */
    private static boolean resolveAdjacentUnit(Cell cell) {
        if (cell == null || cell.getSheet() == null) {
            return false;
        }

        Sheet sheet = cell.getSheet();
        int rowIndex = cell.getRowIndex();
        int colIndex = cell.getColumnIndex();

        // Safely inspect cell immediately to the right (colIndex + 1)
        Row row = sheet.getRow(rowIndex);
        if (row != null) {
            Cell rightCell = row.getCell(colIndex + 1);
            if (rightCell != null && rightCell.getCellType() == CellType.STRING) {
                String val = rightCell.getStringCellValue();
                if (val != null) {
                    String trimmed = val.trim();
                    return trimmed.length() >= 1 && trimmed.length() <= 6;
                }
            }
        }

        return false;
    }

    @Override
    public boolean isLocked() {
        if (cell == null) {
            return false;
        }
        CellStyle style = cell.getCellStyle();
        return style == null || style.getLocked();
    }

    @Override
    public boolean isSheetProtected() {
        if (cell == null || cell.getSheet() == null) {
            return false;
        }
        return cell.getSheet().getProtect();
    }

    @Override
    public boolean hasDataValidation() {
        return dataValidation;
    }

    @Override
    public boolean isClearedByMacro(String macroName) {
        if (cell == null || clearedCells == null || clearedCells.isEmpty()) {
            return false;
        }
        String cellRef = cell.getAddress().formatAsString().toUpperCase();
        return clearedCells.contains(cellRef);
    }

    @Override
    public boolean isFormula() {
        if (cell == null) {
            return false;
        }
        return cell.getCellType() == CellType.FORMULA;
    }

    @Override
    public int getDownstreamDependentsCount() {
        return downstreamDependentsCount;
    }

    @Override
    public boolean matchesInputStyleCluster() {
        if (dominantInputStyle == -1 || cell == null || cell.getCellStyle() == null) {
            return false;
        }
        return cell.getCellStyle().getIndex() == dominantInputStyle;
    }

    @Override
    public boolean hasNamedRange() {
        return namedRange != null && !namedRange.isEmpty();
    }

    @Override
    public String getNamedRange() {
        return namedRange != null ? namedRange : "";
    }

    @Override
    public boolean hasAdjacentLabel() {
        return adjacentLabel != null && !adjacentLabel.isEmpty();
    }

    @Override
    public String getAdjacentLabel() {
        return adjacentLabel != null ? adjacentLabel : "";
    }

    @Override
    public boolean isReferencedByMacro(String macroName) {
        return false;
    }

    @Override
    public boolean matchesOutputStyleCluster() {
        if (dominantOutputStyle == -1 || cell == null || cell.getCellStyle() == null) {
            return false;
        }
        return cell.getCellStyle().getIndex() == dominantOutputStyle;
    }

    @Override
    public String getSheetName() {
        if (cell == null || cell.getSheet() == null) {
            return "";
        }
        return cell.getSheet().getSheetName();
    }

    @Override
    public boolean hasAdjacentUnit() {
        return adjacentUnit;
    }

    @Override
    public boolean isReferencedByChart() {
        return false;
    }

    @Override
    public boolean isHidden() {
        return false;
    }

    @Override
    public int getUpstreamDependenciesCount() {
        return upstreamDependenciesCount;
    }
}
