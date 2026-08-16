package org.reversexl.engine;

/**
 * Abstraction representing a cell's attributes, layout geometry, protection state,
 * style clusters, formula dependency counts, and macro signals used by {@link CellScorer}.
 */
public interface CellContext {

    /**
     * Checks if the cell is locked by its cell style.
     *
     * @return true if the cell is locked, false otherwise
     */
    boolean isLocked();

    /**
     * Checks if the sheet containing this cell is protected.
     *
     * @return true if sheet protection is enabled, false otherwise
     */
    boolean isSheetProtected();

    /**
     * Checks if this cell falls within a Data Validation constraint region.
     *
     * @return true if the cell has data validation applied, false otherwise
     */
    boolean hasDataValidation();

    /**
     * Checks if this cell is targeted for clearing by a VBA macro routine.
     *
     * @param macroName name of the macro routine to check against
     * @return true if the cell is cleared by the macro, false otherwise
     */
    boolean isClearedByMacro(String macroName);

    /**
     * Checks if this cell contains a formula.
     *
     * @return true if the cell type is a formula, false otherwise
     */
    boolean isFormula();

    /**
     * Returns the count of downstream formula cells that directly reference this cell.
     *
     * @return number of downstream formula dependents
     */
    int getDownstreamDependentsCount();

    /**
     * Checks if the cell's style matches the dominant input style cluster of the workbook.
     *
     * @return true if the cell matches the dominant input style, false otherwise
     */
    boolean matchesInputStyleCluster();

    /**
     * Checks if this cell is associated with an explicit Named Range.
     *
     * @return true if a named range is bound to this cell, false otherwise
     */
    boolean hasNamedRange();

    /**
     * Returns the name of the Named Range bound to this cell.
     *
     * @return named range name, or an empty string if none
     */
    String getNamedRange();

    /**
     * Checks if an adjacent descriptive label exists immediately to the left or above.
     *
     * @return true if an adjacent text label is present, false otherwise
     */
    boolean hasAdjacentLabel();

    /**
     * Returns the text value of the adjacent label cell.
     *
     * @return string value of the adjacent label, or an empty string if none
     */
    String getAdjacentLabel();

    /**
     * Checks if this cell is referenced by an export/reporting VBA macro.
     *
     * @param macroName name of the macro routine to check against
     * @return true if referenced by the macro, false otherwise
     */
    boolean isReferencedByMacro(String macroName);

    /**
     * Checks if the cell's style matches the dominant output style cluster of the workbook.
     *
     * @return true if the cell matches the dominant output style, false otherwise
     */
    boolean matchesOutputStyleCluster();

    /**
     * Returns the name of the worksheet containing this cell.
     *
     * @return sheet name
     */
    String getSheetName();

    /**
     * Checks if an adjacent engineering unit string (1–6 characters) exists immediately to the right.
     *
     * @return true if an adjacent unit is present, false otherwise
     */
    boolean hasAdjacentUnit();

    /**
     * Checks if this cell is referenced by any chart series in the workbook.
     *
     * @return true if referenced by a chart, false otherwise
     */
    boolean isReferencedByChart();

    /**
     * Checks if this cell is hidden due to hidden row or column visibility.
     *
     * @return true if the cell is hidden, false otherwise
     */
    boolean isHidden();

    /**
     * Returns the count of upstream cell references contained within this cell's formula.
     *
     * @return number of upstream formula dependencies
     */
    int getUpstreamDependenciesCount();
}
