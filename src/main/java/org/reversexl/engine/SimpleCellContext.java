package org.reversexl.engine;

import java.util.HashSet;
import java.util.Set;

/**
 * In-memory POJO implementation of {@link CellContext} with a fluent builder,
 * primarily used for hermetic unit testing and mocking of heuristic scenarios.
 */
public class SimpleCellContext implements CellContext {

    private boolean locked = true;
    private boolean sheetProtected = false;
    private boolean dataValidation = false;
    private Set<String> clearedByMacros = new HashSet<>();
    private Set<String> referencedByMacros = new HashSet<>();
    private boolean formula = false;
    private int downstreamDependentsCount = 0;
    private int upstreamDependenciesCount = 0;
    private boolean matchesInputStyleCluster = false;
    private boolean matchesOutputStyleCluster = false;
    private String namedRange = null;
    private String adjacentLabel = null;
    private String sheetName = "Sheet1";
    private boolean adjacentUnit = false;
    private boolean referencedByChart = false;
    private boolean hidden = false;

    /**
     * Constructs a default SimpleCellContext instance.
     */
    public SimpleCellContext() {}

    /**
     * Creates a new fluent {@link Builder} instance.
     *
     * @return a new builder
     */
    public static Builder builder() {
        return new Builder();
    }

    @Override
    public boolean isLocked() {
        return locked;
    }

    @Override
    public boolean isSheetProtected() {
        return sheetProtected;
    }

    @Override
    public boolean hasDataValidation() {
        return dataValidation;
    }

    @Override
    public boolean isClearedByMacro(String macroName) {
        return clearedByMacros.contains(macroName);
    }

    @Override
    public boolean isFormula() {
        return formula;
    }

    @Override
    public int getDownstreamDependentsCount() {
        return downstreamDependentsCount;
    }

    @Override
    public boolean matchesInputStyleCluster() {
        return matchesInputStyleCluster;
    }

    @Override
    public boolean hasNamedRange() {
        return namedRange != null && !namedRange.isEmpty();
    }

    @Override
    public String getNamedRange() {
        return namedRange;
    }

    @Override
    public boolean hasAdjacentLabel() {
        return adjacentLabel != null && !adjacentLabel.isEmpty();
    }

    @Override
    public String getAdjacentLabel() {
        return adjacentLabel;
    }

    @Override
    public boolean isReferencedByMacro(String macroName) {
        return referencedByMacros.contains(macroName);
    }

    @Override
    public boolean matchesOutputStyleCluster() {
        return matchesOutputStyleCluster;
    }

    @Override
    public String getSheetName() {
        return sheetName;
    }

    @Override
    public boolean hasAdjacentUnit() {
        return adjacentUnit;
    }

    @Override
    public boolean isReferencedByChart() {
        return referencedByChart;
    }

    @Override
    public boolean isHidden() {
        return hidden;
    }

    @Override
    public int getUpstreamDependenciesCount() {
        return upstreamDependenciesCount;
    }

    /**
     * Fluent builder for constructing {@link SimpleCellContext} instances.
     */
    public static class Builder {
        private final SimpleCellContext ctx = new SimpleCellContext();

        /**
         * Constructs a new Builder instance.
         */
        public Builder() {}

        /**
         * Sets the locked cell style flag.
         *
         * @param locked true if locked
         * @return this builder
         */
        public Builder locked(boolean locked) { ctx.locked = locked; return this; }

        /**
         * Sets the sheet protection flag.
         *
         * @param sheetProtected true if protected
         * @return this builder
         */
        public Builder sheetProtected(boolean sheetProtected) { ctx.sheetProtected = sheetProtected; return this; }

        /**
         * Sets the data validation flag.
         *
         * @param dataValidation true if data validation is present
         * @return this builder
         */
        public Builder dataValidation(boolean dataValidation) { ctx.dataValidation = dataValidation; return this; }

        /**
         * Adds a macro name targeting this cell for clearing.
         *
         * @param macroName macro routine name
         * @return this builder
         */
        public Builder addClearedByMacro(String macroName) { ctx.clearedByMacros.add(macroName); return this; }

        /**
         * Adds a macro name referencing this cell for export.
         *
         * @param macroName macro routine name
         * @return this builder
         */
        public Builder addReferencedByMacro(String macroName) { ctx.referencedByMacros.add(macroName); return this; }

        /**
         * Sets whether the cell contains a formula.
         *
         * @param formula true if formula
         * @return this builder
         */
        public Builder formula(boolean formula) { ctx.formula = formula; return this; }

        /**
         * Sets the downstream dependents count.
         *
         * @param count count of downstream formulas
         * @return this builder
         */
        public Builder downstreamDependentsCount(int count) { ctx.downstreamDependentsCount = count; return this; }

        /**
         * Sets the upstream dependencies count.
         *
         * @param count count of upstream cell references
         * @return this builder
         */
        public Builder upstreamDependenciesCount(int count) { ctx.upstreamDependenciesCount = count; return this; }

        /**
         * Sets whether the cell matches the dominant input style cluster.
         *
         * @param matches true if matches
         * @return this builder
         */
        public Builder matchesInputStyleCluster(boolean matches) { ctx.matchesInputStyleCluster = matches; return this; }

        /**
         * Sets whether the cell matches the dominant output style cluster.
         *
         * @param matches true if matches
         * @return this builder
         */
        public Builder matchesOutputStyleCluster(boolean matches) { ctx.matchesOutputStyleCluster = matches; return this; }

        /**
         * Sets the Named Range bound to the cell.
         *
         * @param namedRange named range string
         * @return this builder
         */
        public Builder namedRange(String namedRange) { ctx.namedRange = namedRange; return this; }

        /**
         * Sets the adjacent label text.
         *
         * @param adjacentLabel label text
         * @return this builder
         */
        public Builder adjacentLabel(String adjacentLabel) { ctx.adjacentLabel = adjacentLabel; return this; }

        /**
         * Sets the sheet name.
         *
         * @param sheetName worksheet name
         * @return this builder
         */
        public Builder sheetName(String sheetName) { ctx.sheetName = sheetName; return this; }

        /**
         * Sets whether the cell has an adjacent unit.
         *
         * @param adjacentUnit true if adjacent unit exists
         * @return this builder
         */
        public Builder adjacentUnit(boolean adjacentUnit) { ctx.adjacentUnit = adjacentUnit; return this; }

        /**
         * Sets whether the cell is referenced by a chart.
         *
         * @param referencedByChart true if chart reference exists
         * @return this builder
         */
        public Builder referencedByChart(boolean referencedByChart) { ctx.referencedByChart = referencedByChart; return this; }

        /**
         * Sets whether the cell is hidden.
         *
         * @param hidden true if cell is hidden
         * @return this builder
         */
        public Builder hidden(boolean hidden) { ctx.hidden = hidden; return this; }

        /**
         * Builds the configured {@link SimpleCellContext} instance.
         *
         * @return built context
         */
        public SimpleCellContext build() {
            return ctx;
        }
    }
}
