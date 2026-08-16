package org.reversexl;

import org.reversexl.engine.CellScorer;
import org.reversexl.engine.CellScorer.Role;
import org.reversexl.engine.SimpleCellContext;

/**
 * Hermetic unit test suite verifying the heuristic scoring engine rules in {@link CellScorer}.
 */
public class CellScorerTest {

    /**
     * Executes all test cases for the heuristic engine.
     *
     * @param args command-line arguments
     */
    public static void main(String[] args) {
        StringTddRunner.reset();
        System.out.println("Starting CellScorer Heuristic Engine TDD Suite...\n");

        testUnlockedProtectedInput();
        testDataValidationAndMacroInput();
        testFormulaLeafOutput();
        testResultSheetOutput();
        testMacroReferencedOutput();
        testIntermediateFormula();
        testInsufficientSignalsUnknown();
        testAdjacentLabelsAndUnits();

        StringTddRunner.summary();
        if (StringTddRunner.getFailedCount() > 0) {
            System.exit(1);
        }
    }

    /**
     * Verifies that an unlocked cell in a protected worksheet with data validation classifies as USER_INPUT.
     */
    private static void testUnlockedProtectedInput() {
        var ctx = SimpleCellContext.builder()
                .locked(false)
                .sheetProtected(true)
                .dataValidation(true)
                .build();
        var result = CellScorer.classify(ctx);
        StringTddRunner.assertEquals("Unlocked cell with validation in protected sheet classifies as USER_INPUT",
                Role.USER_INPUT, result.role);
        StringTddRunner.assertEquals("Confidence is clamped/calculated correctly",
                90, result.confidencePercentage);
    }

    /**
     * Verifies that a cell targeted by a ClearInputs macro with downstream dependents classifies as USER_INPUT.
     */
    private static void testDataValidationAndMacroInput() {
        var ctx = SimpleCellContext.builder()
                .locked(true)
                .sheetProtected(false)
                .addClearedByMacro("ClearInputs")
                .downstreamDependentsCount(2)
                .formula(false)
                .build();
        var result = CellScorer.classify(ctx);
        StringTddRunner.assertEquals("Macro ClearInputs + downstream dependents classifies as USER_INPUT",
                Role.USER_INPUT, result.role);
    }

    /**
     * Verifies that a leaf formula matching the output style cluster classifies as OUTPUT.
     */
    private static void testFormulaLeafOutput() {
        var ctx = SimpleCellContext.builder()
                .formula(true)
                .downstreamDependentsCount(0)
                .matchesOutputStyleCluster(true)
                .build();
        var result = CellScorer.classify(ctx);
        StringTddRunner.assertEquals("Formula leaf node with output style cluster classifies as OUTPUT",
                Role.OUTPUT, result.role);
        StringTddRunner.assertEquals("Confidence is 70% for leaf + style cluster",
                70, result.confidencePercentage);
    }

    /**
     * Verifies that a leaf formula located on a summary sheet with adjacent units classifies as OUTPUT.
     */
    private static void testResultSheetOutput() {
        var ctx = SimpleCellContext.builder()
                .formula(true)
                .downstreamDependentsCount(0)
                .sheetName("Financial_Summary")
                .adjacentUnit(true)
                .build();
        var result = CellScorer.classify(ctx);
        StringTddRunner.assertEquals("Leaf formula on Summary sheet with adjacent unit classifies as OUTPUT",
                Role.OUTPUT, result.role);
    }

    /**
     * Verifies that a cell referenced by an ExportOutputs macro classifies as OUTPUT.
     */
    private static void testMacroReferencedOutput() {
        var ctx = SimpleCellContext.builder()
                .addReferencedByMacro("ExportOutputs")
                .build();
        var result = CellScorer.classify(ctx);
        StringTddRunner.assertEquals("ExportOutputs macro reference classifies as OUTPUT",
                Role.OUTPUT, result.role);
    }

    /**
     * Verifies that a formula cell with both upstream dependencies and downstream dependents classifies as INTERMEDIATE.
     */
    private static void testIntermediateFormula() {
        var ctx = SimpleCellContext.builder()
                .formula(true)
                .upstreamDependenciesCount(2)
                .downstreamDependentsCount(3)
                .build();
        var result = CellScorer.classify(ctx);
        StringTddRunner.assertEquals("Formula with upstream and downstream dependencies classifies as INTERMEDIATE",
                Role.INTERMEDIATE, result.role);
        StringTddRunner.assertEquals("Intermediate formula confidence is 90%",
                90, result.confidencePercentage);
    }

    /**
     * Verifies that a default plain cell without heuristic signals classifies as UNKNOWN.
     */
    private static void testInsufficientSignalsUnknown() {
        var ctx = SimpleCellContext.builder()
                .locked(true)
                .sheetProtected(false)
                .build();
        var result = CellScorer.classify(ctx);
        StringTddRunner.assertEquals("Plain cell without signals classifies as UNKNOWN",
                Role.UNKNOWN, result.role);
        StringTddRunner.assertEquals("Unknown confidence is 0%",
                0, result.confidencePercentage);
    }

    /**
     * Verifies that a leaf formula with an adjacent label and named range classifies as OUTPUT.
     */
    private static void testAdjacentLabelsAndUnits() {
        var ctx = SimpleCellContext.builder()
                .adjacentLabel("Gross Margin")
                .namedRange("TOTAL_MARGIN")
                .formula(true)
                .downstreamDependentsCount(0)
                .build();
        var result = CellScorer.classify(ctx);
        StringTddRunner.assertEquals("Leaf formula with label and named range classifies as OUTPUT",
                Role.OUTPUT, result.role);
    }
}
