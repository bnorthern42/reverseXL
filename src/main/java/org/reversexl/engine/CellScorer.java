package org.reversexl.engine;

import java.util.ArrayList;
import java.util.List;

/**
 * Heuristic classification engine that evaluates cell context signals,
 * scores candidate roles, and infers functional semantic roles (Input, Output, Intermediate).
 */
public class CellScorer {

    /**
     * Private constructor to prevent instantiation of utility class.
     */
    private CellScorer() {}

    /**
     * Functional roles that a spreadsheet cell can fulfill.
     */
    public enum Role {
        /** User-configurable parameter or model boundary input. */
        USER_INPUT,
        /** Static configuration setting. */
        CONFIGURATION,
        /** Lookup table or reference database entry. */
        REFERENCE_DATA,
        /** Multi-tier intermediate calculation formula in the computational DAG. */
        INTERMEDIATE,
        /** Final business calculation, KPI, or exported summary metric. */
        OUTPUT,
        /** Visual display or presentation element. */
        DISPLAY,
        /** Interactive control element (button, checkbox). */
        CONTROL,
        /** Cell without sufficient discriminating heuristic signals. */
        UNKNOWN
    }

    /**
     * Immutable data holder representing the outcome of a cell heuristic classification.
     */
    public static class ClassificationResult {
        /** The inferred functional role. */
        public Role role;
        /** Confidence score (0 to 100 percentage). */
        public int confidencePercentage;
        /** Human-readable rationale and triggering heuristic evidence strings. */
        public List<String> evidence;

        /**
         * Constructs a new ClassificationResult.
         *
         * @param role the classified role
         * @param confidencePercentage confidence level clamped between 0 and 100
         * @param evidence list of supporting heuristic evidence strings
         */
        public ClassificationResult(Role role, int confidencePercentage, List<String> evidence) {
            this.role = role;
            this.confidencePercentage = Math.min(100, Math.max(0, confidencePercentage));
            this.evidence = evidence;
        }
    }

    /**
     * Evaluates the provided {@link CellContext} against the heuristic scoring matrix
     * and assigns an inferred {@link Role}.
     *
     * @param ctx cell attributes and contextual environment
     * @return classification result containing role, confidence, and evidence
     */
    public static ClassificationResult classify(CellContext ctx) {
        int inputScore = 0;
        int outputScore = 0;
        List<String> evidence = new ArrayList<>();

        // --- INPUT SCORING ---
        if (!ctx.isLocked() && ctx.isSheetProtected()) {
            inputScore += 5;
            evidence.add("Unlocked cell in a protected sheet");
        }
        if (ctx.hasDataValidation()) {
            inputScore += 4;
            evidence.add("Has data validation constraints");
        }
        if (ctx.isClearedByMacro("ClearInputs")) {
            inputScore += 4;
            evidence.add("Cleared by Input-clearing VBA macro");
        }
        if (!ctx.isFormula() && ctx.getDownstreamDependentsCount() > 0) {
            inputScore += 3;
            evidence.add("Constant value referenced by formulas");
        }
        if (ctx.matchesInputStyleCluster()) {
            inputScore += 3;
            evidence.add("Matches dominant input style cluster");
        }
        if (ctx.hasNamedRange()) {
            inputScore += 2;
            outputScore += 2; // Named ranges are used for both
            evidence.add("Has named range: " + ctx.getNamedRange());
        }
        if (ctx.hasAdjacentLabel()) {
            inputScore += 2;
            outputScore += 2;
            evidence.add("Has adjacent label: '" + ctx.getAdjacentLabel() + "'");
        }
        if (ctx.getDownstreamDependentsCount() > 3) {
            inputScore += 1;
        }
        if (ctx.isFormula()) {
            inputScore -= 5;
        }

        // --- OUTPUT SCORING ---
        if (ctx.isReferencedByMacro("ExportOutputs") || ctx.isClearedByMacro("ClearOutputs")) {
            outputScore += 5;
            evidence.add("Referenced by Output-related VBA macro");
        }
        if (ctx.isFormula() && ctx.getDownstreamDependentsCount() == 0) {
            outputScore += 4;
            evidence.add("Formula with zero dependents (Leaf node)");
        }
        if (ctx.matchesOutputStyleCluster()) {
            outputScore += 3;
            evidence.add("Matches dominant output style cluster");
        }
        if (ctx.getSheetName().matches("(?i).*(Result|Summary|Performance).*")) {
            outputScore += 3;
            evidence.add("Located on Result/Summary sheet");
        }
        if (ctx.hasAdjacentUnit()) {
            outputScore += 2;
            evidence.add("Unit detected beside cell");
        }
        if (ctx.isReferencedByChart()) {
            outputScore += 2;
            evidence.add("Referenced by a chart");
        }
        if (ctx.isHidden()) {
            outputScore -= 3;
        }

        // --- CLASSIFICATION LOGIC ---
        if (ctx.isFormula() && ctx.getUpstreamDependenciesCount() > 0 && ctx.getDownstreamDependentsCount() > 0) {
            return new ClassificationResult(Role.INTERMEDIATE, 90, List.of("Formula with both upstream and downstream dependencies"));
        }

        if (inputScore > outputScore && inputScore > 4) {
            int confidence = Math.min(100, (inputScore * 10)); // Rough confidence mapping
            return new ClassificationResult(Role.USER_INPUT, confidence, evidence);
        } else if (outputScore > inputScore && outputScore > 4) {
            int confidence = Math.min(100, (outputScore * 10));
            return new ClassificationResult(Role.OUTPUT, confidence, evidence);
        }

        return new ClassificationResult(Role.UNKNOWN, 0, List.of("Insufficient heuristic signals"));
    }
}
