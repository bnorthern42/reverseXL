package org.reversexl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.reversexl.engine.PoiCellContext;

import java.util.Collections;
import java.util.List;
import java.util.Set;

/**
 * Unit test suite verifying the Apache POI adapter logic in {@link PoiCellContext},
 * including data validation intersections, named range mapping, style clustering, and dependency counts.
 */
public class PoiCellContextTest {

    /**
     * Executes all test cases for the POI context adapter.
     *
     * @param args command-line arguments
     * @throws Exception if POI workbook generation fails
     */
    public static void main(String[] args) throws Exception {
        StringTddRunner.reset();
        System.out.println("Starting PoiCellContext TDD Suite...\n");

        testDataValidationAndNamedRanges();
        testStyleClustering();
        testDependencies();
        testMacroClearing();

        StringTddRunner.summary();
        if (StringTddRunner.getFailedCount() > 0) {
            System.exit(1);
        }
    }

    /**
     * Verifies that cells within Data Validation regions and Named Ranges are correctly recognized.
     *
     * @throws Exception if workbook manipulation fails
     */
    private static void testDataValidationAndNamedRanges() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("TestSheet");
            Row row = sheet.createRow(2);
            Cell b3 = row.createCell(1); // B3 (row 2, col 1)
            Cell c3 = row.createCell(2); // C3 (row 2, col 2)

            // Add DataValidation on B3:B5
            DataValidationHelper dvHelper = sheet.getDataValidationHelper();
            DataValidationConstraint constraint = dvHelper.createNumericConstraint(
                    DataValidationConstraint.ValidationType.INTEGER,
                    DataValidationConstraint.OperatorType.BETWEEN, "1", "100"
            );
            CellRangeAddressList addressList = new CellRangeAddressList(2, 4, 1, 1);
            DataValidation validation = dvHelper.createValidation(constraint, addressList);
            sheet.addValidationData(validation);

            List<? extends DataValidation> validations = sheet.getDataValidations();

            // Test B3 (inside validation region, has named range)
            PoiCellContext ctxB3 = new PoiCellContext(b3, validations, "USER_AGE");
            StringTddRunner.assertEquals("B3 has Data Validation", true, ctxB3.hasDataValidation());
            StringTddRunner.assertEquals("B3 has Named Range", true, ctxB3.hasNamedRange());
            StringTddRunner.assertEquals("B3 named range name matches", "USER_AGE", ctxB3.getNamedRange());

            // Test C3 (outside validation region, no named range)
            PoiCellContext ctxC3 = new PoiCellContext(c3, validations, null);
            StringTddRunner.assertEquals("C3 does not have Data Validation", false, ctxC3.hasDataValidation());
            StringTddRunner.assertEquals("C3 does not have Named Range", false, ctxC3.hasNamedRange());
            StringTddRunner.assertEquals("C3 named range string is empty", "", ctxC3.getNamedRange());
        }
    }

    /**
     * Verifies that cell styles accurately match dominant input/output style cluster indices.
     *
     * @throws Exception if workbook manipulation fails
     */
    private static void testStyleClustering() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("StyleTestSheet");
            Row row = sheet.createRow(0);

            CellStyle inputStyle = wb.createCellStyle();
            CellStyle outputStyle = wb.createCellStyle();

            Cell cellA1 = row.createCell(0);
            cellA1.setCellStyle(inputStyle);

            Cell cellB1 = row.createCell(1);
            cellB1.setCellStyle(outputStyle);

            short inputIdx = inputStyle.getIndex();
            short outputIdx = outputStyle.getIndex();

            PoiCellContext ctxA1 = new PoiCellContext(cellA1, Collections.emptyList(), null, inputIdx, outputIdx);
            StringTddRunner.assertEquals("A1 matches input style cluster", true, ctxA1.matchesInputStyleCluster());
            StringTddRunner.assertEquals("A1 does not match output style cluster", false, ctxA1.matchesOutputStyleCluster());

            PoiCellContext ctxB1 = new PoiCellContext(cellB1, Collections.emptyList(), null, inputIdx, outputIdx);
            StringTddRunner.assertEquals("B1 does not match input style cluster", false, ctxB1.matchesInputStyleCluster());
            StringTddRunner.assertEquals("B1 matches output style cluster", true, ctxB1.matchesOutputStyleCluster());
        }
    }

    /**
     * Verifies that upstream and downstream dependency counts are accurately reflected in the context.
     *
     * @throws Exception if workbook manipulation fails
     */
    private static void testDependencies() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("DepTestSheet");
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);

            PoiCellContext ctx = new PoiCellContext(cell, Collections.emptyList(), null, (short) -1, (short) -1, 5, 3);
            StringTddRunner.assertEquals("Downstream count matches", 5, ctx.getDownstreamDependentsCount());
            StringTddRunner.assertEquals("Upstream count matches", 3, ctx.getUpstreamDependenciesCount());
        }
    }

    /**
     * Verifies that cells targeted by VBA macro clearing routines return true from isClearedByMacro.
     *
     * @throws Exception if workbook manipulation fails
     */
    private static void testMacroClearing() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("MacroSheet");
            Row row = sheet.createRow(9);
            Cell c10 = row.createCell(2); // C10
            Cell d10 = row.createCell(3); // D10

            Set<String> cleared = Set.of("C10");
            PoiCellContext ctxC10 = new PoiCellContext(
                    c10, Collections.emptyList(), null, (short) -1, (short) -1, 0, 0, cleared
            );
            StringTddRunner.assertEquals("C10 is cleared by macro", true, ctxC10.isClearedByMacro("ClearInputs"));

            PoiCellContext ctxD10 = new PoiCellContext(
                    d10, Collections.emptyList(), null, (short) -1, (short) -1, 0, 0, cleared
            );
            StringTddRunner.assertEquals("D10 is not cleared by macro", false, ctxD10.isClearedByMacro("ClearInputs"));
        }
    }
}
