package org.reversexl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.reversexl.engine.CellScorer.Role;
import org.reversexl.parser.WorkbookParser;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

/**
 * End-to-end integration test verifying that {@link WorkbookParser} accurately analyzes
 * in-memory Apache POI workbooks and produces correct role classifications.
 */
public class WorkbookParserTest {

    /**
     * Executes the workbook parser integration test suite.
     *
     * @param args command-line arguments
     * @throws Exception if POI workbook creation or serialization fails
     */
    public static void main(String[] args) throws Exception {
        StringTddRunner.reset();
        System.out.println("Starting WorkbookParser Integration TDD Suite...\n");

        testWorkbookParsingAndClassification();

        StringTddRunner.summary();
        if (StringTddRunner.getFailedCount() > 0) {
            System.exit(1);
        }
    }

    /**
     * Creates an in-memory workbook with input constants, intermediate calculation formulas,
     * leaf outputs, and spatial labels/units, verifying end-to-end classification.
     *
     * @throws Exception if workbook manipulation fails
     */
    private static void testWorkbookParsingAndClassification() throws Exception {
        // Build an in-memory workbook:
        // A1: "Price" (Label)
        // B1: 100 (Constant Input referenced by formula)
        // A2: "Quantity" (Label)
        // B2: 5 (Constant Input referenced by formula)
        // A3: "Subtotal" (Label)
        // B3: "=B1*B2" (Intermediate formula)
        // A4: "Tax Rate" (Label)
        // B4: 0.1 (Input)
        // A5: "Total Revenue" (Label)
        // B5: "=B3*(1+B4)" (Output formula - Leaf node)
        // C5: "USD" (Unit)
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Summary_Report");
            sheet.protectSheet("password");

            CellStyle unlockedStyle = wb.createCellStyle();
            unlockedStyle.setLocked(false);

            CellStyle lockedStyle = wb.createCellStyle();
            lockedStyle.setLocked(true);

            // Row 0
            Row r0 = sheet.createRow(0);
            Cell a1 = r0.createCell(0);
            a1.setCellValue("Price");
            a1.setCellStyle(lockedStyle);

            Cell b1 = r0.createCell(1);
            b1.setCellValue(100.0);
            b1.setCellStyle(unlockedStyle); // Unlocked in protected sheet

            // Row 1
            Row r1 = sheet.createRow(1);
            Cell a2 = r1.createCell(0);
            a2.setCellValue("Quantity");
            a2.setCellStyle(lockedStyle);

            Cell b2 = r1.createCell(1);
            b2.setCellValue(5.0);
            b2.setCellStyle(unlockedStyle); // Unlocked in protected sheet

            // Row 2
            Row r2 = sheet.createRow(2);
            Cell a3 = r2.createCell(0);
            a3.setCellValue("Subtotal");
            a3.setCellStyle(lockedStyle);

            Cell b3 = r2.createCell(1);
            b3.setCellFormula("B1*B2");
            b3.setCellStyle(lockedStyle);

            // Row 3
            Row r3 = sheet.createRow(3);
            Cell a4 = r3.createCell(0);
            a4.setCellValue("Tax Rate");
            a4.setCellStyle(lockedStyle);

            Cell b4 = r3.createCell(1);
            b4.setCellValue(0.10);
            b4.setCellStyle(unlockedStyle);

            // Row 4
            Row r4 = sheet.createRow(4);
            Cell a5 = r4.createCell(0);
            a5.setCellValue("Total Revenue");
            a5.setCellStyle(lockedStyle);

            Cell b5 = r4.createCell(1);
            b5.setCellFormula("B3*(1+B4)");
            b5.setCellStyle(lockedStyle);

            Cell c5 = r4.createCell(2);
            c5.setCellValue("USD");
            c5.setCellStyle(lockedStyle);

            // Serialize to bytes and parse
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            wb.write(baos);
            byte[] bytes = baos.toByteArray();

            WorkbookParser parser = new WorkbookParser();
            WorkbookParser.WorkbookAnalysisReport report = parser.analyze(new ByteArrayInputStream(bytes), "test_model.xlsx");

            StringTddRunner.assertEquals("Total sheets is 1", 1, report.totalSheets());

            // Check B1 classified as USER_INPUT
            var b1Entry = report.entries().stream().filter(e -> e.cellRef().equals("B1")).findFirst().orElseThrow();
            StringTddRunner.assertEquals("B1 (unlocked constant in protected sheet) is USER_INPUT",
                    Role.USER_INPUT, b1Entry.classification().role);

            // Check B3 classified as INTERMEDIATE
            var b3Entry = report.entries().stream().filter(e -> e.cellRef().equals("B3")).findFirst().orElseThrow();
            StringTddRunner.assertEquals("B3 (formula with upstream & downstream dependencies) is INTERMEDIATE",
                    Role.INTERMEDIATE, b3Entry.classification().role);

            // Check B5 classified as OUTPUT
            var b5Entry = report.entries().stream().filter(e -> e.cellRef().equals("B5")).findFirst().orElseThrow();
            StringTddRunner.assertEquals("B5 (formula leaf node on Summary sheet with USD unit) is OUTPUT",
                    Role.OUTPUT, b5Entry.classification().role);
        }
    }
}
