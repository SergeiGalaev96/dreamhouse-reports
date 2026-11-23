package com.jasper.jasper_service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import java.io.ByteArrayInputStream;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import static org.junit.jupiter.api.Assertions.*;

class EstimateNumberFormatTests {
    @Test
    void blanksZerosAndFormatsWholeAndFractionalNumbers() throws Exception {
        Map<String, Object> item = Map.of("item_name", "Материал", "total_quantity", 85,
                "unit_price", 12.5, "total_amount", 127500,
                "subsection_quantities", Map.of("1", 0));
        Map<String, Object> stage = Map.of("stage_name", "Этап",
                "subsections", List.of(Map.of("subsection_id", 1, "subsection_name", "Колонны")),
                "material_rows", List.of(item), "material_total_amount", 127500, "total_amount", 127500);
        byte[] bytes = new ReportService().generateEstimateStage(new ReportService.ReportData(
                Map.of("stages", List.of(stage)), Map.of()), "xlsx");
        try (XSSFWorkbook book = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            var row = book.getSheetAt(0).getRow(7);
            assertEquals("Материал", row.getCell(1).getStringCellValue());
            assertEquals(CellType.BLANK, row.getCell(3).getCellType());
            DataFormatter formatter = new DataFormatter(Locale.forLanguageTag("ru-RU"));
            assertEquals("85", formatter.formatCellValue(row.getCell(4)));
            assertEquals("12,5", formatter.formatCellValue(row.getCell(5)));
            assertEquals(CellType.NUMERIC, row.getCell(6).getCellType());
            for (var sheet : book) for (var r : sheet) for (Cell cell : r) {
                if (cell.getCellType() == CellType.NUMERIC) {
                    assertNotEquals(0, cell.getNumericCellValue());
                    assertFalse(formatter.formatCellValue(cell).endsWith(","));
                }
            }
        }
    }
}
