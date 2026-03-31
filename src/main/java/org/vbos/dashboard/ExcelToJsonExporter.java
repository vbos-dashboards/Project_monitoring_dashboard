package org.vbos.dashboard;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class ExcelToJsonExporter {
    private static final String INPUT_FILE = "Projects.xlsx";
    private static final String OUTPUT_FILE = "docs/data/projects.json";

    public static void main(String[] args) throws IOException {
        Path inputPath = args.length > 0 ? Paths.get(args[0]) : Paths.get(INPUT_FILE);
        Path outputPath = args.length > 1 ? Paths.get(args[1]) : Paths.get(OUTPUT_FILE);

        if (!Files.exists(inputPath)) {
            throw new IllegalArgumentException("Excel file not found: " + inputPath.toAbsolutePath());
        }

        Map<String, Object> payload = readWorkbook(inputPath);

        Files.createDirectories(outputPath.getParent());
        ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
        mapper.writeValue(outputPath.toFile(), payload);

        System.out.println("Exported workbook data to " + outputPath.toAbsolutePath());
    }

    private static Map<String, Object> readWorkbook(Path inputPath) throws IOException {
        DataFormatter formatter = new DataFormatter();
        List<Map<String, Object>> sheetsPayload = new ArrayList<>();
        List<Map<String, String>> firstSheetRows = new ArrayList<>();
        int totalRecords = 0;

        try (FileInputStream fis = new FileInputStream(inputPath.toFile());
             Workbook workbook = new XSSFWorkbook(fis)) {
            for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
                Sheet sheet = workbook.getSheetAt(s);
                if (sheet == null) {
                    continue;
                }

                int headerRowIndex = detectHeaderRowIndex(sheet, formatter);
                if (headerRowIndex < 0) {
                    continue;
                }

                Row headerRow = sheet.getRow(headerRowIndex);
                List<String> headers = extractHeaders(headerRow, formatter);
                List<Map<String, String>> rows = readRows(sheet, headers, headerRowIndex + 1, formatter);
                totalRecords += rows.size();

                Map<String, Object> sheetEntry = new LinkedHashMap<>();
                sheetEntry.put("name", sheet.getSheetName());
                sheetEntry.put("headerRow", headerRowIndex + 1);
                sheetEntry.put("headers", headers);
                sheetEntry.put("totalRows", rows.size());
                sheetEntry.put("rows", rows);
                sheetsPayload.add(sheetEntry);

                if (s == 0) {
                    firstSheetRows = rows;
                }
            }
        }

        Map<String, Object> payload = new HashMap<>();
        payload.put("generatedAt", java.time.Instant.now().atOffset(ZoneOffset.UTC).format(DateTimeFormatter.ISO_OFFSET_DATE_TIME));
        payload.put("totalSheets", sheetsPayload.size());
        payload.put("totalProjects", firstSheetRows.size());
        payload.put("totalRecords", totalRecords);
        payload.put("projects", firstSheetRows);
        payload.put("sheets", sheetsPayload);
        return payload;
    }

    private static List<String> extractHeaders(Row headerRow, DataFormatter formatter) {
        List<String> headers = new ArrayList<>();
        short minCell = headerRow.getFirstCellNum();
        short maxCell = headerRow.getLastCellNum();

        for (int c = minCell; c < maxCell; c++) {
            Cell cell = headerRow.getCell(c);
            if (cell == null || cell.getCellType() == CellType.BLANK) {
                headers.add("column_" + c);
                continue;
            }
            String value = formatter.formatCellValue(cell).trim();
            headers.add(value.isEmpty() ? "column_" + c : value);
        }
        return headers;
    }

    private static int detectHeaderRowIndex(Sheet sheet, DataFormatter formatter) {
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            int nonEmpty = 0;
            int stringLike = 0;
            boolean hasProjectKeyword = false;

            for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
                if (c < 0) {
                    continue;
                }
                Cell cell = row.getCell(c);
                String value = formatter.formatCellValue(cell).trim();
                if (value.isEmpty()) {
                    continue;
                }
                nonEmpty++;
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    stringLike++;
                }
                if (value.toLowerCase().contains("project") || value.toLowerCase().contains("status")) {
                    hasProjectKeyword = true;
                }
            }

            if (nonEmpty >= 6 && stringLike >= 4 && hasProjectKeyword) {
                return i;
            }
        }
        return -1;
    }

    private static List<Map<String, String>> readRows(
        Sheet sheet,
        List<String> headers,
        int startRowIndex,
        DataFormatter formatter
    ) {
        List<Map<String, String>> rows = new ArrayList<>();
        for (int i = startRowIndex; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }

            Map<String, String> entry = new LinkedHashMap<>();
            boolean hasValue = false;
            for (int c = 0; c < headers.size(); c++) {
                String value = formatter.formatCellValue(row.getCell(c));
                if (!value.isBlank()) {
                    hasValue = true;
                }
                entry.put(headers.get(c), value);
            }

            if (hasValue) {
                rows.add(entry);
            }
        }
        return rows;
    }
}
