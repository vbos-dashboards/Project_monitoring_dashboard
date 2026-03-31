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

        List<Map<String, String>> projects = readProjects(inputPath);
        Map<String, Object> payload = buildPayload(projects);

        Files.createDirectories(outputPath.getParent());
        ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
        mapper.writeValue(outputPath.toFile(), payload);

        System.out.println("Exported " + projects.size() + " records to " + outputPath.toAbsolutePath());
    }

    private static List<Map<String, String>> readProjects(Path inputPath) throws IOException {
        List<Map<String, String>> rows = new ArrayList<>();
        DataFormatter formatter = new DataFormatter();

        try (FileInputStream fis = new FileInputStream(inputPath.toFile());
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null) {
                return rows;
            }

            Row headerRow = sheet.getRow(sheet.getFirstRowNum());
            if (headerRow == null) {
                return rows;
            }

            List<String> headers = extractHeaders(headerRow, formatter);

            for (int i = sheet.getFirstRowNum() + 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    continue;
                }

                Map<String, String> entry = new LinkedHashMap<>();
                boolean hasValue = false;

                for (int c = 0; c < headers.size(); c++) {
                    String header = headers.get(c);
                    String value = formatter.formatCellValue(row.getCell(c));
                    if (!value.isBlank()) {
                        hasValue = true;
                    }
                    entry.put(header, value);
                }

                if (hasValue) {
                    rows.add(entry);
                }
            }
        }

        return rows;
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

    private static Map<String, Object> buildPayload(List<Map<String, String>> projects) {
        Map<String, Object> payload = new HashMap<>();
        payload.put("generatedAt", java.time.Instant.now().toString());
        payload.put("totalProjects", projects.size());
        payload.put("projects", projects);
        return payload;
    }
}
