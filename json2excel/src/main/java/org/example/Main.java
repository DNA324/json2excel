package org.example;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;

/**
 * @author ${USER}
 * @date ${DATE} ${TIME}
 */
public class Main {
    public static void main(String[] args) throws IOException {
        String area = "eu";
        String json = IOUtils.resourceToString("basic_" + area + ".json", Charset.defaultCharset(), Main.class.getClassLoader());
        ObjectMapper mapper = new ObjectMapper();
        JsonNode root = mapper.readTree(json).get("aaData");

        // create a new workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Packages");

        // create header row
        Row headerRow = sheet.createRow(0);
        Cell headerCellName = headerRow.createCell(0);
        headerCellName.setCellValue("Package Name");
        Cell headerCellId = headerRow.createCell(1);
        headerCellId.setCellValue("Package Id");
        Cell headerCellMode = headerRow.createCell(2);
        headerCellMode.setCellValue("Charge Modes");

        // populate data rows
        int rowNum = 1;
        for (JsonNode node : root) {
            String packageName = node.get("package_name").asText();
            String packageId = node.get("package_id").asText();
            String chargeModes = node.get("charge_modes").toString();
            Row row = sheet.createRow(rowNum++);
            Cell cell1 = row.createCell(0);
            cell1.setCellValue(packageName);
            Cell cell2 = row.createCell(1);
            cell2.setCellValue(packageId);
            Cell cell3 = row.createCell(2);
            cell3.setCellValue(chargeModes);
        }

        // autosize columns
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);

        // write to file
        File outputFile = new File("packages_" + area + ".xlsx");
        FileOutputStream outputStream = new FileOutputStream(outputFile);
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }
}