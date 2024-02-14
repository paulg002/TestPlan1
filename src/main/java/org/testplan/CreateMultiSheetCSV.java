
package org.testplan;



import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;

public class CreateMultiSheetCSV {

    public static void main(String[] args) {

        ClassLoader classloader =
                org.apache.poi.poifs.filesystem.POIFSFileSystem.class.getClassLoader();
        URL resPath = classloader.getResource(
                "org/apache/poi/poifs/filesystem/POIFSFileSystem.class");
        String path = resPath.getPath();
        System.out.println("The actual POI Path is " + path);

        String[] defaultColumns = {"Section", "Title", "Precondition", "Steps", "Expected", "Actual", "Status", "Automation", "Assignee", "Estimate", "Type", "Notes", "4"};

        String[] columns;
        if (args.length == 0) {
            columns = defaultColumns;
        } else {
            columns = args;
        }

        String filename = "TestPlan.xlsx";
        String url = "https://github.com/paulg002"; // GitHub link
        try (Workbook workbook = new XSSFWorkbook()) {

            // Creating Sheet 1 (Test)
            Sheet sheet1 = workbook.createSheet("Test");
            Row row = sheet1.createRow(0);
            row.createCell(0).setCellValue("Test");
            // Writing headers and data
            for (int i = 1; i < columns.length-1; i++) {
                row.createCell(i).setCellValue(columns[i]);
            }
            String[] heuristics = {
                    "API Testing",
                    "Compatibility Testing",
                    "Data Testing",
                    "Domain Testing",
                    "Flow Testing",
                    "Installability Testing",
                    "Info / Data Testing",
                    "Localizability Testing",
                    "Maintainability Testing",
                    "Operations Testing",
                    "Platforms Testing",
                    "Portability Testing",
                    "Risk Testing",
                    "Scenario Testing",
                    "Supportability Testing",
                    "User Testing",
                    "Usability Testing"
            };

            int rowindex =1;
            for (int i = 0; i < heuristics.length;  i++) {
                sheet1.createRow(rowindex);
                row = sheet1.createRow(++rowindex);
                Cell cell = row.createCell(0); // Create a cell in the first column
                cell.setCellValue(heuristics[i]); // Set the cell value
                Font font = cell.getSheet().getWorkbook().createFont();
                font.setItalic(true); // Set italic to true
                CellStyle style = cell.getCellStyle();
                style.setFont(font);
                cell.setCellStyle(style);
                rowindex++;
            }

            // Creating Sheet 2 (Schedule)
            Sheet sheet2 = workbook.createSheet("Schedule");
            row = sheet2.createRow(0);
            String[] headers = {
                    "Weeks",
                    "Estimate",
                    "Functional",
                    "Performance",
                    "Security",
                    "Notes",
                    "Vacation"
            };
            for (int i = 0; i < headers.length; i++) {
                row.createCell(i).setCellValue(headers[i]);
            }

            row.createCell(0).setCellValue("Weeks");
            for (int i = 1; i <= Integer.parseInt(columns[columns.length - 1]); i++) {
                row = sheet2.createRow(i);
                Cell cell = row.createCell(0); // Create a cell in the first column
                cell.setCellValue("Week " + i); // Set the cell value
            }

            // Creating Sheet 3 (Performance)
            Sheet sheet3 = workbook.createSheet("Performance");
            row = sheet3.createRow(0);
            row.createCell(0).setCellValue("Test Item");
            row.createCell(1).setCellValue("Metric");
            // Writing headers for Performance
            for (int i = 2; i < 8; i++) {
                row.createCell(i).setCellValue("Stats");
            }
            row.createCell(8).setCellValue("Status");
            row.createCell(9).setCellValue("Date");

            // Creating Sheet 4 (Security)
            Sheet sheet4 = workbook.createSheet("Security");
            row = sheet4.createRow(0);
            String[] securityColumns = {
                    "Feature Name",
                    "Description",
                    "Test Scenario",
                    "Vulnerability Type",
                    "Severity",
                    "Affected Components",
                    "Repro",
                    "Recommendations",
                    "Priority",
                    "Status",
                    "Assignee",
                    "Deadline",
                    "Comments/Notes"
            };
            // Writing headers for Security
            for (int i = 0; i < securityColumns.length; i++) {
                row.createCell(i).setCellValue(securityColumns[i]);
            }

            // Adding hyperlink to the last cell
            row.createCell(12).setCellFormula("HYPERLINK(\"" + url + "\")");

            // Writing to file
            try (FileOutputStream fileOut = new FileOutputStream(filename)) {
                workbook.write(fileOut);
                System.out.println("Multi-sheet XLSX file created successfully: " + filename);
            }
        } catch (IOException e) {
            System.out.println("Error occurred while creating XLSX file: " + e.getMessage());
        }
    }
}
