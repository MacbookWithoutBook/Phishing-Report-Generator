package com.reli.PhishingReportGenerator;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.io.input.BOMInputStream;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStreamReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class App {
    private static final Logger logger = LogManager.getLogger(App.class);

    public static void main(String[] args) {
        // Scanner scanner = new Scanner(System.in);
        // logger.info("Enter the path to the CSV file:");
        // String csvFilePath = scanner.nextLine();
        String csvFilePath = "Phishing-Report-Generator\\report-generator\\testfile.csv";
        String excelFilePath = "phishing_report.xlsx";

        List<Employee> reportedPhishing = new ArrayList<>();

        try (BOMInputStream bomInputStream = new BOMInputStream(new FileInputStream(csvFilePath));
             InputStreamReader reader = new InputStreamReader(bomInputStream)) {
            Iterable<CSVRecord> records = CSVFormat.DEFAULT
                    .withHeader("First Name", "Last Name", "Email Address", "Reported")
                    .withFirstRecordAsHeader()
                    .parse(reader);

            for (CSVRecord record : records) {
                String firstname = record.get("First Name");
                String lastname = record.get("Last Name");
                String email = record.get("Email Address");
                boolean reported = Boolean.parseBoolean(record.get("Reported"));

                Employee employee = new Employee(firstname, lastname, email);

                if (reported) {
                    reportedPhishing.add(employee);
                }
            }
        } catch (IOException e) {
            logger.error("Error reading CSV file", e);
        }

        try (Workbook workbook = new XSSFWorkbook()) {
            createSheet(workbook, "Reported Phishing", reportedPhishing);

            try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
                workbook.write(fileOut);
                logger.info("Report generated: " + excelFilePath);
            }
        } catch (IOException e) {
            logger.error("Error writing Excel file", e);
        }
    }

    private static void createSheet(Workbook workbook, String sheetName, List<Employee> employees) {
        Sheet sheet = workbook.createSheet(sheetName);
        int rowNum = 0;
        Row header = sheet.createRow(rowNum++);
        header.createCell(0).setCellValue("First Name");
        header.createCell(1).setCellValue("Last Name");
        header.createCell(2).setCellValue("Email Address");

        for (Employee employee : employees) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(employee.getFirstname());
            row.createCell(1).setCellValue(employee.getLastname());
            row.createCell(2).setCellValue(employee.getEmail());
        }
    }
}
