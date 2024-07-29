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
    private static int count_sent;
    private static int count_clicked;
    private static int count_reported;
    private static int count_no_action;
    private static int count_bounced;

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
        logger.info("Enter the path to the CSV file:");
        String csvFilePath = scanner.nextLine();
        // String csvFilePath = "Phishing-Report-Generator\\report-generator\\deleteme.csv";
        String excelFilePath = "phishing_report.xlsx";

        List<Employee> clickedList = new ArrayList<>();
        List<Employee> reportedList = new ArrayList<>();
        List<Employee> noActionList = new ArrayList<>();
        

        try (BOMInputStream bomInputStream = new BOMInputStream(new FileInputStream(csvFilePath));
             InputStreamReader reader = new InputStreamReader(bomInputStream)) {
            Iterable<CSVRecord> records = CSVFormat.DEFAULT
                    .withHeader("First Name", "Last Name","Primary Email Opened", "Primary Clicked",
                     "Multi Click Event", "Email Address", "Reported", "Email Bounced", "Passed?", "Phishing Template")
                    .withFirstRecordAsHeader()
                    .parse(reader);

            for (CSVRecord record : records) {
                // parse data
                String firstname = record.get("First Name");
                String lastname = record.get("Last Name");
                boolean open = Boolean.parseBoolean(record.get("Primary Email Opened"));
                boolean click = Boolean.parseBoolean(record.get("Primary Clicked"));
                int multiClick = Integer.parseInt(record.get("Multi Click Event"));
                String email = record.get("Email Address");
                boolean reported = Boolean.parseBoolean(record.get("Reported"));
                boolean bounced = Boolean.parseBoolean(record.get("Email Bounced"));
                boolean pass = Boolean.parseBoolean(record.get("Passed?"));
                String template = record.get("Phishing Template");

                Employee employee = new Employee(firstname, lastname, email, open, click, multiClick, reported, bounced, pass, template);
                count_sent++;

                if (bounced) {
                    count_bounced++;
                } else {
                    if (click) {
                        clickedList.add(employee);
                        count_clicked++;
                    }
                    if (reported) {
                        reportedList.add(employee);
                        count_reported++;
                    }
                    if (!click && !reported) {
                        noActionList.add(employee);
                        count_no_action++;
                    }
                }
            }
        } catch (IOException e) {
            logger.error("Error reading CSV file", e);
        }

        // System.out.println("Email Sent: " + count_sent);
        // System.out.println("Email Bounced: " + count_bounced);
        System.out.println("Email Delivered: " + (count_sent - count_bounced));
        System.out.println("Clicked: " + count_clicked + " (%" + String.format("%.2f", 100.0*46.0/817) + ")");
        System.out.println("Reported: " + count_reported + " (Should manually add # of users reported via MS)");
        System.out.println("No Action: " + count_no_action + " (Should manually subtract # of users reported via MS)");

        try (Workbook workbook = new XSSFWorkbook()) {
            createSheet(workbook, "Clicked", clickedList);
            createSheet(workbook, "Reported", reportedList);
            createSheet(workbook, "No Action", reportedList);

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
        header.createCell(0).setCellValue("Firstname");
        header.createCell(1).setCellValue("Lastname");
        header.createCell(2).setCellValue("Email");
        header.createCell(3).setCellValue("Primary Email Opened");
        header.createCell(4).setCellValue("Primary Clicked");
        header.createCell(5).setCellValue("Multi Click Event");
        header.createCell(6).setCellValue("Reported");
        header.createCell(7).setCellValue("Email Bounced");
        header.createCell(8).setCellValue("Passed?");
        header.createCell(9).setCellValue("Phishing Template");

        for (Employee employee : employees) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(employee.getFirstname());
            row.createCell(1).setCellValue(employee.getLastname());
            row.createCell(2).setCellValue(employee.getEmail());
            row.createCell(3).setCellValue(employee.isOpen());
            row.createCell(4).setCellValue(employee.isClick());
            row.createCell(5).setCellValue(employee.getMultiClick());
            row.createCell(6).setCellValue(employee.isReported());
            row.createCell(7).setCellValue(employee.isBounced());
            row.createCell(8).setCellValue(employee.isPass());
            row.createCell(9).setCellValue(employee.getTemplate());
        }
    }
}
