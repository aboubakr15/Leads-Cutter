package com.cutter.cutter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import javafx.scene.control.ButtonType;


import javafx.scene.control.Alert;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Splitter {
    private final Config config = new Config();

    public boolean splitAndSave(List<ExcelSheet> excelSheets, int divisor) throws IOException {
        int sheetsNotDone = 0;
        // Iterate over each ExcelSheet
        for (ExcelSheet excelSheet : excelSheets) {

            // Check if the divisor is greater than the companies "Will lead to empty sheets"
            int totalCompanies = excelSheet.getTotalCompanies();
            String workbookName = excelSheet.getFileName().replace(".xlsx", " ");

            if (totalCompanies < divisor) {
                saveCompanies(excelSheet.getHeadings(), excelSheet.getCenCompanies(), excelSheet.getEstCompanies(), excelSheet.getPacCompanies(), config.DEFAULT_SAVE_RESULT_DIR_1.getAbsolutePath(), workbookName, excelSheet);
                saveEmails(excelSheet.getHeadings(), excelSheet.getEmailCompanies(), config.DEFAULT_SAV_EMAILS_DIR.getAbsolutePath(), workbookName);
                moveOriginalWorkbook(excelSheet);
                continue;
            }
            // Extract time-zoned companies from the ExcelSheet
            List<Company> cenCompanies = excelSheet.getCenCompanies();
            List<Company> estCompanies = excelSheet.getEstCompanies();
            List<Company> pacCompanies = excelSheet.getPacCompanies();


            // Check if there are any companies to split
            if (cenCompanies.isEmpty() && estCompanies.isEmpty() && pacCompanies.isEmpty()) {
                Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
                alert.setTitle("No Companies");
                alert.setHeaderText(excelSheet.getFileName());
                alert.setContentText("There are no leads to split in the Excel sheet.\nDo you want to continue and save emails?");

                Optional<ButtonType> result = alert.showAndWait();
                if (result.isPresent() && result.get() == ButtonType.OK) {
                    // User chose to continue and save emails
                    saveEmails(excelSheet.getHeadings(), excelSheet.getEmailCompanies(), config.DEFAULT_SAV_EMAILS_DIR.getAbsolutePath(), workbookName);
                } else {
                    // User chose to cancel
                    sheetsNotDone++;
                    continue; // Move to the next ExcelSheet
                }
            }


            // Calculate chunk size for each group based on the divisor
            int cenChunkSize = cenCompanies.isEmpty() ? 0 : cenCompanies.size() / divisor;
            int estChunkSize = estCompanies.isEmpty() ? 0 : estCompanies.size() / divisor;
            int pacChunkSize = pacCompanies.isEmpty() ? 0 : pacCompanies.size() / divisor;

            // Calculate the number of remaining companies
            int cenRemaining = cenCompanies.size() % divisor;
            int estRemaining = estCompanies.size() % divisor;
            int pacRemaining = pacCompanies.size() % divisor;

            // Split cenCompanies
            List<List<Company>> splitCenCompanies = splitCompanies(cenCompanies, cenChunkSize, cenRemaining);

            // Split estCompanies
            List<List<Company>> splitEstCompanies = splitCompanies(estCompanies, estChunkSize, estRemaining);

            // Split pacCompanies
            List<List<Company>> splitPacCompanies = splitCompanies(pacCompanies, pacChunkSize, pacRemaining);

            for (int i = 0; i < divisor; i++) {
                List<Company> cenChunk = i < splitCenCompanies.size() ? splitCenCompanies.get(i) : new ArrayList<>();
                List<Company> estChunk = i < splitEstCompanies.size() ? splitEstCompanies.get(i) : new ArrayList<>();
                List<Company> pacChunk = i < splitPacCompanies.size() ? splitPacCompanies.get(i) : new ArrayList<>();
                switch (i) {
                    case 0 -> saveCompanies(excelSheet.getHeadings(), cenChunk, estChunk, pacChunk, config.DEFAULT_SAVE_RESULT_DIR_1.getAbsolutePath(), workbookName, excelSheet);
                    case 1 -> saveCompanies(excelSheet.getHeadings(), cenChunk, estChunk, pacChunk, config.DEFAULT_SAVE_RESULT_DIR_2.getAbsolutePath(), workbookName, excelSheet);
                    case 2 -> saveCompanies(excelSheet.getHeadings(), cenChunk, estChunk, pacChunk, config.DEFAULT_SAVE_RESULT_DIR_3.getAbsolutePath(), workbookName, excelSheet);
                }

            }

            saveEmails(excelSheet.getHeadings(), excelSheet.getEmailCompanies(), config.DEFAULT_SAV_EMAILS_DIR.getAbsolutePath(), workbookName);

            moveOriginalWorkbook(excelSheet);
        }
        return (sheetsNotDone == 0);
    }

    private List<List<Company>> splitCompanies(List<Company> companies, int chunkSize, int remaining) {
        // Split the companies into chunks
        List<List<Company>> dividedCompanies = new ArrayList<>();

        // Handle the case where chunkSize is 0 or 1
        if (chunkSize < 1) {
            dividedCompanies.add(new ArrayList<>(companies)); // Add the whole list as a single chunk
            return dividedCompanies;
        }

        int startIndex = 0;
        while (startIndex < companies.size()) {
            // Adjust the chunk size if there are remaining companies
            int adjustedChunkSize = chunkSize;
            if (remaining > 0) {
                adjustedChunkSize++;
                remaining--;
            }
            int endIndex = Math.min(startIndex + adjustedChunkSize, companies.size());
            dividedCompanies.add(companies.subList(startIndex, endIndex));
            startIndex = endIndex;
        }

        return dividedCompanies;
    }

    private void saveCompanies(List<String> headings, List<Company> cenCompanies, List<Company> estCompanies, List<Company> pacCompanies, String outputDirectory, String workbookName, ExcelSheet excelSheet) throws IOException {
        // Create a new workbook
        Workbook workbook = new XSSFWorkbook();

        // Create a sheet for the companies
        Sheet sheet = workbook.createSheet("Sheet 1");

        // Create the styled font for the headers
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short)11);
        headerFont.setColor(IndexedColors.BLACK.index);

        // Create cell style with the font
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(headerFont);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setFillForegroundColor(IndexedColors.YELLOW.index);

        CellStyle CB = workbook.createCellStyle();
        CB.setFont(headerFont);
        CB.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        CB.setFillForegroundColor(IndexedColors.PALE_BLUE.index);
        CB.setBorderTop(BorderStyle.THIN);
        CB.setBorderBottom(BorderStyle.THIN);
        CB.setBorderLeft(BorderStyle.THIN);
        CB.setBorderRight(BorderStyle.THIN);
        // Assuming newCellStyle is the cell style you want to modify to center the text
        CB.setAlignment(HorizontalAlignment.CENTER);
        // Set the vertical alignment as well if needed
        CB.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle RD = workbook.createCellStyle();
        RD.setFont(headerFont);
        RD.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        RD.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
        RD.setBorderTop(BorderStyle.THIN);
        RD.setBorderBottom(BorderStyle.THIN);
        RD.setBorderLeft(BorderStyle.THIN);
        RD.setBorderRight(BorderStyle.THIN);
        // Assuming newCellStyle is the cell style you want to modify to center the text
        RD.setAlignment(HorizontalAlignment.CENTER);
        // Set the vertical alignment as well if needed
        RD.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle VM = workbook.createCellStyle();
        VM.setFont(headerFont);
        VM.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        VM.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
        VM.setBorderTop(BorderStyle.THIN);
        VM.setBorderBottom(BorderStyle.THIN);
        VM.setBorderLeft(BorderStyle.THIN);
        VM.setBorderRight(BorderStyle.THIN);
        // Assuming newCellStyle is the cell style you want to modify to center the text
        VM.setAlignment(HorizontalAlignment.CENTER);
        // Set the vertical alignment as well if needed
        VM.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle DNC = workbook.createCellStyle();
        DNC.setFont(headerFont);
        DNC.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        DNC.setFillForegroundColor(IndexedColors.RED.index);
        DNC.setBorderTop(BorderStyle.THIN);
        DNC.setBorderBottom(BorderStyle.THIN);
        DNC.setBorderLeft(BorderStyle.THIN);
        DNC.setBorderRight(BorderStyle.THIN);
        // Assuming newCellStyle is the cell style you want to modify to center the text
        DNC.setAlignment(HorizontalAlignment.CENTER);
        // Set the vertical alignment as well if needed
        DNC.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle CD = workbook.createCellStyle();
        CD.setFont(headerFont);
        CD.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        CD.setFillForegroundColor(IndexedColors.SEA_GREEN.index);
        CD.setBorderTop(BorderStyle.THIN);
        CD.setBorderBottom(BorderStyle.THIN);
        CD.setBorderLeft(BorderStyle.THIN);
        CD.setBorderRight(BorderStyle.THIN);
        // Assuming newCellStyle is the cell style you want to modify to center the text
        CD.setAlignment(HorizontalAlignment.CENTER);
        // Set the vertical alignment as well if needed
        CD.setVerticalAlignment(VerticalAlignment.CENTER);


        CellStyle NA = workbook.createCellStyle();
        NA.setFont(headerFont);
        NA.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        NA.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.index);
        NA.setBorderTop(BorderStyle.THIN);
        NA.setBorderBottom(BorderStyle.THIN);
        NA.setBorderLeft(BorderStyle.THIN);
        NA.setBorderRight(BorderStyle.THIN);
        // Assuming newCellStyle is the cell style you want to modify to center the text
        NA.setAlignment(HorizontalAlignment.CENTER);
        // Set the vertical alignment as well if needed
        NA.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle WN = workbook.createCellStyle();
        WN.setFont(headerFont);
        WN.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        WN.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
        WN.setBorderTop(BorderStyle.THIN);
        WN.setBorderBottom(BorderStyle.THIN);
        WN.setBorderLeft(BorderStyle.THIN);
        WN.setBorderRight(BorderStyle.THIN);
        // Assuming newCellStyle is the cell style you want to modify to center the text
        WN.setAlignment(HorizontalAlignment.CENTER);
        // Set the vertical alignment as well if needed
        WN.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle PR = workbook.createCellStyle();
        PR.setFont(headerFont);
        PR.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        PR.setFillForegroundColor(IndexedColors.GOLD.index);
        PR.setBorderTop(BorderStyle.THIN);
        PR.setBorderBottom(BorderStyle.THIN);
        PR.setBorderLeft(BorderStyle.THIN);
        PR.setBorderRight(BorderStyle.THIN);
        // Assuming newCellStyle is the cell style you want to modify to center the text
        PR.setAlignment(HorizontalAlignment.CENTER);
        // Set the vertical alignment as well if needed
        PR.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle percent10 = workbook.createCellStyle();
        percent10.setFont(headerFont);
        percent10.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        percent10.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.index);
        percent10.setBorderTop(BorderStyle.THIN);
        percent10.setBorderBottom(BorderStyle.THIN);
        percent10.setBorderLeft(BorderStyle.THIN);
        percent10.setBorderRight(BorderStyle.THIN);
        // Assuming newCellStyle is the cell style you want to modify to center the text
        percent10.setAlignment(HorizontalAlignment.CENTER);
        // Set the vertical alignment as well if needed
        percent10.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle percent40 = workbook.createCellStyle();
        percent40.setFont(headerFont);
        percent40.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        percent40.setFillForegroundColor(IndexedColors.BLUE.index);
        percent40.setBorderTop(BorderStyle.THIN);
        percent40.setBorderBottom(BorderStyle.THIN);
        percent40.setBorderLeft(BorderStyle.THIN);
        percent40.setBorderRight(BorderStyle.THIN);
        // Assuming newCellStyle is the cell style you want to modify to center the text
        percent40.setAlignment(HorizontalAlignment.CENTER);
        // Set the vertical alignment as well if needed
        percent40.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle percent90 = workbook.createCellStyle();
        percent90.setFont(headerFont);
        percent90.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        percent90.setFillForegroundColor(IndexedColors.VIOLET.index);
        percent90.setBorderTop(BorderStyle.THIN);
        percent90.setBorderBottom(BorderStyle.THIN);
        percent90.setBorderLeft(BorderStyle.THIN);
        percent90.setBorderRight(BorderStyle.THIN);
        // Assuming newCellStyle is the cell style you want to modify to center the text
        percent90.setAlignment(HorizontalAlignment.CENTER);
        // Set the vertical alignment as well if needed
        percent90.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle sale = workbook.createCellStyle();
        sale.setFont(headerFont);
        sale.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        sale.setFillForegroundColor(IndexedColors.SEA_GREEN.index);
        sale.setBorderTop(BorderStyle.THIN);
        sale.setBorderBottom(BorderStyle.THIN);
        sale.setBorderLeft(BorderStyle.THIN);
        sale.setBorderRight(BorderStyle.THIN);
        // Assuming newCellStyle is the cell style you want to modify to center the text
        sale.setAlignment(HorizontalAlignment.CENTER);
        // Set the vertical alignment as well if needed
        sale.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle blank = workbook.createCellStyle();
        blank.setFont(headerFont);
        blank.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        blank.setFillForegroundColor(IndexedColors.WHITE.index);
        blank.setBorderTop(BorderStyle.THIN);
        blank.setBorderBottom(BorderStyle.THIN);
        blank.setBorderLeft(BorderStyle.THIN);
        blank.setBorderRight(BorderStyle.THIN);


        CellStyle[][] terminationColors = {
                                        {blank ,blank ,blank ,blank ,blank},
                                        {blank ,blank ,blank ,blank ,blank},
                                        {CB, blank ,blank ,blank ,blank},
                                        {RD, blank ,blank ,blank ,blank},
                                        {VM, blank ,blank ,blank ,blank},
                                        {DNC, blank ,blank ,blank ,blank},
                                        {PR, blank ,blank ,blank ,blank},
                                        {CD, blank ,blank ,blank ,blank},
                                        {NA, blank ,blank ,blank ,blank},
                                        {WN, blank ,blank ,blank ,blank},
                                        {blank ,blank ,blank ,blank ,blank},
                                        {blank ,blank ,blank ,blank ,blank},
                                        {percent10, blank ,blank ,blank ,blank},
                                        {percent40, blank ,blank ,blank ,blank},
                                        {percent90, blank ,blank ,blank ,blank},
                                        {sale, blank ,blank ,blank ,blank}
        };


        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headings.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headings.get(i));
            cell.setCellStyle(headerStyle);
        }

        // Write companies to the sheet
        int rowNum = 1;
        for (Company company : cenCompanies) {
            Row row = sheet.createRow(rowNum++);
            // Write company data to cells in the row
            row.createCell(0).setCellValue(company.getName());
            row.createCell(1).setCellValue(company.getNumber());
            row.createCell(2).setCellValue(company.getTimeZone());
            row.createCell(3).setCellValue(company.getDirect());
            row.createCell(4).setCellValue("");
            row.createCell(5).setCellValue(company.getDmName());
            row.createCell(6).setCellValue("");
            row.createCell(7).setCellValue("");
            row.createCell(8).setCellValue("");
            row.createCell(9).setCellValue("");
        }

        for (Company company : estCompanies) {
            Row row = sheet.createRow(rowNum++);
            // Write company data to cells in the row
            row.createCell(0).setCellValue(company.getName());
            row.createCell(1).setCellValue(company.getNumber());
            row.createCell(2).setCellValue(company.getTimeZone());
            row.createCell(3).setCellValue(company.getDirect());
            row.createCell(4).setCellValue("");
            row.createCell(5).setCellValue(company.getDmName());
            row.createCell(6).setCellValue("");
            row.createCell(7).setCellValue("");
            row.createCell(8).setCellValue("");
            row.createCell(9).setCellValue("");
        }

        for (Company company : pacCompanies) {
            Row row = sheet.createRow(rowNum++);
            // Write company data to cells in the row
            row.createCell(0).setCellValue(company.getName());
            row.createCell(1).setCellValue(company.getNumber());
            row.createCell(2).setCellValue(company.getTimeZone());
            row.createCell(3).setCellValue(company.getDirect());
            row.createCell(4).setCellValue("");
            row.createCell(5).setCellValue(company.getDmName());
            row.createCell(6).setCellValue("");
            row.createCell(7).setCellValue("");
            row.createCell(8).setCellValue("");
            row.createCell(9).setCellValue("");
        }


        int startRow = 5; // Row index for K6
        int endRow = 30; // Row index for O21
        int startCell = 10; // Column index for column K
        int endCell = 14; // Column index for column O
        for (int rowNum1 = startRow; rowNum1 <= endRow && rowNum1 - startRow < excelSheet.terminationCodes.size(); rowNum1++) {
            Row newRow = sheet.getRow(rowNum1);
            if (newRow == null) {
                newRow = sheet.createRow(rowNum1); // Create a new row in the new sheet if it doesn't exist
            }

            ArrayList<String> rowData = excelSheet.terminationCodes.get(rowNum1 - startRow); // Adjust row index for terminationCodes

            if (rowData != null) {
                // Iterate over the termination codes in the row
                for (int cellNum = startCell; cellNum <= endCell && cellNum - startCell < rowData.size(); cellNum++) {
                    // Get the corresponding style from terminationColors array
                    CellStyle style = terminationColors[rowNum1 - startRow][cellNum - startCell];

                    // Apply the style to the cell if it's not blank
                    Cell newCell = newRow.createCell(cellNum); // Create a new cell in the new row
                    newCell.setCellValue(rowData.get(cellNum - startCell)); // Set the termination code value
                    newCell.setCellStyle(style);
                }
            }
        }

        // Save the workbook to a file
        String fileName = workbookName + ".xlsx";
        try (FileOutputStream fileOut = new FileOutputStream(outputDirectory + File.separator + fileName)) {
            workbook.write(fileOut);
        }
        // Close the workbook
        workbook.close();
    }

    private void moveOriginalWorkbook(ExcelSheet originalWorkbook) throws IOException {

        // Define the destination file in the "Done" folder
        Path fromPath = new File(originalWorkbook.getFilePath()).toPath();
        Path cutDestinationPath = new File(config.DEFAULT_INVENTORY_DIR, originalWorkbook.getFileName()).toPath();
        Path copyDestinationPathX = new File(config.DEFAULT_SAVE_BASE_DIR_X, originalWorkbook.getFileName()).toPath();
        Path copyDestinationPathN = new File(config.DEFAULT_SAVE_BASE_DIR_N, originalWorkbook.getFileName()).toPath();

        if (originalWorkbook.getType().equals("XSHOW")) {
            // Copy the original workbook to the "XDone" folder
            Files.copy(fromPath, copyDestinationPathX, StandardCopyOption.REPLACE_EXISTING);
        }
        else if (originalWorkbook.getType().equals("N")) {
            // Copy the original workbook to the "NDone" folder
            Files.copy(fromPath, copyDestinationPathN, StandardCopyOption.REPLACE_EXISTING);
        }
        // Move the original workbook to the "Done" folder
        Files.move(fromPath, cutDestinationPath, StandardCopyOption.REPLACE_EXISTING);
    }

    private void saveEmails(List<String> headings, List<Company> companies, String outputDirectory, String workbookName) throws IOException {
        // Create a new workbook
        Workbook workbook = new XSSFWorkbook();

        // Create a sheet for the companies
        Sheet sheet = workbook.createSheet("Sheet 1");

        // Create the styled font for the headers
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short)12);
        headerFont.setColor(IndexedColors.BLACK.index);

        // Create cell style with the font
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(headerFont);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);

        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headings.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headings.get(i));
            cell.setCellStyle(headerStyle);
        }

        // Write companies to the sheet
        int rowNum = 1;
        for (Company company : companies) {
            Row row = sheet.createRow(rowNum++);
            // Write company data to cells in the row
            row.createCell(0).setCellValue(company.getName());
            row.createCell(1).setCellValue("");
            row.createCell(2).setCellValue("");
            row.createCell(3).setCellValue("");
            row.createCell(4).setCellValue(company.getEmail());
            row.createCell(5).setCellValue("");
            row.createCell(6).setCellValue("");
            row.createCell(7).setCellValue("");
            row.createCell(8).setCellValue("");
            row.createCell(9).setCellValue("");
        }

        // Save the workbook to a file
        String fileName = workbookName + ".xlsx";
        try (FileOutputStream fileOut = new FileOutputStream(outputDirectory + File.separator + fileName)) {
            workbook.write(fileOut);
        }
        // Close the workbook
        workbook.close();

    }

}
