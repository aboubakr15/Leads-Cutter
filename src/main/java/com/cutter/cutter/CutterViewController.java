package com.cutter.cutter;

import javafx.beans.property.SimpleStringProperty;
import javafx.beans.value.ObservableValue;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;

public class CutterViewController implements Initializable {
    public Stage primaryStage;

    private Config config;

    @FXML
    private Button cutButton;

    @FXML
    private TextField workbooksCount;

    @FXML
    private TextField baseWorkbookDirN;

    @FXML
    private TextField baseWorkbookDirX;

    @FXML
    private TextField saveResultDir1;

    @FXML
    private TextField saveResultDir2;

    @FXML
    private TextField saveResultDir3;

    @FXML
    private TextField emailsDir;

    @FXML
    private TextField inventoryDir;

    @FXML
    private TableView<ExcelSheet> tableView;

    @FXML
    private TableColumn<ExcelSheet, String> fileNameColumn;

    @FXML
    private TableColumn<ExcelSheet, String> easternLeadsColumn;

    @FXML
    private TableColumn<ExcelSheet, String> pacificLeadsColumn;

    @FXML
    private TableColumn<ExcelSheet, String> centralLeadsColumn;

    private final ObservableList<ExcelSheet> excelSheets = FXCollections.observableArrayList();

    private final ArrayList<File> files = new ArrayList<>();

    void initData(List<File> files, Stage primaryStage) {
        this.primaryStage = primaryStage;

        for (File f : files) {
            this.files.add(f);
            try {
                gatherData(f);
            } catch (IOException e) {
                (new Alert(AlertType.ERROR, "Failed to get data from [" + f.getAbsolutePath() + "]")).show();
                System.out.println(e.getLocalizedMessage());
            }
        }
    }

    @Override
    public void initialize(URL url, ResourceBundle resourceBundle) {
        fileNameColumn.setCellValueFactory(new PropertyValueFactory<>("fileName"));
        easternLeadsColumn.setCellValueFactory(
                cellData -> new SimpleStringProperty(String.valueOf(cellData.getValue().getEstCompanies().size())));
        pacificLeadsColumn.setCellValueFactory(
                cellData -> new SimpleStringProperty(String.valueOf(cellData.getValue().getPacCompanies().size())));
        centralLeadsColumn.setCellValueFactory(
                cellData -> new SimpleStringProperty(String.valueOf(cellData.getValue().getCenCompanies().size())));

        tableView.setItems(excelSheets);

        this.config = new Config();

        // Initialize text fields with config values
        baseWorkbookDirN.setText(config.DEFAULT_SAVE_BASE_DIR_N.getAbsolutePath());
        baseWorkbookDirX.setText(config.DEFAULT_SAVE_BASE_DIR_X.getAbsolutePath());
        saveResultDir1.setText(config.DEFAULT_SAVE_RESULT_DIR_1.getAbsolutePath());
        saveResultDir2.setText(config.DEFAULT_SAVE_RESULT_DIR_2.getAbsolutePath());
        saveResultDir3.setText(config.DEFAULT_SAVE_RESULT_DIR_3.getAbsolutePath());
        emailsDir.setText(config.DEFAULT_SAV_EMAILS_DIR.getAbsolutePath());
        inventoryDir.setText(config.DEFAULT_INVENTORY_DIR.getAbsolutePath());
    }

    private void gatherData(File excelFile) throws IOException {
        // Read data from Excel File.

        // Get the Excel workbook
        FileInputStream file = new FileInputStream(excelFile.getAbsolutePath());
        Workbook workbook = new XSSFWorkbook(file);

        // Create a Formatter
        DataFormatter dataFormatter = new DataFormatter();

        ArrayList<Company> companies = new ArrayList<>();
        ArrayList<Company> emailCompanies = new ArrayList<>();
        ArrayList<String> headings = new ArrayList<>();

        // Get the first sheet
        Sheet sheet = workbook.getSheetAt(0);

        // Create an iterator to iterate over the rows
        for (Row row : sheet) {

            // Create an iterator to iterate over the cells
            Cell nameCell = row.getCell(0);
            String nameCellValue = dataFormatter.formatCellValue(nameCell);

            Cell numberCell = row.getCell(1);
            String numberCellValue = dataFormatter.formatCellValue(numberCell);

            Cell timeZoneCell = row.getCell(2);
            String timeZoneCellValue = dataFormatter.formatCellValue(timeZoneCell);

            Cell directCell = row.getCell(3);
            String directCellValue = dataFormatter.formatCellValue(directCell);

            Cell emailCell = row.getCell(4);
            String emailCellValue = dataFormatter.formatCellValue(emailCell);

            Cell dmNameCell = row.getCell(5);
            String dmNameCellValue = dataFormatter.formatCellValue(dmNameCell);

            Cell terminationCodeCell = row.getCell(6);
            String terminationCodeCellValue = dataFormatter.formatCellValue(terminationCodeCell);

            Cell DateCell =  row.getCell(7);
            String DateValue = String.valueOf((DateCell));

            Cell specialNotesCell = row.getCell(8);
            String specialNotesCellValue = dataFormatter.formatCellValue(specialNotesCell);

            Cell opportunitySystemCell = row.getCell(9);
            String opportunitySystemCellValue = dataFormatter.formatCellValue(opportunitySystemCell);


            if (row.getRowNum() == 0) {
                headings.add(nameCellValue);
                headings.add(numberCellValue);
                headings.add(timeZoneCellValue);
                headings.add(directCellValue);
                headings.add(emailCellValue);
                headings.add(dmNameCellValue);
                headings.add(terminationCodeCellValue);
                headings.add(DateValue);
                headings.add(specialNotesCellValue);
                headings.add(opportunitySystemCellValue);
                continue;
            }
            if (timeZoneCellValue.isEmpty()) {
                continue;
            }
            if (specialNotesCellValue.equals("new")) {
                Company company = new Company(nameCellValue, numberCellValue, timeZoneCellValue, directCellValue,
                        emailCellValue, dmNameCellValue, terminationCodeCellValue, specialNotesCellValue,
                        opportunitySystemCellValue, DateValue);

                companies.add(company);
//                if (!emailCellValue.isEmpty()) emailCompanies.add(company);
            }
        }

        // Check if there are not any new keyword
        if (companies.isEmpty()) {
            // Create an iterator to iterate over the rows
            for (Row row : sheet) {
                // Iterator over the cells
                Cell nameCell = row.getCell(0);
                String nameCellValue = dataFormatter.formatCellValue(nameCell);

                Cell numberCell = row.getCell(1);
                String numberCellValue = dataFormatter.formatCellValue(numberCell);

                Cell timeZoneCell = row.getCell(2);
                String timeZoneCellValue = dataFormatter.formatCellValue(timeZoneCell);

                Cell directCell = row.getCell(3);
                String directCellValue = dataFormatter.formatCellValue(directCell);

                Cell emailCell = row.getCell(4);
                String emailCellValue = dataFormatter.formatCellValue(emailCell);

                Cell dmNameCell = row.getCell(5);
                String dmNameCellValue = dataFormatter.formatCellValue(dmNameCell);

                Cell terminationCodeCell = row.getCell(6);
                String terminationCodeCellValue = dataFormatter.formatCellValue(terminationCodeCell);

                Cell DateCell = row.getCell(7);
                String DateValue = String.valueOf((DateCell));

                Cell specialNotesCell = row.getCell(8);
                String specialNotesCellValue = dataFormatter.formatCellValue(specialNotesCell);

                Cell opportunitySystemCell = row.getCell(9);
                String opportunitySystemCellValue = dataFormatter.formatCellValue(opportunitySystemCell);


                if (row.getRowNum() == 0) {
                    continue;
                }

                if (timeZoneCellValue.isEmpty()) {
                    break;
                } else {
                    Company company = new Company(nameCellValue, numberCellValue, timeZoneCellValue, directCellValue,
                            emailCellValue, dmNameCellValue, terminationCodeCellValue, specialNotesCellValue,
                            opportunitySystemCellValue, DateValue);

                    companies.add(company);
//                    if (!emailCellValue.isEmpty()) emailCompanies.add(company);
                }
            }
        }

        emailCompanies = getEmails(dataFormatter, sheet);

        System.out.println("THE SIZE:" + companies.size());

        workbook.close();

        ArrayList<Company> cenCompanies = new ArrayList<>();
        ArrayList<Company> estCompanies = new ArrayList<>();
        ArrayList<Company> pacCompanies = new ArrayList<>();

        for (var cmp : companies) {
            switch (cmp.getTimeZone().toLowerCase()) {
                case "cen" -> cenCompanies.add(cmp);
                case "est" -> estCompanies.add(cmp);
                case "pac" -> pacCompanies.add(cmp);
            }
        }

        // implement a method to get the terminationCode
        ArrayList<ArrayList<String>> terminationCodes = getTerminationCodes(dataFormatter,sheet);


        String excelFileName = excelFile.getName();
        String type = "";
        if (excelFileName.startsWith("X_") || excelFileName.startsWith("X _") || excelFileName.startsWith("X-") || excelFileName.startsWith("X -")) {
            type = "XSHOW";
        } else {type = "N";}

        ExcelSheet excelSheet = new ExcelSheet(excelFile.getAbsolutePath(), excelFile.getName(), type, headings, emailCompanies, cenCompanies,
                estCompanies, pacCompanies, terminationCodes);
        excelSheets.add(excelSheet);
    }

    private ArrayList<ArrayList<String>> getTerminationCodes(DataFormatter dataFormatter, Sheet excelSheet) {
        ArrayList<ArrayList<String>> terminationCodes = new ArrayList<>();
        for (int rowNum = 5; rowNum < 21; rowNum++) { // Loop through rows 6 to 21 (inclusive)
            Row row = excelSheet.getRow(rowNum);
            ArrayList<String> rowData = new ArrayList<>();
            if (row != null) {
                for (int cellNum = 10; cellNum < 15; cellNum++) { // Loop through columns K to O (inclusive)
                    Cell cell = row.getCell(cellNum);
                    if (cell != null) {
                        rowData.add(dataFormatter.formatCellValue(cell));
                    } else {
                        rowData.add(""); // or handle null cells as needed
                    }
                }
                terminationCodes.add(rowData);
            }
        }
        return terminationCodes;
    }

    private ArrayList<Company> getEmails(DataFormatter dataFormatter, Sheet sheet) {
        ArrayList<Company> emailCompanies = new ArrayList<>();
        // Create an iterator to iterate over the rows
        for (Row row : sheet) {
            // Iterator over the cells
            Cell nameCell = row.getCell(0);
            String nameCellValue = dataFormatter.formatCellValue(nameCell);

            Cell numberCell = row.getCell(1);
            String numberCellValue = dataFormatter.formatCellValue(numberCell);

            Cell timeZoneCell = row.getCell(2);
            String timeZoneCellValue = dataFormatter.formatCellValue(timeZoneCell);

            Cell directCell = row.getCell(3);
            String directCellValue = dataFormatter.formatCellValue(directCell);

            Cell emailCell = row.getCell(4);
            String emailCellValue = dataFormatter.formatCellValue(emailCell);

            Cell dmNameCell = row.getCell(5);
            String dmNameCellValue = dataFormatter.formatCellValue(dmNameCell);

            Cell terminationCodeCell = row.getCell(6);
            String terminationCodeCellValue = dataFormatter.formatCellValue(terminationCodeCell);

            Cell specialNotesCell = row.getCell(7);
            String specialNotesCellValue = dataFormatter.formatCellValue(specialNotesCell);

            Cell opportunitySystemCell = row.getCell(8);
            String opportunitySystemCellValue = dataFormatter.formatCellValue(opportunitySystemCell);

            Cell DateCell = row.getCell(9);
            String DateValue = String.valueOf((DateCell));

            if (row.getRowNum() == 0 || emailCellValue.isEmpty()) continue;

            Company company = new Company(nameCellValue, numberCellValue, timeZoneCellValue, directCellValue,
                    emailCellValue, dmNameCellValue, terminationCodeCellValue, specialNotesCellValue,
                    opportunitySystemCellValue, DateValue);
            emailCompanies.add(company);
        }
        return emailCompanies;
    }
    @FXML
    public void onSelectBaseWorkbookDirN(ActionEvent event) {
        configureAndSaveDirectory(baseWorkbookDirN, "DEFAULT_SAVE_BASE_DIR_N");
    }

    @FXML
    public void onSelectBaseWorkbookDirX(ActionEvent event) {
        configureAndSaveDirectory(baseWorkbookDirX, "DEFAULT_SAVE_BASE_DIR_X");
    }

    @FXML
    public void onSelectResultDir1(ActionEvent event) {
        configureAndSaveDirectory(saveResultDir1, "DEFAULT_SAVE_RESULT_DIR_1");
    }

    @FXML
    public void onSelectResultDir2(ActionEvent event) {
        configureAndSaveDirectory(saveResultDir2, "DEFAULT_SAVE_RESULT_DIR_2");
    }

    @FXML
    public void onSelectResultDir3(ActionEvent event) {
        configureAndSaveDirectory(saveResultDir3, "DEFAULT_SAVE_RESULT_DIR_3");
    }

    @FXML
    public void onSelectEmailDir(ActionEvent event) {
        configureAndSaveDirectory(emailsDir, "DEFAULT_SAV_EMAILS_DIR");
    }

    @FXML
    public void onSelectInventoryDir(ActionEvent event) {
        configureAndSaveDirectory(inventoryDir, "DEFAULT_INVENTORY_DIR");
    }

    private void configureAndSaveDirectory(TextField textField, String configField) {
        DirectoryChooser directoryChooser = new DirectoryChooser();
        directoryChooser.setTitle("Select Directory");

        File selectedDirectory = directoryChooser.showDialog(primaryStage);

        if (selectedDirectory != null) {
            textField.setText(selectedDirectory.getAbsolutePath());
            updateConfig(configField, selectedDirectory.getAbsolutePath());
        }
    }

    private void updateConfig(String configField, String value) {
        switch (configField) {
            case "DEFAULT_SAVE_BASE_DIR_N" -> config.DEFAULT_SAVE_BASE_DIR_N = new File(value);
            case "DEFAULT_SAVE_BASE_DIR_X" -> config.DEFAULT_SAVE_BASE_DIR_X = new File(value);
            case "DEFAULT_SAVE_RESULT_DIR_1" -> config.DEFAULT_SAVE_RESULT_DIR_1 = new File(value);
            case "DEFAULT_SAVE_RESULT_DIR_2" -> config.DEFAULT_SAVE_RESULT_DIR_2 = new File(value);
            case "DEFAULT_SAVE_RESULT_DIR_3" -> config.DEFAULT_SAVE_RESULT_DIR_3 = new File(value);
            case "DEFAULT_SAV_EMAILS_DIR" -> config.DEFAULT_SAV_EMAILS_DIR = new File(value);
            case "DEFAULT_INVENTORY_DIR" -> config.DEFAULT_INVENTORY_DIR = new File(value);
        }
        config.save();
    }

    @FXML
    public void onCut(ActionEvent event) {
        // Check if the entered value is valid
        try {
            Splitter splitter = new Splitter();
            boolean success = splitter.splitAndSave(excelSheets.stream().toList(), 3);

            if (success) {
                Alert alert = new Alert(AlertType.INFORMATION);
                alert.setTitle("Splitting Complete");
                alert.setHeaderText(null);
                alert.setContentText("Splitting and saving operation completed successfully.");
                alert.showAndWait();

                // Navigate to main window
                FXMLLoader loader = new FXMLLoader(getClass().getResource("main-view.fxml"));
                loader.load();
                MainViewController mainViewController = loader.getController();
                mainViewController.setPrimaryStage(primaryStage);
                primaryStage.setScene(new Scene(loader.getRoot()));
                primaryStage.show();
            }

        } catch (NumberFormatException e) {

            Alert alert = new Alert(AlertType.ERROR);
            alert.setTitle("Input Error");
            alert.setHeaderText("Invalid Input");
            alert.setContentText("Please enter a valid INTEGER value.");
            alert.showAndWait();

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}