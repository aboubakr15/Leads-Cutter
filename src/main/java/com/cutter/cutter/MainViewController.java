package com.cutter.cutter;


import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;
import java.io.IOException;
import java.util.List;

public class MainViewController {
    public Stage primaryStage;

    private Config config;

    @FXML
    private ImageView imageView;

    @FXML
    private Button selectFilesButton;

    @FXML
    private void handleSelectFilesButtonClick(ActionEvent event) throws IOException {
        // Your code to open a FileChooser and handle selected files
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));

        List<File> selectedFiles = fileChooser.showOpenMultipleDialog(selectFilesButton.getScene().getWindow());

        if (selectedFiles != null) {
            FXMLLoader loader = new FXMLLoader(getClass().getResource("cutter-view.fxml"));
            Parent root = loader.load();

            CutterViewController CutterController = loader.getController();
            CutterController.initData(selectedFiles, primaryStage);

            Scene scene = new Scene(root);
            primaryStage.setScene(scene);
            primaryStage.show();
            primaryStage.setMaximized(true);

        } else {
            System.out.println("No files selected.");
        }
    }

    public void initialize() {
        Image image = new Image(getClass().getResourceAsStream("/images/file.png"));
        imageView.setImage(image);

        config = new Config();
    }

    public void setPrimaryStage(Stage primaryStage) {
        this.primaryStage = primaryStage;
    }
}