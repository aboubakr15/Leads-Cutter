package com.cutter.cutter;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.io.IOException;

public class Main extends Application {

    @Override
    public void start(Stage primaryStage) {
        try {
            FXMLLoader loader = new FXMLLoader(getClass().getResource("main-view.fxml"));
            Parent root = loader.load();

            MainViewController mainController = loader.getController();
            mainController.setPrimaryStage(primaryStage);

            Scene scene = new Scene(root);
            primaryStage.setScene(scene);
            primaryStage.setTitle("Leads Cutter");
            primaryStage.show();

//            primaryStage.setMaximized(true);

        } catch (IOException e) {
            System.out.println(e.getLocalizedMessage());
        }
    }

    public static void main(String[] args) {
        launch(args);
    }
}