module com.cutter.cutter {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;
    requires com.fasterxml.jackson.databind;

    opens com.cutter.cutter to javafx.fxml;
    exports com.cutter.cutter;
}