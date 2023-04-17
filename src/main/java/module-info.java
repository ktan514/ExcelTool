module com.example.exceltool {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;


    opens com.example.exceltool to javafx.fxml;
    exports com.example.exceltool;
}