module com.bsa {
    requires javafx.controls;
    requires javafx.fxml;

    opens com.bsa to javafx.fxml;
    exports com.bsa;
}
