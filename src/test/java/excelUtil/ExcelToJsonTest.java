package excelUtil;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import static org.junit.jupiter.api.Assertions.*;

@SpringBootTest
class ExcelToJsonTest {

    @Autowired ExcelToJson excelToJson;

    @Test
    void exportJsonFileInBase10Test() {
        try {
            excelToJson.exportJsonFileInBase10();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void exportJsonFileInBinary() {
        try {
            excelToJson.exportJsonFileInBinary();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void exportCSVFileIncludeCombined() {
        try {
            excelToJson.exportCSVFileIncludeCombined();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}