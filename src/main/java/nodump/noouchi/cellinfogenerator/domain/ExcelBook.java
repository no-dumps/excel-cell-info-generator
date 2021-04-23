package nodump.noouchi.cellinfogenerator.domain;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelBook {
    private XSSFWorkbook workbook;

    public ExcelBook(String path) {
        if (path == null) {
            throw new IllegalArgumentException("エクセルファイルのパスが不正");
        }
        try {
            workbook = new XSSFWorkbook(new FileInputStream(path));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public List<String> getSheetNames() {
        List<String> sheetNames = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            sheetNames.add(workbook.getSheetName(i));
        }
        return sheetNames;
    }

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }
}
