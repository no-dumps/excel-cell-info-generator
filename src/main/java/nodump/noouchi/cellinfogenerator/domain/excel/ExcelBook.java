package nodump.noouchi.cellinfogenerator.domain.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

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
        return IntStream.range(0, workbook.getNumberOfSheets()).mapToObj(i -> workbook.getSheetName(i)).collect(Collectors.toList());
    }

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }
}
