package nodump.noouchi.cellinfogenerator.domain;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelSheet {
    private XSSFSheet sheet;

    public ExcelSheet(XSSFWorkbook workbook, String sheetName) {
        sheet = workbook.getSheet(sheetName);
    }

    public XSSFSheet getSheet() {
        return sheet;
    }
}
