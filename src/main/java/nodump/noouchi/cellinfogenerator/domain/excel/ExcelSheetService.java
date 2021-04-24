package nodump.noouchi.cellinfogenerator.domain.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.Iterator;

public class ExcelSheetService {
    public static ArrayList<ExcelCellInfo> getAllCellInfo(XSSFSheet sheet) {
        ArrayList<ExcelCellInfo> excelCellInfoList = new ArrayList<>();

        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            XSSFRow row = (XSSFRow) rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                XSSFCell cell = (XSSFCell) cellIterator.next();
                ExcelCellInfo excelCellInfo = new ExcelCellInfo(cell);
                excelCellInfoList.add(excelCellInfo);
            }
        }
        return excelCellInfoList;
    }
}
