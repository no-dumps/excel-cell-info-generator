package nodump.noouchi.cellinfogenerator.domain;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.Iterator;

public class ExcelSheetService {
    public static String getAllCellInfo(XSSFSheet sheet) throws JsonProcessingException {
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
        ObjectMapper mapper = new ObjectMapper();
        String json = mapper.writeValueAsString(excelCellInfoList);
        return json;
    }
}
