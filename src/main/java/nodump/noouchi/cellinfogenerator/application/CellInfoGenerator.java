package nodump.noouchi.cellinfogenerator.application;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import nodump.noouchi.cellinfogenerator.domain.excel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.List;

public class CellInfoGenerator {
    public static void GenerateExcelInfo(String path) throws JsonProcessingException {
        ExcelBook excelBook = new ExcelBook(path);
        List<String> sheetNames = excelBook.getSheetNames();
        XSSFWorkbook workbook = excelBook.getWorkbook();
        for (String sheetName :
                sheetNames
        ) {
            ExcelSheet excelSheet = new ExcelSheet(workbook, sheetName);

            ArrayList<ExcelCellInfo> cellInfos = ExcelSheetService.getAllCellInfo(excelSheet.getSheet());
            ExcelInfo excelInfo = new ExcelInfo(sheetName, cellInfos);
            ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
            System.out.print(mapper.writeValueAsString(excelInfo));
        }
    }
}