package nodump.noouchi.cellinfogenerator.application;

import com.fasterxml.jackson.core.JsonProcessingException;
import nodump.noouchi.cellinfogenerator.domain.ExcelBook;
import nodump.noouchi.cellinfogenerator.domain.ExcelSheet;
import nodump.noouchi.cellinfogenerator.domain.ExcelSheetService;

import java.util.List;

public class CellInfoGenerator {
    public static void GenerateExcelInfo(String path) throws JsonProcessingException {
        ExcelBook excelBook = new ExcelBook(path);
        List<String> sheetNames = excelBook.getSheetNames();
        System.out.print(sheetNames);
        var workbook = excelBook.getWorkbook();
        for (String sheetName :
                sheetNames
        ) {
            ExcelSheet excelSheet = new ExcelSheet(workbook, sheetName);
            String cellInfo = ExcelSheetService.getAllCellInfo(excelSheet.getSheet());
            System.out.print(cellInfo);
        }
    }
}