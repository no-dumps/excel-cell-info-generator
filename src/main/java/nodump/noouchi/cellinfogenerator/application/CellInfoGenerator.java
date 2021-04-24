package nodump.noouchi.cellinfogenerator.application;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import nodump.noouchi.cellinfogenerator.domain.excel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class CellInfoGenerator {
    public static void GenerateExcelInfo(String path) throws IOException {
        System.out.println("Loading WorkBook：" + path + " .");
        ExcelBook excelBook = new ExcelBook(path);
        List<String> sheetNames = excelBook.getSheetNames();
        ArrayList<ExcelInfo> result = new ArrayList<>();
        XSSFWorkbook workbook = excelBook.getWorkbook();
        for (String sheetName :
                sheetNames
        ) {
            System.out.println("Sheet：" + sheetName + " .");
            ExcelSheet excelSheet = new ExcelSheet(workbook, sheetName);

            ArrayList<ExcelCellInfo> cellInfos = ExcelSheetService.getAllCellInfo(excelSheet.getSheet());
            ExcelInfo excelInfo = new ExcelInfo(sheetName, cellInfos);
            result.add(excelInfo);
        }
        if (result.size() > 0) {
            ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
            // TODO want to format list.
            createJsonFile(path, mapper.writerWithDefaultPrettyPrinter().writeValueAsString(result));
        } else {
            System.out.println("No output.");
        }
    }

    private static void createJsonFile(String path, String json) throws IOException {
        File file = new File(path);
        String outputPath = getFileName(file.getAbsolutePath()) + ".json";
        FileWriter fw = new FileWriter(outputPath, false);
        fw.write(json);
        fw.close();
        System.out.println("Output Json：" + outputPath);
    }

    private static String getFileName(final String fullPathString) {
        File file = new File(fullPathString);
        String basename = file.getName();
        return basename.substring(0, basename.lastIndexOf('.'));
    }
}