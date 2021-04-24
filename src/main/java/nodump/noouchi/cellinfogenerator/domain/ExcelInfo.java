package nodump.noouchi.cellinfogenerator.domain;

import java.util.ArrayList;

public class ExcelInfo {
    public String SheetName;
    public ArrayList<ExcelCellInfo> CellInfo;

    public ExcelInfo(String sheetName, ArrayList<ExcelCellInfo> excelCellInfos) {
        this.SheetName = sheetName;
        this.CellInfo = excelCellInfos;
    }
}
