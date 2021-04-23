package nodump.noouchi;

import com.fasterxml.jackson.core.JsonProcessingException;
import nodump.noouchi.cellinfogenerator.application.CellInfoGenerator;

public class Main {

    public static void main(String[] args) {
        try {
            CellInfoGenerator.GenerateExcelInfo("resource/test/test.xlsx");
        } catch (JsonProcessingException e) {
            e.printStackTrace();
        }
    }
}
