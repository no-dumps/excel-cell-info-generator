package nodump.noouchi;

import nodump.noouchi.cellinfogenerator.application.CellInfoGenerator;

import java.io.IOException;

public class Main {

    public static void main(String[] args) {
        System.out.println("Start.");
        try {
            CellInfoGenerator.GenerateExcelInfo("resource/test/test.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            System.out.println("End.");
        }
    }
}
