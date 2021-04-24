package nodump.noouchi;

import nodump.noouchi.cellinfogenerator.application.CellInfoGenerator;

import java.io.IOException;

public class Main {

    public static void main(String[] args) {
        System.out.println("Start.");
        try {
            if (args.length == 0) {
                System.out.println("エクセルファイルパスを指定してください。");
                return;
            }
            CellInfoGenerator.GenerateExcelInfo(args[0]);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            System.out.println("End.");
        }
    }
}
