package org.example;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.blueprint.FoodInfo;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Main {

    //https://www.baeldung.com/java-convert-excel-data-into-list

    private static final String FILE_LOCATION = "/Users/jim.wu/Downloads/222.xlsx";

    public static void main(String[] args) throws IOException {

        System.out.printf("Hello and welcome!");
        List<FoodInfo> result = excelDataToListOfObjets(FILE_LOCATION);
        for (FoodInfo foodInfo : result) {
            System.out.println(foodInfo.getCategory() + " " + foodInfo.getName() + " " + foodInfo.getMeasure() + " " + foodInfo.getCalories());
        }
    }

    public static List<FoodInfo> excelDataToListOfObjets(String fileLocation)
            throws IOException {
        FileInputStream file = new FileInputStream(fileLocation);
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);
        List<FoodInfo> foodData = new ArrayList<>();
        DataFormatter dataFormatter = new DataFormatter();
        for (int n = 1; n < sheet.getPhysicalNumberOfRows(); n++) {
            Row row = sheet.getRow(n);
            FoodInfo foodInfo = new FoodInfo();
            int i = row.getFirstCellNum();

            foodInfo.setCategory(dataFormatter.formatCellValue(row.getCell(i)));
            foodInfo.setName(dataFormatter.formatCellValue(row.getCell(++i)));
            foodInfo.setMeasure(dataFormatter.formatCellValue(row.getCell(++i)));
            foodInfo.setCalories(row.getCell(++i).getNumericCellValue());

            foodData.add(foodInfo);
        }
        return foodData;
    }
}