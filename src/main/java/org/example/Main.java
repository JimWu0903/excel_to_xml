package org.example;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.blueprint.FoodInfo;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Text;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Main {

    //https://www.baeldung.com/java-convert-excel-data-into-list

    private static final String FILE_LOCATION = "/Users/jim.wu/Downloads/222.xlsx";
    private static final String XML_FILE_LOCATION = "/Users/jim.wu/Downloads/222.xml";

    public static void main(String[] args) throws IOException {

        System.out.printf("Hello and welcome!");
        List<FoodInfo> result = excelDataToListOfObjets(FILE_LOCATION);
        writeFoodInfoListToXml(result, XML_FILE_LOCATION);
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

    public static void writeFoodInfoListToXml(List<FoodInfo> foodInfoList, String xmlFilePath) {
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document document = builder.newDocument();

            // Create root element,名稱自定義
            Element rootElement = document.createElement("FoodInfoList");
            document.appendChild(rootElement);

            for (FoodInfo foodInfo : foodInfoList) {
                // 每一 row 的 root,名稱自定義
                Element foodElement = document.createElement("row");
                rootElement.appendChild(foodElement);

                // Add category
                Element categoryElement = document.createElement("Category");
                Text categoryText = document.createTextNode(foodInfo.getCategory());
                categoryElement.appendChild(categoryText);
                foodElement.appendChild(categoryElement);

                // Add name
                Element nameElement = document.createElement("Name");
                Text nameText = document.createTextNode(foodInfo.getName());
                nameElement.appendChild(nameText);
                foodElement.appendChild(nameElement);

                // Add measure
                Element measureElement = document.createElement("Measure");
                Text measureText = document.createTextNode(foodInfo.getMeasure());
                measureElement.appendChild(measureText);
                foodElement.appendChild(measureElement);

                // Add calories
                Element caloriesElement = document.createElement("Calories");
                Text caloriesText = document.createTextNode(String.valueOf(foodInfo.getCalories()));
                caloriesElement.appendChild(caloriesText);
                foodElement.appendChild(caloriesElement);
            }

            // Write to XML file
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            DOMSource source = new DOMSource(document);
            StreamResult result = new StreamResult(xmlFilePath);
            transformer.transform(source, result);

            System.out.println("FoodInfo data written to " + xmlFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}