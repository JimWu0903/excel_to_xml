package org.example.business;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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

public class Recycle {

    public static List<RecycleInfo> excelDataToListOfObjets(String fileLocation)
            throws IOException {
        FileInputStream file = new FileInputStream(fileLocation);
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);
        List<RecycleInfo> recycleInfoLs = new ArrayList<>();
        DataFormatter dataFormatter = new DataFormatter();
        for (int n = 1; n < sheet.getPhysicalNumberOfRows(); n++) {
            Row row = sheet.getRow(n);
            RecycleInfo recycleInfo = new RecycleInfo();
            int i = row.getFirstCellNum();

            recycleInfo.setIsBuzEntity((int) row.getCell(i).getNumericCellValue());
            recycleInfo.setRequestAddress(dataFormatter.formatCellValue(row.getCell(++i)));
            recycleInfo.setCleanerCode(dataFormatter.formatCellValue(row.getCell(++i)));
            recycleInfo.setStoreAddress(dataFormatter.formatCellValue(row.getCell(++i)));

            recycleInfoLs.add(recycleInfo);
        }
        return recycleInfoLs;
    }


    public static void writeInfoListToXml(List<RecycleInfo> infoList, String xmlFilePath) {
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document document = builder.newDocument();

            // Create root element,名稱自定義
            Element rootElement = document.createElement("FoodInfoList");
            document.appendChild(rootElement);

            for (RecycleInfo info : infoList) {
                // 每一 row 的 root,名稱自定義
                Element foodElement = document.createElement("ReuMonReportN");
                rootElement.appendChild(foodElement);

                // Add category
                Element categoryElement = document.createElement("廢棄物來源是否為事業單位");
                Text categoryText = document.createTextNode(String.valueOf(info.getIsBuzEntity()));
                categoryElement.appendChild(categoryText);
                foodElement.appendChild(categoryElement);

                // Add name
                Element nameElement = document.createElement("委託單位地址");
                Text nameText = document.createTextNode(info.getRequestAddress());
                nameElement.appendChild(nameText);
                foodElement.appendChild(nameElement);

                // Add measure
                Element measureElement = document.createElement("清除者代碼");
                Text measureText = document.createTextNode(info.getCleanerCode());
                measureElement.appendChild(measureText);
                foodElement.appendChild(measureElement);

                // Add calories
                Element caloriesElement = document.createElement("貯存地點");
                Text caloriesText = document.createTextNode(info.getStoreAddress());
                caloriesElement.appendChild(caloriesText);
                foodElement.appendChild(caloriesElement);
            }

            // Write to XML file
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            //transformer.setOutputProperty(OutputKeys.ENCODING, "Big5"); // 指定编码
            DOMSource source = new DOMSource(document);
            StreamResult result = new StreamResult(xmlFilePath);
            transformer.transform(source, result);


            System.out.println("Info data written to " + xmlFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
