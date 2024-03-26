package org.example.business;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;

public class Recycle {

    public static List<Map<String, String>> excelDataToListOfObjets(String fileLocation)
            throws IOException {
        FileInputStream file = new FileInputStream(fileLocation);
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        List<Map<String, String>> data = new ArrayList<>();

        Iterator<Row> iterator = sheet.iterator();
        Row headerRow = iterator.next(); // Assuming the first row contains column headers

        while (iterator.hasNext()) {
            Row currentRow = iterator.next();
            Map<String, String> rowData = new LinkedHashMap<>();
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell headerCell = headerRow.getCell(i);
                Cell currentCell = currentRow.getCell(i);
                rowData.put(headerCell.getStringCellValue(), currentCell.getStringCellValue());
            }
            data.add(rowData);
        }

        return data;
    }


    public static void writeDataToXml(List<Map<String, String>> data, String xmlFilePath) throws IOException {
        BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(xmlFilePath), StandardCharsets.UTF_8));
        writer.write("<?xml version=\"1.0\" encoding=\"Big5\"?>\n");
        writer.write("<Working xmlns=\"urn:ReuSchema\">\n");
        for (Map<String, String> row : data) {
            writer.write("\t<ReuMonReportN>\n");
            for (Map.Entry<String, String> entry : row.entrySet()) {
//                String tagName = entry.getKey().replaceAll("[\\\\u4E00-\\\\u9FA5\\\\w\\\\s]", "");
                String tagName = entry.getKey();
                writer.write("\t\t<" + tagName + ">" + entry.getValue() + "</" + tagName + ">\n");
            }
            writer.write("\t</ReuMonReportN>\n");
        }
        writer.write("</Working>");
        writer.close();
    }
}
