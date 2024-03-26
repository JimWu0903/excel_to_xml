package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.soap.Node;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import static org.example.Constant.EXCEL_OUTPUT;
import static org.example.Constant.XML_FILE_LOCATION;

public class XmlToExcel {

    public static final String ELEMENT_NAME = "ReuMonReportN";

    public static void main(String[] args) {
        NodeList nodeList = extractDataFromXml(XML_FILE_LOCATION);
        writeDataToExcel(nodeList, EXCEL_OUTPUT);
    }

    static NodeList extractDataFromXml(String xmlFileLocation) {
        try {
            // 加載 XML 文件
            File xmlFile = new File(xmlFileLocation);
            FileInputStream fis = new FileInputStream(xmlFile);

            // 創建 DocumentBuilder 和 Document 對象以解析 XML
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(new InputSource(fis));

            // 獲取 XML 中存放資料的 node
            NodeList nodeList = doc.getElementsByTagName(ELEMENT_NAME);

            if (nodeList.getLength() == 0) {
                System.out.println("No data found in the XML file.");
                return null;
            }

            return nodeList;

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    public static void writeDataToExcel(NodeList nodeList, String excelOutputLocation) {
        try{
            // 創建 Excel 工作簿和工作表
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Data");

            List<String> columnHeaders = getColumnHeaders(nodeList);

            // // 寫入 Excel 表頭 (欄位名稱)
            Row headerRow = sheet.createRow(0);
            int columnIdx = 0;
            for (String header : columnHeaders) {
                Cell cell = headerRow.createCell(columnIdx++);
                cell.setCellValue(header);
            }

            // 遍歷每個 node 並將數據寫入 Excel 工作表
            int rowNum = 1;
            for (int i = 0; i < nodeList.getLength(); i++) {
                Row row = sheet.createRow(rowNum++);
                NodeList childNodes = nodeList.item(i).getChildNodes();
                for (int j = 0; j < childNodes.getLength(); j++) {
                    if (childNodes.item(j).getNodeType() == Node.ELEMENT_NODE) {
                        String nodeName = childNodes.item(j).getNodeName();
                        String nodeValue = childNodes.item(j).getTextContent();
                        int columnIndex = new ColumnIndex(columnHeaders, nodeName).getIndex();
                        Cell cell = row.createCell(columnIndex);
                        cell.setCellValue(nodeValue);
                    }
                }
            }

            // 將Excel 數據寫入文件
            FileOutputStream fos = new FileOutputStream(EXCEL_OUTPUT);
            workbook.write(fos);
            fos.close();
            workbook.close();

            System.out.println("Excel 文件已生成成功！");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static List<String> getColumnHeaders(NodeList nodeList) {
        List<String> columnHeaders = new ArrayList<>();
        NodeList firstNode = nodeList.item(0).getChildNodes();
        for (int i = 0; i < firstNode.getLength(); i++) {
            if (firstNode.item(i).getNodeType() == Node.ELEMENT_NODE) {
                String nodeName = firstNode.item(i).getNodeName();
                columnHeaders.add(nodeName);
            }
        }
        return columnHeaders;
    }

    /**
     * Helper class to get the index of a column header in a list of column headers.
     */
    static class ColumnIndex {
        private final List<String> columnHeaders;
        private final String nodeName;

        public ColumnIndex(List<String> columnHeaders, String nodeName) {
            this.columnHeaders = columnHeaders;
            this.nodeName = nodeName;
        }

        public int getIndex() {
            int index = 0;
            for (String header : columnHeaders) {
                if (header.equals(nodeName)) {
                    return index;
                }
                index++;
            }
            return -1; // Not found
        }
    }
}
