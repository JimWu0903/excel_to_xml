package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.example.Constant.RECYCLE_FORM_LOCATION;
import static org.example.Constant.RECYCLE_REAL_DATA;


public class ExcelToExcel {

    public static final List<String> columnNames = Arrays.asList("再利用數量", "再利用作業完成日期", "清運(除)機具車號", "委託單位名稱");


    public static void main(String[] args) throws Exception {

        List<ExcelParserUtil.recycleData> result = ExcelParserUtil.ExcelParser(RECYCLE_REAL_DATA);
        readAndUpdateExcel(RECYCLE_FORM_LOCATION,result);
    }


    public static void readAndUpdateExcel(String excelFilePath, List<ExcelParserUtil.recycleData> result) {

        try {
            // 創建一個 FileInputStream 對象來讀取 Excel 文件
            FileInputStream fis = new FileInputStream(excelFilePath);

            // 創建一個 XSSFWorkbook 對象來讀取 Excel 工作簿
            XSSFWorkbook workbook = new XSSFWorkbook(fis);

            // 獲取第一個工作表
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            Map<String, Integer> columnHeaders = getColumnMapper(headerRow);

            for (int i = 0; i < result.size(); i++) {
                // 從第二行開始創建新的行
                Row row = sheet.createRow(i+1);
                final int count = i;
                columnHeaders.forEach((k, v) -> {
                    // Excel只有row0 表頭，因此 cell 都要新增
                    Cell cell = row.createCell(v);
                    switch (k) {
                        case "再利用數量":
                            cell.setCellValue(result.get(count).getWeight());
                            break;
                        case "再利用作業完成日期":
                            cell.setCellValue(result.get(count).getDate());
                            break;
                        case "清運(除)機具車號":
                            cell.setCellValue(result.get(count).getCarNo());
                            break;
                        case "委託單位名稱":
                            cell.setCellValue(result.get(count).getVendorNm());
                            break;
                    }
                });
            }

            // 將更新後的 Excel 工作簿寫入文件
            FileOutputStream fos = new FileOutputStream(excelFilePath);
            workbook.write(fos);
            fos.close();
            workbook.close();

            System.out.println("Excel 文件已更新成功！");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static Map<String, Integer> getColumnMapper(Row headerRow) {
        Map<String, Integer> columnIndexMap = new HashMap<>();
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            // 獲取當前單元格
            Cell cell = headerRow.getCell(i);
            for (String columnName : columnNames) {
                if(columnName.equals(cell.getStringCellValue())) {
                    columnIndexMap.put(columnName, i);
                    break;
                }
            }
        }
        return columnIndexMap;
    }

}
