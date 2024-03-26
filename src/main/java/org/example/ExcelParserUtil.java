package org.example;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

import static org.example.Constant.RECYCLE_FILE_LOCATION;

public class ExcelParserUtil {

    public static void main(String[] args) throws Exception {
        Path path = Paths.get(RECYCLE_FILE_LOCATION);

        XSSFWorkbook workbook = new XSSFWorkbook(path.toFile());
        XSSFSheet sheet = workbook.getSheetAt(2);

        int lastRowNum = sheet.getLastRowNum();
        String month = getMonth(sheet);

        List<Data> dataList = new ArrayList<>();

        // 從第二列開始讀取
        for (int rowNum = 1; rowNum <= lastRowNum; rowNum++) {
            XSSFRow row = sheet.getRow(rowNum);
            short lastCellNum = row.getLastCellNum();
            switch (rowNum) {
                case 0:
                    break;
                case 1:
                    for (int cellNum = 2; cellNum < lastCellNum; cellNum++) {
                        String carNo = row.getCell(cellNum).toString();

                        Data d = new Data();
                        d.carNo = carNo;
                        d.month = month;

                        dataList.add(d);
                    }

                    break;
                case 2:
                    for (int cellNum = 2; cellNum < lastCellNum; cellNum++) {
                        dataList.get(cellNum - 2).vendorNm = row.getCell(cellNum).toString();
                    }

                    break;
                default:
                    for (int cellNum = 2 ; cellNum < lastCellNum; cellNum++ ) {
                        String date = getDateOrWeek(row.getCell(0)).orElse(StringUtils.EMPTY);
                        String week = getDateOrWeek(row.getCell(1)).orElse(StringUtils.EMPTY);

                        if (StringUtils.isBlank(date) || StringUtils.isBlank(week)) {
                            continue;
                        }

                        double weight = getWeightFromCell(row.getCell(cellNum)).orElse(0.0);
                        dataList.get(cellNum - 2).setDetail(date, week, weight);
                    }
            }
        }

        dataList.forEach(Data::setSumWeight);
        for (Data data : dataList) {
            System.out.println(data.carNo + " " + data.vendorNm + " " + data.month + " " + data.sumWeight);
        }

    }


    /**
     * 取得日期, 星期儲存格的值
     *
     * @param cell 儲存格
     * @return 日期 或 星期
     */
    public static Optional<String> getDateOrWeek(XSSFCell cell) {
        if (!Objects.isNull(cell)) {
            CellType type = cell.getCellType();

            switch (type) {
                case NUMERIC -> {
                    return Optional.of(String.valueOf((int) cell.getNumericCellValue()));
                }

                case STRING -> {
                    return Optional.of(String.valueOf(cell.getStringCellValue()));
                }
                default -> {
                }

            }
        }

        return Optional.empty();
    }

    /**
     * 取得垃圾清運重量
     *
     * @param cell 儲存格
     * @return 清運重量
     */
    public static Optional<Double> getWeightFromCell(XSSFCell cell) {
        if (!Objects.isNull(cell)) {
            CellType type = cell.getCellType();

            switch (type) {
                case NUMERIC -> {
                    return Optional.of(cell.getNumericCellValue());
                }

                case STRING -> {
                    String value = cell.getStringCellValue();
                    if (value.matches("\\d+")) {
                        return Optional.of(Double.parseDouble(value));
                    }
                }
                default -> {
                }

            }
        }

        return Optional.empty();
    }

    /**
     * 取得月份
     *
     * @param sheet Excel工作表
     * @return 月份
     */
    public static String getMonth(XSSFSheet sheet) {
        return Optional.ofNullable(sheet.getRow(2))
                .stream()
                .map(row -> row.getCell(0))
                .map(XSSFCell::toString)
                .findFirst()
                .orElse(StringUtils.EMPTY);
    }

    static class Data {

        // 車號
        String carNo;

        // 廠商名稱
        String vendorNm;

        // 月份
        String month;

        // 清運合計
        double sumWeight;

        List<DataDetail> details = new ArrayList<>();

        /**
         * 計錄每日的清運重量
         *
         * @param week 星期
         * @param weight 清運重量
         */
        public void setDetail(String day, String week, double weight) {
            details.add(new DataDetail(day, week, weight));
        }

        void setSumWeight() {
            sumWeight = details.stream()
                    .mapToDouble(DataDetail::weight)
                    .sum();
        }

    }

    static class DataDetail {

        // 日期
        String date;

        // 星期
        String week;

        // 清運重量
        double weight;

        public DataDetail(String date, String week, double weight) {
            this.date = date;
            this.week = week;
            this.weight = weight;
        }

        public double weight() {
            return this.weight;
        }
    }
}
