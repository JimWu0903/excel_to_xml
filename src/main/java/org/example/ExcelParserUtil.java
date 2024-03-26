package org.example;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

import static org.example.Constant.RECYCLE_FILE_LOCATION;

public class ExcelParserUtil {

    public static List<recycleData> ExcelParser(String fileLocation) throws Exception {
        Path path = Paths.get(RECYCLE_FILE_LOCATION);

        XSSFWorkbook workbook = new XSSFWorkbook(path.toFile());
        XSSFSheet sheet = workbook.getSheetAt(2);

        int lastRowNum = sheet.getLastRowNum();
        String month = getMonth(sheet).replace("月","");;

        List<recycleData> dataList = new ArrayList<>();

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

                        recycleData data = new recycleData();
                        data.setCarNo(carNo);
                        data.setMonth(month);

                        dataList.add(data);
                    }

                    break;
                case 2:
                    for (int cellNum = 2; cellNum < lastCellNum; cellNum++) {
                        dataList.get(cellNum - 2).setVendorNm(row.getCell(cellNum).toString());
                    }

                    break;
                default:
                    for (int cellNum = 2 ; cellNum < lastCellNum; cellNum++ ) {

                        String date = getDateOrWeek(row.getCell(0)).orElse(StringUtils.EMPTY);
                        String week = getDateOrWeek(row.getCell(1)).orElse(StringUtils.EMPTY);

                        if (StringUtils.isBlank(date) || StringUtils.isBlank(week)) {
                            continue;
                        }
                        if (row.getCell(cellNum) == null) {
                            continue;
                        }
                        double weight = getWeightFromCell(row.getCell(cellNum)).orElse(0.0);
                        if (weight == 0.0) {
                            continue;
                        }
                        dataList.get(cellNum - 2).setWeight(weight/1000); // 回報單位是公斤，xml單位是公噸
                        dataList.get(cellNum - 2).setDate(getDateData(month, date));
                    }
            }
        }

        return dataList;
    }

    public static String getDateData(String month, String date){
        // Combine month and date
        String combinedDate = month + "-" + date;

        // Format the combined date to yyyy-mm-dd
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        try {
            // Assuming year can be retrieved from another source (e.g., another column in the sheet)
            String year = "2024"; // Replace with your logic to get the year
            String formattedDate = sdf.format(sdf.parse(year + "-" + combinedDate));
            return formattedDate;
        } catch (ParseException e) {
            e.printStackTrace();
            return StringUtils.EMPTY;
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
        Optional<Double> weight = Optional.of(0.0);
        if (!Objects.isNull(cell)) {
            CellType type = cell.getCellType();


            switch (type) {
                case NUMERIC -> {
                    weight = Optional.of(cell.getNumericCellValue());
                }

                case STRING -> {
                    String value = cell.getStringCellValue();
                    if (value.matches("\\d+")) {
                        weight = Optional.of(Double.parseDouble(value));
                    }
                }
                default -> {
                }

            }
        }

        return weight;
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

    public static class recycleData {

        // 車號
        String carNo;

        // 廠商名稱
        String vendorNm;

        // 月份
        String month;

        // 清運日期
        String date;
        // 清運
        double weight;

        public String getCarNo() {
            return carNo;
        }

        public void setCarNo(String carNo) {
            this.carNo = carNo;
        }

        public String getVendorNm() {
            return vendorNm;
        }

        public void setVendorNm(String vendorNm) {
            this.vendorNm = vendorNm;
        }

        public String getMonth() {
            return month;
        }

        public void setMonth(String month) {
            this.month = month;
        }

        public String getDate() {
            return date;
        }

        public void setDate(String date) {
            this.date = date;
        }

        public double getWeight() {
            return weight;
        }

        public void setWeight(double weight) {
            this.weight = weight;
        }
    }

    static class Data {

        // 車號
        String carNo;

        // 廠商名稱
        String vendorNm;

        // 月份
        String month;

        // 清運日期
        String date;

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
