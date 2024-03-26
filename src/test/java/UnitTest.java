import org.example.ExcelParserUtil;
import org.example.ExcelToXml;
import org.example.blueprint.FoodInfo;
import org.junit.Test;

import java.io.IOException;
import java.util.List;

import static org.example.Constant.FOOD_FILE_LOCATION;
import static org.example.Constant.RECYCLE_FILE_LOCATION;
import static org.junit.Assert.assertEquals;

public class UnitTest {

    @Test
    public void whenParsingExcelFileWithApachePOI_thenConvertsToList() throws IOException {
        List<FoodInfo> foodInfoList = ExcelToXml.excelDataToListOfObjets(FOOD_FILE_LOCATION);

        assertEquals("1", foodInfoList.get(0).getCategory());
        assertEquals("row3-measure", foodInfoList.get(3).getMeasure());
    }

    @Test
    public void testExcelParser() throws Exception {
        List<ExcelParserUtil.recycleData> dataList = ExcelParserUtil.ExcelParser(RECYCLE_FILE_LOCATION);

        for (ExcelParserUtil.recycleData recycleData : dataList) {
            System.out.print(recycleData.getCarNo()+" ");
            System.out.print(recycleData.getMonth()+" ");
            System.out.print(recycleData.getVendorNm()+" ");
            System.out.print(recycleData.getDate()+" ");
            System.out.println(recycleData.getWeight());
        }

    }


}
