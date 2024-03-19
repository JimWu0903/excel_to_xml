import org.example.ExcelToXml;
import org.example.blueprint.FoodInfo;
import org.junit.Test;

import java.io.IOException;
import java.util.List;

import static org.junit.Assert.assertEquals;

public class UnitTest {

    @Test
    public void whenParsingExcelFileWithApachePOI_thenConvertsToList() throws IOException {
        List<FoodInfo> foodInfoList = ExcelToXml.excelDataToListOfObjets(ExcelToXml.FILE_LOCATION);

        assertEquals("row1-0", foodInfoList.get(0).getCategory());
        assertEquals("row3-measure", foodInfoList.get(3).getMeasure());
    }
}
