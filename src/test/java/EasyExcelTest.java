import com.alibaba.excel.EasyExcel;
import excel.CustomSheetWriteHandler;
import excel.StyleWriteHandler;
import org.junit.Test;

import java.io.File;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @author by liangzj
 * @since 2022/9/17 15:32
 */
public class EasyExcelTest {

    @Test
    public void testWriteExcel() {
        String pathname = "E:\\liangzj\\Desktop\\test.xlsx";
        EasyExcel.write(new File(pathname))
                .head(header())
                .registerWriteHandler(new StyleWriteHandler())
                .registerWriteHandler(new CustomSheetWriteHandler())
                .sheet("Sheet1")
                .doWrite(data());
    }

    /**
     * 数据是先行后列
     *
     * @return
     */
    public List<List<String>> data() {
        List<List<String>> data = getRowColMatrix(3, 5);

        data.get(0).set(0, "用户1");
        data.get(0).set(1, "1234567890");
        data.get(0).set(2, "合同1");
        data.get(0).set(3, "文本1");
        data.get(0).set(4, "210283202209078615");

        data.get(1).set(0, "用户2");
        data.get(1).set(1, "1234553478");
        data.get(1).set(2, null);
        data.get(1).set(3, "文本2");
        data.get(1).set(4, "210211202209073951");

        data.get(2).set(0, "用户3");
        data.get(2).set(1, "8332675567");
        data.get(2).set(2, null);
        data.get(2).set(3, "文本3");
        data.get(2).set(4, "120221202209076790");

        return data;
    }

    /**
     * 表头是先列后行
     *
     * @return
     */
    public List<List<String>> header() {
        List<List<String>> header = getColRowMatrix(2, 5);
        header.get(0).set(0, "姓名");
        header.get(0).set(1, "姓名");
        header.get(1).set(0, "手机/邮箱");
        header.get(1).set(1, "手机/邮箱");
        header.get(2).set(0, "合同名称");
        header.get(2).set(1, "合同名称");
        header.get(3).set(0, "文件1");
        header.get(3).set(1, "单行文本");
        header.get(4).set(0, "文件1");
        header.get(4).set(1, "身份证号");

        return header;
    }

    /**
     * 生成一个先行后列的矩阵数组
     *
     * @param maxRow
     * @param maxCol
     * @return
     */
    private static List<List<String>> getRowColMatrix(int maxRow, int maxCol) {
        List<List<String>> header =
                Stream.generate(
                                () ->
                                        Stream.generate(() -> "")
                                                .limit(maxCol)
                                                .collect(Collectors.toList()))
                        .limit(maxRow)
                        .collect(Collectors.toList());
        return header;
    }

    /**
     * 生成一个先列后行的矩阵数组
     *
     * @param maxRow
     * @param maxCol
     * @return
     */
    private static List<List<String>> getColRowMatrix(int maxRow, int maxCol) {
        return getRowColMatrix(maxCol, maxRow);
    }
}
