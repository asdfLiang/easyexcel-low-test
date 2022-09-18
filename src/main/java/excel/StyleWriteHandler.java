package excel;

import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.util.StringUtils;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.List;

/**
 * @author by liangzj
 * @since 2022/9/17 16:26
 */
public class StyleWriteHandler extends LongestMatchColumnWidthStyleStrategy {

    @Override
    public void afterCellDispose(
            WriteSheetHolder writeSheetHolder,
            WriteTableHolder writeTableHolder,
            List<CellData> cellDataList,
            Cell cell,
            Head head,
            Integer relativeRowIndex,
            Boolean isHead) {
        super.afterCellDispose(
                writeSheetHolder,
                writeTableHolder,
                cellDataList,
                cell,
                head,
                relativeRowIndex,
                isHead);

        if (isHead) {
            headerStyle(cell);
        } else {
            contentStyle(cell);
        }
    }

    private void contentStyle(Cell cell) {
        CellStyle cellStyle = cell.getSheet().getWorkbook().createCellStyle();
        cellStyle.setLocked(!StringUtils.isEmpty(cell.getStringCellValue()));
        cell.setCellStyle(cellStyle);
        cell.getSheet().protectSheet("123");

        if (cell.getCellStyle().getLocked()) {
            Font font = cell.getSheet().getWorkbook().createFont();
            font.setColor(IndexedColors.GREY_40_PERCENT.getIndex());
            cell.getCellStyle().setFont(font);
        } else {
            cell.setCellValue("可填写");
        }
    }

    private static void headerStyle(Cell cell) {
        int colWidth = cell.getStringCellValue().length() * 1500;

        // 根据表头文字设置列宽
        cell.getSheet().setColumnWidth(cell.getColumnIndex(), colWidth);
        // 冻结表头
        cell.getSheet().createFreezePane(1, 2);
    }
}
