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

        if (isHead) { // 设置表头属性
            headerStyle(cell);
        } else { // 设置数据行属性
            contentStyle(cell);
        }
    }

    private void contentStyle(Cell cell) {
        // 创建新的单元格样式
        CellStyle cellStyle = cell.getSheet().getWorkbook().createCellStyle();
        // 复制原单元格的样式
        cellStyle.cloneStyleFrom(cell.getCellStyle());
        /* !! 注意：这行就是解锁单元格的代码，locked == true为锁定，locked == false为不锁定 */
        cellStyle.setLocked(!StringUtils.isEmpty(cell.getStringCellValue()));
        // 把新的单元格样式设置为当前单元格样式
        cell.setCellStyle(cellStyle);

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
