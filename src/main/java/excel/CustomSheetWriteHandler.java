package excel;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import org.apache.poi.xssf.streaming.SXSSFSheet;

/**
 * @author by liangzj
 * @since 2022/9/17 16:08
 */
public class CustomSheetWriteHandler implements SheetWriteHandler {

    @Override
    public void beforeSheetCreate(
            WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {}

    @Override
    public void afterSheetCreate(
            WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        writeSheetHolder.getSheet().protectSheet("123456");
        //        writeWorkbookHolder.getWriteWorkbook().setPassword("123456");
        //        ((SXSSFSheet) writeSheetHolder.getSheet()).enableLocking();
        ((SXSSFSheet) writeSheetHolder.getSheet()).lockSelectLockedCells(true);
        //        ((SXSSFSheet) writeSheetHolder.getSheet()).lockSelectUnlockedCells(true);
    }
}
