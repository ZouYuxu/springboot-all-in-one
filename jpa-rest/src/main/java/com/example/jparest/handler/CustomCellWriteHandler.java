package com.example.jparest.handler;

import com.alibaba.excel.util.BooleanUtils;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

@Slf4j
public class CustomCellWriteHandler implements CellWriteHandler {

    @Override
    public void afterCellDispose(CellWriteHandlerContext context) {
        log.info("start");
        Cell cell = context.getCell();
        Sheet sheet = cell.getSheet();
        if (cell.getRowIndex() == 0 && cell.getColumnIndex() == 0) {
            Row row = sheet.getRow(0);

            int columnIndex = cell.getColumnIndex();
//            copyCell(cell, row, columnIndex);

        }

        //        Sheet sheet = context.getWriteSheetHolder()
//        String stringCellValue = cell.getStringCellValue();
//        WriteSheetHolder writeSheetHolder = context.getWriteSheetHolder();
//        Sheet sheet = writeSheetHolder.getSheet();
////        sheet.shiftColumns();
//        // 这里可以对cell进行任何操作
        log.info("第{}行，第{}列写入完成。", cell.getRowIndex(), cell.getColumnIndex());
        if (BooleanUtils.isTrue(context.getHead()) && cell.getColumnIndex() == 0) {
//            CreationHelper createHelper = context.getWriteSheetHolder().getSheet().getWorkbook().getCreationHelper();
//            Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.URL);
//            hyperlink.setAddress("https://github.com/alibaba/easyexcel");
//            cell.setHyperlink(hyperlink);
        }
    }

    private static void copyCell(Cell cell, Row row, int columnIndex) {
        row.shiftCellsRight(columnIndex, row.getLastCellNum(), 1);
        String value = cell.getStringCellValue();
        CellStyle cellStyle = cell.getCellStyle();
        Cell last = row.createCell(columnIndex);
        log.info("第{}行，第{}列写入完成。", last.getRowIndex(), last.getColumnIndex());
        last.setCellValue(value);
        last.setCellStyle(cellStyle);
    }

}