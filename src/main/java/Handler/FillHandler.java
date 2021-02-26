package Handler;

import com.alibaba.excel.analysis.v07.handlers.CellInlineStringValueTagHandler;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

public class FillHandler implements CellWriteHandler
{
    private static final Logger LOGGER = LoggerFactory.getLogger(CustomCellWriteHandler.class);

    public void beforeCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Head head, Integer integer,
            Integer integer1, Boolean aBoolean) {

    }

    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Cell cell, Head head, Integer integer,
            Boolean aBoolean) {

    }

    public void afterCellDataConverted(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, CellData cellData, Cell cell, Head head,
            Integer integer, Boolean aBoolean) {

    }

    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> list, Cell cell, Head head,
            Integer integer, Boolean aBoolean) {

        Workbook  workbook = writeSheetHolder.getSheet().getWorkbook();


        CellStyle cellStyle = workbook.createCellStyle();

        if (cell.getRowIndex() == 3 && 1 == cell.getColumnIndex()) {
            System.out.println("Row is " +  cell.getRowIndex() + "And Cell is " + cell.getColumnIndex());
            int actualCellRowNum = cell.getRowIndex() + 1;
            cell.setCellFormula("Data!$A$2");
        }

        if (cell.getRowIndex() == 4 && 1 == cell.getColumnIndex()) {
            System.out.println("Row is " +  cell.getRowIndex() + "And Cell is " + cell.getColumnIndex());
            int actualCellRowNum = cell.getRowIndex() + 1;
            cell.setCellFormula("Data!$B$2");
        }

        if (cell.getRowIndex() == 3 && 2 == cell.getColumnIndex()) {

            int actualCellRowNum = cell.getRowIndex() + 1;
            cell.setCellFormula("B" + actualCellRowNum +"*$F$3");
        }
        if ((cell.getRowIndex() == 3 || cell.getRowIndex()==5) && 0 == cell.getColumnIndex() ) {
            Font font = workbook.createFont();
            font.setBold(true);
            cellStyle.setFont(font);
            cell.setCellStyle(cellStyle);
        }
        if (cell.getRowIndex() == 4 && 2 == cell.getColumnIndex()) {
            int actualCellRowNum = cell.getRowIndex() + 1;
            cell.setCellFormula("B" + actualCellRowNum +"*$F$3");
        }
        if (cell.getRowIndex() == 5 && 2 == cell.getColumnIndex()) {
            int actualCellRowNum = cell.getRowIndex() + 1;
            cell.setCellFormula("B" + actualCellRowNum +"*$F$3");
        }
        if (cell.getRowIndex() == 7 && 2 == cell.getColumnIndex()) {
            int actualCellRowNum = cell.getRowIndex() + 1;
            cell.setCellFormula("B" + actualCellRowNum +"*$F$3");
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(IndexedColors.YELLOW1.getIndex());
            cell.setCellStyle(cellStyle);
        }
        if (cell.getRowIndex() == 6 && 2 == cell.getColumnIndex()) {
            int actualCellRowNum = cell.getRowIndex() + 1;
            cell.setCellFormula("B" + actualCellRowNum +"*$F$3");
        }
        if(cell.getRowIndex()==5 && 1== cell.getColumnIndex()){
            int BeforeCellRowNumber = cell.getRowIndex();
            int BeforeBeforeCellRowNumber = cell.getRowIndex()-1;
            cell.setCellFormula("B"+BeforeCellRowNumber+"+B"+BeforeBeforeCellRowNumber);
        }
        if(cell.getRowIndex()==7 && 1== cell.getColumnIndex()) {
            int BeforeCellRowNumber = cell.getRowIndex();
            int BeforeBeforeCellRowNumber = cell.getRowIndex()-1;
            cell.setCellFormula("B"+BeforeBeforeCellRowNumber+"-B"+BeforeCellRowNumber);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(IndexedColors.YELLOW1.getIndex());
            cell.setCellStyle(cellStyle);
        }
    }
}