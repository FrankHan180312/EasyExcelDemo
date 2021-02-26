package Handler;

import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.zip.DeflaterOutputStream;

public class CustomCellWriteHandler implements CellWriteHandler {

    private static CellStyle cellStyle = null;

    private static Workbook workbook = null;
    private static CreationHelper creationHelper  = null;
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

        if(workbook ==null){
            workbook = writeSheetHolder.getSheet().getWorkbook();
        }

        if(cellStyle==null) {
            cellStyle = workbook.createCellStyle();
        }
        if(creationHelper ==null){
            creationHelper = workbook.getCreationHelper();
        }
        try {
            if(cell.getRowIndex() >0 && cell.getColumnIndex()==28){
                cellStyle.setDataFormat(
                        creationHelper.createDataFormat().getFormat("dd-MM-yyyy")
                                       );

                final short format = creationHelper.createDataFormat().getFormat("dd-MM-yyyy");
                cell.setCellStyle(cellStyle);
                cell.setCellValue(StringToDate(list.get(0).getStringValue()));
            }
            if(cell.getRowIndex() >0 && cell.getColumnIndex()==26){
                cellStyle.setDataFormat(
                        creationHelper.createDataFormat().getFormat("0.00")
                                       );
                cell.setCellStyle(cellStyle);
                cell.setCellValue(Double.parseDouble(list.get(0).getNumberValue().toString()));
            }
        }
        catch (Exception e){
            e.printStackTrace();
        }

    }

    private static String GetCellValue(CellData cellData){
        String cellValue = "";
        if(cellData.getType()== CellDataTypeEnum.STRING){
            cellValue = cellData.getStringValue();
        }
        if(cellData.getType()==CellDataTypeEnum.NUMBER){

        }
        return cellValue;
    }

    private static Date StringToDate(String inputDate) {
        SimpleDateFormat formatter = new SimpleDateFormat("dd-MMM-yyyy");
        Date date = null;
        try {
            date = formatter.parse(inputDate);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return date;
    }
}
