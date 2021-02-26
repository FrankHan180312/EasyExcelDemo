package Handler;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class DateTimeConverter implements Converter<String> {

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


    public Class supportJavaTypeKey() {
        return String.class;
    }

    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    public String convertToJavaData(CellData cellData, ExcelContentProperty excelContentProperty, GlobalConfiguration globalConfiguration)
            throws Exception {
        return null;
    }

    public CellData<Date> convertToExcelData(String s, ExcelContentProperty excelContentProperty, GlobalConfiguration globalConfiguration)
            throws Exception {
        Date date = StringToDate(s);
        return new CellData(date);
    }
}
