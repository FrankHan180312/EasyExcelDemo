package Handler;

import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class ExcekUtils {
    public static HorizontalCellStyleStrategy setHorizontalCellStyleStrategy(){
        /*******自定义列标题和内容的样式******/
        // 头的策略
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        // 背景设置为红色
        //headWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        headWriteCellStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        WriteFont headWriteFont = new WriteFont();
        headWriteFont.setBold(true);
        headWriteFont.setFontName("Calibri");
        headWriteFont.setFontHeightInPoints((short)12);
        headWriteCellStyle.setWriteFont(headWriteFont);
        headWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.LEFT);
        headWriteCellStyle.setBorderRight(BorderStyle.NONE);
        headWriteCellStyle.setBorderLeft(BorderStyle.NONE);
        headWriteCellStyle.setBorderTop(BorderStyle.NONE);
        headWriteCellStyle.setBorderBottom(BorderStyle.NONE);
        // 内容的策略
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND 不然无法显示背景颜色.头默认了 FillPatternType所以可以不指定
        //  contentWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        // 背景绿色
        //  contentWriteCellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        WriteFont contentWriteFont = new WriteFont();
        // 字体大小
        contentWriteFont.setFontHeightInPoints((short)10);
        contentWriteCellStyle.setWriteFont(contentWriteFont);
        //contentWriteCellStyle.setDataFormat();

        //设置 自动换行
        contentWriteCellStyle.setWrapped(true);
        //设置 垂直居中
        contentWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置 水平居中
        contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.LEFT);
        //        //设置边框样式
        //       contentWriteCellStyle.setBorderLeft(BorderStyle.MEDIUM);
        //      contentWriteCellStyle.setBottomBorderColor(IndexedColors.BLUE.getIndex()); 颜色
        //      contentWriteCellStyle.setBorderTop(DASHED);
        //        contentWriteCellStyle.setBorderRight(DASHED);
        //       contentWriteCellStyle.setBorderBottom(DASHED);

        // 这个策略是 头是头的样式 内容是内容的样式 其他的策略可以自己实现
        HorizontalCellStyleStrategy horizontalCellStyleStrategy =
                new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
        return horizontalCellStyleStrategy;
    }
}
