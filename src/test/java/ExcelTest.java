import Handler.AutoWidthWriteHandler;
import Handler.CustomCellWriteHandler;
import Handler.ExcekUtils;
import Handler.FillHandler;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import model.Data;
import model.FillDate;
import model.SummaryBalance;
import model.TemplateData;
import org.apache.poi.ss.usermodel.*;
import org.junit.Test;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;

import java.util.*;

public class ExcelTest {


    @Test
    public void test(){
        SimpleDateFormat df1 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        SimpleDateFormat df = new SimpleDateFormat("dd-mm-yyyy");
        String fileName = "text"+df.format(new Date())+".xlsx";
        System.out.println("begin at " + df1.format(new Date()));

        EasyExcel.write(fileName, Data.class)
                 .registerWriteHandler(new AutoWidthWriteHandler())
                 //.registerWriteHandler(new CustomCellWriteHandler())
                 .sheet("page")
                 .doWrite(data());
        System.out.println("end at " + df1.format(new Date()));
    }

    @Test
    public void TestExcelTemplate(){
        String templateName = "C:/Frank/Big Data/prac/aa.xlsx";
        String fileName = System.currentTimeMillis()+".xlsx";
        TemplateData data = new TemplateData();
        data.setDate(new Date());
        data.setString("Frank");
        data.setDoubleData(11.22);
        EasyExcel.write(fileName).withTemplate(templateName).sheet().doFill(data);

    }

    @Test
    public void TestExcelTemplateFill(){
        // 1. 路径定义
        // 输入模板路径
        String templatePath = "C:/Frank/Big Data/prac/VRS_Template.xlsx";
        // 结果输出路径
        SimpleDateFormat df = new SimpleDateFormat("dd-mm-yyyy");
        String fileName = "Recon"+df.format(new Date())+".xlsx";

        // 2. 真实的模拟数据
        List<SummaryBalance> list = SummaryBalanceList();

        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();

        // 3. 数据写出
        ExcelWriter writer = EasyExcel.write(byteArrayOutputStream).withTemplate(new File(templatePath)).build();
        // 3.1 操作第一个sheet(记得注册自定义的CellWriteHandler)
        WriteSheet sheet = EasyExcel.writerSheet("1_Summary Reconciliation")
                .registerWriteHandler(new FillHandler())
                .build();
        WriteSheet sheet2 = EasyExcel.writerSheet("1_Summary Reconciliation Copy")
                .registerWriteHandler(new FillHandler())
                .build();
        // 3.2 填充列表数据
        writer.fill(list, FillConfig.builder().forceNewRow(Boolean.TRUE).build(), sheet);
        writer.fill(list, FillConfig.builder().forceNewRow(Boolean.TRUE).build(), sheet2);
        // 3.3 填充其它动态信息
        Map<String, Object> extra = new LinkedHashMap<>();
        // 单位：10%
        extra.put("exchcangeRate", 1);
        //extra.put("glBalance", 1234);
        //extra.put("supportBalance", 5678);
        //extra.put("glUSD",0);
        //extra.put("sbUSD",0);
        //3.4 设置强制计算公式：不然公式会以字符串的形式显示在excel中
        Workbook workbook = writer.writeContext().writeWorkbookHolder().getWorkbook();
        workbook.setForceFormulaRecalculation(true);


        // 3.5 数据刷新
        writer.fill(extra, sheet);
        writer.fill(extra, sheet2);

        writer.finish();

        //System.out.println(byteArrayOutputStream.toByteArray());

    }

    private static List<SummaryBalance> SummaryBalanceList()
    {
        ArrayList<SummaryBalance> list = new ArrayList<SummaryBalance>();
        SummaryBalance summaryBalance = new SummaryBalance();
        summaryBalance.setBalanceDescr("G/L End Balance");
        summaryBalance.setBalanceLocal("");
        //summaryBalance.setBalanceUSD("Fomular");
        list.add(summaryBalance);
        SummaryBalance summaryBalance1 = new SummaryBalance();
        summaryBalance1.setBalanceDescr("Balance Supporting Documentation -");
        summaryBalance1.setBalanceLocal("");
        //summaryBalance1.setBalanceUSD(" ");
        list.add(summaryBalance1);
        SummaryBalance summaryBalance2 = new SummaryBalance();
        summaryBalance2.setBalanceDescr("Difference =");
        summaryBalance2.setBalanceLocal(" ");
        //summaryBalance2.setBalanceUSD("");
        list.add(summaryBalance2);
        SummaryBalance summaryBalance3 = new SummaryBalance();
        summaryBalance3.setBalanceDescr("Reconciling items +");
        summaryBalance3.setBalanceLocal("9999");
        //summaryBalance3.setBalanceUSD("");
        list.add(summaryBalance3);
        SummaryBalance summaryBalance4 = new SummaryBalance();
        summaryBalance4.setBalanceDescr("Net Unreconciled difference");
        summaryBalance4.setBalanceLocal("");
        summaryBalance4.setBalanceUSD("");
        list.add(summaryBalance4);
        return list;
    }

    private static List<FillDate> dataList()
    {
        ArrayList<FillDate> list = new ArrayList<FillDate>();
        for (int i = 1; i <= 20; i++)
        {
            FillDate demoData = new FillDate();
            demoData.setOrderId("202001180000" + i);
            demoData.setGoodName("商品名称" + i);
            demoData.setGoodPrice(new BigDecimal(i));
            demoData.setNum(i);

            list.add(demoData);
        }
        return list;
    }

    private List<Data> data() {
        SimpleDateFormat df1 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        List<Data> list = new ArrayList<Data>();
        for (int i = 0; i < 12000; i++) {
            Data dataTemp = new Data("2020002", "2020-10 (I)", "2020-09 (D)", "3861", "", "20200218", "0010091010", "9300021555", "9300021555", "", "XX", "18-02-20", "18-02-20",
                    "18-Feb-2020", "18-Feb-2020", "50", "A1", Double.valueOf("-4500"), "VND", Double.valueOf("-1"), "USD", Double.valueOf("0"), "", Double.valueOf("-4500"),
                    "VND", Double.valueOf("4500"), "9300021572", "", Double.valueOf("107.10"), "", "30-Dec-2020");
            list.add(dataTemp);
        }

        return list;
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
