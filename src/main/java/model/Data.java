package model;

import Handler.DateTimeConverter;
import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.annotation.format.NumberFormat;
import com.alibaba.excel.annotation.write.style.ContentStyle;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.oracle.webservices.internal.api.databinding.DatabindingMode;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.Date;

@HeadStyle(horizontalAlignment = HorizontalAlignment.LEFT, fillForegroundColor= 40, wrapped = false, fillPatternType = FillPatternType.SOLID_FOREGROUND,borderBottom = BorderStyle.NONE,borderLeft = BorderStyle.NONE, borderTop = BorderStyle.NONE, borderRight = BorderStyle.NONE)
@HeadFontStyle(color = 12 ,fontHeightInPoints = 11,fontName = "Calibri")
public class Data {
    @ExcelIgnore
    private final String period;
    @ExcelProperty("Period International")
    private final String periodInternational;
    @ExcelProperty("Period Domestic")
    private final String periodDomestic;
    @ExcelProperty("Company Code")
    private final String companyCode;
    @ExcelProperty("Company Descr")
    private final String companyDescription;
    @ExcelProperty("Assignment")
    private final String assignment;
    @ExcelProperty("GL Account")
    private final String glAccount;
    @ExcelProperty("Reference Document Number")
    private final String refDocNumber;
    @ExcelProperty("Accounting Document")
    private final String accountingDocument;
    @ExcelProperty("Business Area")
    private final String businessArea;
    @ExcelProperty("Document Type")
    private final String documentType;
    @ExcelIgnore
    private final String documentDate;
    @ExcelIgnore
    private final String postingDate;
    @ExcelProperty("Document Date")
    @ContentStyle(horizontalAlignment = HorizontalAlignment.RIGHT,dataFormat = 14)
    private final String documentDateExcel;
    @ExcelProperty("Posting Date")
    @ContentStyle(horizontalAlignment = HorizontalAlignment.RIGHT,dataFormat = 14)
    private final String postingDateExcel;
    @ExcelProperty("Posting Key")
    private final String postingKey;
    @ExcelProperty("Tax Code")
    private final String taxCode;
    @ExcelProperty("Local Amount 1")
    @ContentStyle(dataFormat = 2,horizontalAlignment = HorizontalAlignment.RIGHT)
    private final Double localAmount1;
    @ExcelProperty("Local CCY 1")
    private final String localCurrency1;
    @ExcelProperty("Local Amount 2")
    @ContentStyle(dataFormat = 2,horizontalAlignment = HorizontalAlignment.RIGHT)
    private final Double localAmount2;
    @ExcelProperty("Local CCY 2")
    private final String localCurrency2;
    @ExcelProperty("Local Amount 3")
    @ContentStyle(dataFormat = 2,horizontalAlignment = HorizontalAlignment.RIGHT)
    private final Double localAmount3;
    @ExcelProperty("Local CCY 3")
    private final String localCurrency3;
    @ExcelProperty("Foreign CCY Amount")
    @ContentStyle(dataFormat = 2,horizontalAlignment = HorizontalAlignment.RIGHT)
    private final Double foreignCurrencyAmt;
    @ExcelProperty("Document CCY")
    private final String documentCurrency;
    @ExcelProperty("Document Amount")
    @ContentStyle(dataFormat = 2,horizontalAlignment = HorizontalAlignment.RIGHT)
    private final Double documentAmount;
    @ExcelProperty("Clearing Document")
    private final String clearingDocument;
    @ExcelProperty("Document Header Text")
    private final String documentHeaderText;
    @ExcelProperty("Global CCY (USD)")
    @ContentStyle(dataFormat = 2,horizontalAlignment = HorizontalAlignment.RIGHT)
    private final Double globalCurrencyUSD;
    @ExcelProperty("Local Account")
    private final String localAccount;
    @ContentStyle(horizontalAlignment = HorizontalAlignment.RIGHT,dataFormat = 15)
    @ExcelProperty(value = "Import Date (UTC)")
    private final String importDateUtc;

    public Data(String period, String periodInternational, String periodDomestic, String companyCode, String companyDescription,
            String assignment, String glAccount, String refDocNumber, String accountingDocument, String businessArea, String documentType,
            String documentDate, String postingDate, String documentDateExcel, String postingDateExcel, String postingKey, String taxCode,
            Double localAmount1, String localCurrency1, Double localAmount2, String localCurrency2, Double localAmount3, String localCurrency3,
            Double foreignCurrencyAmount, String documentCurrency, Double documentAmount, String clearingDocument, String documentHeaderText,
            Double globalCurrencyUSD, String localAccount, String importDateUtc) {
        this.period = period;
        this.periodInternational = periodInternational;
        this.periodDomestic = periodDomestic;
        this.companyCode = companyCode;
        this.companyDescription = companyDescription;
        this.assignment = assignment;
        this.glAccount = glAccount;
        this.refDocNumber = refDocNumber;
        this.accountingDocument = accountingDocument;
        this.businessArea = businessArea;
        this.documentType = documentType;
        this.documentDate = documentDate;
        this.postingDate = postingDate;
        this.documentDateExcel = documentDateExcel;
        this.postingDateExcel = postingDateExcel;
        this.postingKey = postingKey;
        this.taxCode = taxCode;
        this.localAmount1 = localAmount1;
        this.localCurrency1 = localCurrency1;
        this.localAmount2 = localAmount2;
        this.localCurrency2 = localCurrency2;
        this.localAmount3 = localAmount3;
        this.localCurrency3 = localCurrency3;
        this.foreignCurrencyAmt = foreignCurrencyAmount;
        this.documentCurrency = documentCurrency;
        this.documentAmount = documentAmount;
        this.clearingDocument = clearingDocument;
        this.documentHeaderText = documentHeaderText;
        this.globalCurrencyUSD = globalCurrencyUSD;
        this.localAccount = localAccount;
        this.importDateUtc = importDateUtc;
    }

    public String getPeriod() {
        return period;
    }

    public String getPeriodInternational() {
        return periodInternational;
    }

    public String getPeriodDomestic() {
        return periodDomestic;
    }

    public String getCompanyCode() {
        return companyCode;
    }

    public String getCompanyDescription() {
        return companyDescription;
    }

    public String getAssignment() {
        return assignment;
    }

    public String getGlAccount() {
        return glAccount;
    }

    public String getRefDocNumber() {
        return refDocNumber;
    }

    public String getAccountingDocument() {
        return accountingDocument;
    }

    public String getBusinessArea() {
        return businessArea;
    }

    public String getDocumentType() {
        return documentType;
    }

    public String getDocumentDate() {
        return documentDate;
    }

    public String getPostingDate() {
        return postingDate;
    }

    public String getDocumentDateExcel() {
        return documentDateExcel;
    }

    public String getPostingDateExcel() {
        return postingDateExcel;
    }

    public String getPostingKey() {
        return postingKey;
    }

    public String getTaxCode() {
        return taxCode;
    }

    public Double getLocalAmount1() {
        return localAmount1;
    }

    public String getLocalCurrency1() {
        return localCurrency1;
    }

    public Double getLocalAmount2() {
        return localAmount2;
    }

    public String getLocalCurrency2() {
        return localCurrency2;
    }

    public Double getLocalAmount3() {
        return localAmount3;
    }

    public String getLocalCurrency3() {
        return localCurrency3;
    }

    public Double getForeignCurrencyAmount() {
        return foreignCurrencyAmt;
    }

    public String getDocumentCurrency() {
        return documentCurrency;
    }

    public Double getDocumentAmount() {
        return documentAmount;
    }

    public String getClearingDocument() {
        return clearingDocument;
    }

    public String getDocumentHeaderText() {
        return documentHeaderText;
    }

    public Double getGlobalCurrencyUSD() {
        return globalCurrencyUSD;
    }

    public String getLocalAccount() {
        return localAccount;
    }

    public String getImportDateUtc() {
        return importDateUtc;
    }


}
