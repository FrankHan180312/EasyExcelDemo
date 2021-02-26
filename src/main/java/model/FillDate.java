package model;

import java.io.Serializable;
import java.math.BigDecimal;

public class FillDate implements Serializable
{

    private String orderId;
    private String goodName;
    private BigDecimal goodPrice;
    private Integer num;

    /**
     * 税价：公式单元格（给一个空串的默认值）
     */
    private String taxMoney = "";

    /**
     * 总价：公式单元格（给一个空串的默认值）
     */
    private String totalMoney = "";

    public String getOrderId() {
        return orderId;
    }

    public void setOrderId(String orderId) {
        this.orderId = orderId;
    }

    public String getGoodName() {
        return goodName;
    }

    public void setGoodName(String goodName) {
        this.goodName = goodName;
    }

    public BigDecimal getGoodPrice() {
        return goodPrice;
    }

    public void setGoodPrice(BigDecimal goodPrice) {
        this.goodPrice = goodPrice;
    }

    public Integer getNum() {
        return num;
    }

    public void setNum(Integer num) {
        this.num = num;
    }

    public String getTaxMoney() {
        return taxMoney;
    }

    public void setTaxMoney(String taxMoney) {
        this.taxMoney = taxMoney;
    }

    public String getTotalMoney() {
        return totalMoney;
    }

    public void setTotalMoney(String totalMoney) {
        this.totalMoney = totalMoney;
    }
}
