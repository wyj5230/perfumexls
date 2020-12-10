public class Title
{
    private String date;
    private String store;
    private String storeCode;
    private String barCode;
    private String saleType;
    private String qty;
    private String netAmount1;
    private String netAmount2;
    private String grossAmount;
    private String unitListPrice;
    private String style;
    private String fabric;
    private String color;
    private String size;

    public Title(String date, String store, String storeCode, String barCode, String saleType, String qty,
                 String netAmount1, String netAmount2, String grossAmount, String unitListPrice, String style,
                 String fabric, String color, String size)
    {
        this.date = date;
        this.store = store;
        this.storeCode = storeCode;
        this.barCode = barCode;
        this.saleType = saleType;
        this.qty = qty;
        this.netAmount1 = netAmount1;
        this.netAmount2 = netAmount2;
        this.grossAmount = grossAmount;
        this.unitListPrice = unitListPrice;
        this.style = style;
        this.fabric = fabric;
        this.color = color;
        this.size = size;
    }

    public String getDate()
    {
        return date;
    }

    public void setDate(String date)
    {
        this.date = date;
    }

    public String getStore()
    {
        return store;
    }

    public void setStore(String store)
    {
        this.store = store;
    }

    public String getStoreCode()
    {
        return storeCode;
    }

    public void setStoreCode(String storeCode)
    {
        this.storeCode = storeCode;
    }

    public String getBarCode()
    {
        return barCode;
    }

    public void setBarCode(String barCode)
    {
        this.barCode = barCode;
    }

    public String getSaleType()
    {
        return saleType;
    }

    public void setSaleType(String saleType)
    {
        this.saleType = saleType;
    }

    public String getQty()
    {
        return qty;
    }

    public void setQty(String qty)
    {
        this.qty = qty;
    }

    public String getNetAmount1()
    {
        return netAmount1;
    }

    public void setNetAmount1(String netAmount1)
    {
        this.netAmount1 = netAmount1;
    }

    public String getNetAmount2()
    {
        return netAmount2;
    }

    public void setNetAmount2(String netAmount2)
    {
        this.netAmount2 = netAmount2;
    }

    public String getGrossAmount()
    {
        return grossAmount;
    }

    public void setGrossAmount(String grossAmount)
    {
        this.grossAmount = grossAmount;
    }

    public String getUnitListPrice()
    {
        return unitListPrice;
    }

    public void setUnitListPrice(String unitListPrice)
    {
        this.unitListPrice = unitListPrice;
    }

    public String getStyle()
    {
        return style;
    }

    public void setStyle(String style)
    {
        this.style = style;
    }

    public String getFabric()
    {
        return fabric;
    }

    public void setFabric(String fabric)
    {
        this.fabric = fabric;
    }

    public String getColor()
    {
        return color;
    }

    public void setColor(String color)
    {
        this.color = color;
    }

    public String getSize()
    {
        return size;
    }

    public void setSize(String size)
    {
        this.size = size;
    }
}
