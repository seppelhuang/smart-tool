
package cn.seppel.smarttool.pdfexcel;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

@Data
public class LineItem {

    @ExcelProperty("Line No")
    private String lineNo;

    @ExcelProperty("Quantity")
    private String quantity;

    @ExcelProperty("Unit")
    private String unit;

    @ExcelProperty("Delta Item")
    private String deltaItem;

    @ExcelProperty("Description")
    private String description;

    @ExcelProperty("Rev Level")
    private String revLevel;

    @ExcelProperty("Delivery Date")
    private String deliveryDate;

    @ExcelProperty("Unit Price")
    private String unitPrice;

    @ExcelProperty("Extended Price")
    private String extendedPrice;
}
