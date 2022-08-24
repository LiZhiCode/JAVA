package com.lz.parse;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

/**
 * @author qiliz
 */

@Data
public class ExcelData {
    @ExcelProperty("Dim#")
    private String dim;
    @ExcelProperty("Nominal")
    private String nominal;
    @ExcelProperty("Upper Tol")
    private String upperTol;
    @ExcelProperty("Lower Tol")
    private String lowerTol;
    @ExcelProperty("FCD")
    private String fcd;
    @ExcelProperty("Measured Value")
    private String measuredValue;
    @ExcelProperty("Out of Tol.")
    private String outOfTol;
}
