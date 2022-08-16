package com.lz.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

/**
 * @author qiliz
 */

@Data
public class ExcelData {
    @ExcelProperty("Dimension #")
    private String dimension;
    @ExcelProperty("Est. Cpk")
    private String estCpk;
}
