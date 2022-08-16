package com.Excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Test;


import java.io.FileInputStream;
import java.util.Arrays;

/**
 * @author qiliz
 */
public class ExcelReadTest {
    String filePath = "E:\\_Info\\Work\\HP\\upload\\7QR88-40126_13Jun2022_CPK.xlsm";
    String fileTestPath = "E:\\_Info\\Work\\HP\\upload\\test.xlsx";
    String cellValue = "";

    @Test
    public void testReadXlsm() throws Exception {
        FileInputStream fileInputStream = new FileInputStream(filePath);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        workbook.setForceFormulaRecalculation(true);
        Sheet sheet = workbook.getSheet("P1");
        sheet.setForceFormulaRecalculation(true);
        Row rowData = sheet.getRow(82);
        Row rowDataTarget = sheet.getRow(13);
        if (rowData != null) {
            int cellCount = rowDataTarget.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowData.getCell(cellNum);
                CellType cellType = cell.getCellType();
                CellStyle cellStyle = cell.getCellStyle();

                //后面使用它来执行计算公式
                FormulaEvaluator formulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
                //获取单元格内容的类型
                CellType cellType2 = cell.getCellType();
                //判断是否存储的为公式，此处本可以不加判断
                if (cellType2.equals(CellType.FORMULA)){
                    //获取公式，可以理解为已String类型获取cell的值输出
                    String cellFormula = cell.getCellFormula();
                    System.out.println(cellFormula);
                    //执行公式，此处cell的值就是公式
//                    CellValue evaluate = formulaEvaluator.evaluate(cell);
//                    System.out.println(evaluate.formatAsString());
                }


                //背景颜色
                XSSFColor xssfColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
                byte[] bytes;
                if (xssfColor != null) {
                    bytes = xssfColor.getRGB();
                    System.out.println("bg: " + String.format("#%02X%02X%02X", bytes[0], bytes[1], bytes[2]));
                }

                //获取字体
                XSSFFont eFont = workbook.getFontAt(cell.getCellStyle().getFontIndex());
                XSSFColor xssfFontColor = eFont.getXSSFColor();

                byte[] rgb;
                if (xssfFontColor != null) {
                    rgb = xssfFontColor.getRGB(); //得到rgb的byte数组
                    System.out.println("rgb: " + String.format("#%02X%02X%02X", rgb[0], rgb[1], rgb[2]));
                    System.out.println("rgb String: " + Arrays.toString(rgb));
                }

                switch (cellType) {
                    //如果该excel是公式
                    case FORMULA:
                        try {
                            cellValue = String.valueOf(cell.getNumericCellValue());
                        } catch (IllegalStateException e) {
                            if (e.getMessage().contains("from a STRING cell")) {
                                cellValue = String.valueOf(cell.getStringCellValue());
                            } else if (e.getMessage().contains("from a ERROR formula cell")) {
                                cellValue = String.valueOf(cell.getErrorCellValue());
                            }
                        }
                        break;
                    //如果该excel是字符串
                    case STRING:
                        cellValue = String.valueOf(cell);
                        break;
                    //如果是double型
                    case NUMERIC:
                        cellValue = String.valueOf(cell.getNumericCellValue());
                        break;
                    //如果是空格
                    case BLANK:
                        break;
                    //如果是布尔类型
                    case BOOLEAN:
                        cellValue = String.valueOf(cell.getBooleanCellValue());
                        break;
                    case _NONE:
                        cellValue = "EXCEPTION:NONE";
                        break;
                    case ERROR:
                        cellValue = "EXCEPTION:ERROR";
                        break;
                }
                System.out.println(cellType + " | " + cellValue);
            }
        }

        fileInputStream.close();
    }

    @Test
    public void testReadXlsx() throws Exception {
        FileInputStream fileInputStream = new FileInputStream(fileTestPath);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheet("P1");
        Row rowData = sheet.getRow(7);
        Row rowDataTarget = sheet.getRow(7);
        if (rowData != null) {
            int cellCount = rowDataTarget.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowData.getCell(cellNum);
                CellType cellType = cell.getCellType();
                CellStyle cellStyle = cell.getCellStyle();


                //背景颜色
                XSSFColor xssfColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
                byte[] bytes;
                if (xssfColor != null) {
                    bytes = xssfColor.getRGB();
                    System.out.println("bg: " + String.format("#%02X%02X%02X", bytes[0], bytes[1], bytes[2]));
                }


                //获取字体
                XSSFFont eFont = workbook.getFontAt(cell.getCellStyle().getFontIndex());
                XSSFColor xssfFontColor = eFont.getXSSFColor();
                XSSFColor color = eFont.getXSSFColor();
                System.out.println("font: " + Arrays.toString(color.getARGB()));

                byte[] rgb;
                if (xssfFontColor != null) {
                    rgb = xssfFontColor.getRGB(); //得到rgb的byte数组
                    System.out.println("rgb: " + String.format("#%02X%02X%02X", rgb[0], rgb[1], rgb[2]));
                    System.out.println("rgb String: " + Arrays.toString(rgb));
                }



                switch (cellType) {
                    //如果该excel是公式
                    case FORMULA:
                        try {
                            System.out.println("FORMULA");
                            cellValue = String.valueOf(cell.getNumericCellValue());
                        } catch (IllegalStateException e) {
                            if (e.getMessage().contains("from a STRING cell")) {
                                cellValue = String.valueOf(cell.getStringCellValue());
                            } else if (e.getMessage().contains("from a ERROR formula cell")) {
                                cellValue = String.valueOf(cell.getErrorCellValue());
                            }
                        }
                        break;
                    //如果该excel是字符串
                    case STRING:
                        cellValue = String.valueOf(cell);
                        break;
                    //如果是double型
                    case NUMERIC:
                        cellValue = String.valueOf(cell.getNumericCellValue());
                        break;
                    //如果是空格
                    case BLANK:
                        break;
                    //如果是布尔类型
                    case BOOLEAN:
                        cellValue = String.valueOf(cell.getBooleanCellValue());
                        break;
                    case _NONE:
                        cellValue = "EXCEPTION:NONE";
                        break;
                    case ERROR:
                        cellValue = "EXCEPTION:ERROR";
                        break;
                }
                System.out.println(cellType + " | " + cellValue);
            }
        }

        fileInputStream.close();
    }
}
