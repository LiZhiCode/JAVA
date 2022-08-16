package com.lz.parse;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @author qiliz
 */
public class ExcelPoi {
    public static void main(String[] args) throws IOException {
        String filePath = "E:\\_Info\\Work\\HP\\upload\\7QR88-40126_13Jun2022_CPK.xlsm";
        String cellValue = "";

        FileInputStream fileInputStream = new FileInputStream(filePath);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheet("P1");
        List<Object> estCpkList = new ArrayList<>();
        Iterator<Row> rows = sheet.rowIterator();
        System.out.println(sheet.getRow(7).getPhysicalNumberOfCells());

        int rowCount = sheet.getRow(7).getPhysicalNumberOfCells();
        Row row;
        Cell cell;
        while(rows.hasNext()){
            row = rows.next();
            //获取单元格
            Iterator<Cell> cells =row.cellIterator();
            for (int rowNum = 0; rowNum < rowCount; rowNum++) {
                if (cells.hasNext()) {
                    cell = cells.next();
                    CellType cellType = cell.getCellType();

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
                        default: cellValue = " ";
                    }
                    if (cellValue.contains("Est. Cpk")) {
                        estCpkList.add(cellValue);
                    }
//                    System.out.print(cellValue + ", ");
                }
            }
            System.out.println(estCpkList);
        }





//        Row rowData = sheet.getRow(7);
//        Row rowDataTarget = sheet.getRow(7);
//        if (rowData != null) {
//            int cellCount = rowDataTarget.getPhysicalNumberOfCells();
//            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
//                Cell cell = rowData.getCell(cellNum);
//                CellType cellType = cell.getCellType();
//                System.out.println(cell);
//
//
////                CellStyle cellStyle = cell.getCellStyle();
////                //背景颜色
////                XSSFColor xssfColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
////                byte[] bytes;
////                if (xssfColor != null) {
////                    bytes = xssfColor.getRGB();
////                    System.out.println("bg: " + String.format("#%02X%02X%02X", bytes[0], bytes[1], bytes[2]));
////                }
////                //获取字体
////                XSSFFont eFont = workbook.getFontAt(cell.getCellStyle().getFontIndex());
////                XSSFColor xssfFontColor = eFont.getXSSFColor();
////                XSSFColor color = eFont.getXSSFColor();
////                byte[] rgb;
////                if (xssfFontColor != null) {
////                    rgb = xssfFontColor.getRGB();
////                    System.out.println("rgb: " + String.format("#%02X%02X%02X", rgb[0], rgb[1], rgb[2]));
////                    System.out.println("rgb String: " + Arrays.toString(rgb));
////                }
//
//
//                switch (cellType) {
//                    //如果该excel是公式
//                    case FORMULA:
//                        try {
//                            System.out.println("FORMULA");
//                            cellValue = String.valueOf(cell.getNumericCellValue());
//                        } catch (IllegalStateException e) {
//                            if (e.getMessage().contains("from a STRING cell")) {
//                                cellValue = String.valueOf(cell.getStringCellValue());
//                            } else if (e.getMessage().contains("from a ERROR formula cell")) {
//                                cellValue = String.valueOf(cell.getErrorCellValue());
//                            }
//                        }
//                        break;
//                    //如果该excel是字符串
//                    case STRING:
//                        cellValue = String.valueOf(cell);
//                        break;
//                    //如果是double型
//                    case NUMERIC:
//                        cellValue = String.valueOf(cell.getNumericCellValue());
//                        break;
//                    //如果是空格
//                    case BLANK:
//                        break;
//                    //如果是布尔类型
//                    case BOOLEAN:
//                        cellValue = String.valueOf(cell.getBooleanCellValue());
//                        break;
//                    case _NONE:
//                        cellValue = "EXCEPTION:NONE";
//                        break;
//                    case ERROR:
//                        cellValue = "EXCEPTION:ERROR";
//                        break;
//                    default: cellValue = " ";
//                }
//                System.out.println(cellType + " | " + cellValue);
//            }
//        }

        fileInputStream.close();
    }
}

