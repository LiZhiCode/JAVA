package com.lz.parse;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.text.DecimalFormat;
import java.util.*;

/**
 * @author qiliz
 */
public class ExcelFaiParse {
    static List<Map<Integer, String>> specMapList = new LinkedList<>();
    static List<Object> resultMapList = new ArrayList<>();
    static String cavitiesValue = "";
    static String fcdValue = "";
    static String cavity = "Cavity # :";


    public static void main(String[] args) {
        String start = "Dim#";
        String end = "End of Report";
        final boolean[] flag = {false};
        EasyExcel.read("E:\\_Info\\Work\\HP\\upload\\7QR88-40126_11Jun2022_FAI.xlsm")
                .registerReadListener(new AnalysisEventListener<Map<Integer, String>>() {
                    @Override
                    public void invoke(Map<Integer, String> integerStringMap, AnalysisContext analysisContext) {
                        if (Objects.equals(integerStringMap.get(8), cavity)) {
                            cavitiesValue = integerStringMap.get(10);
                        }
                        if (Objects.equals(integerStringMap.get(0), start)) {
                            flag[0] = true;
                        }
                        if (Objects.equals(integerStringMap.get(0), end)) {
                            flag[0] = false;
                        }
                        if (flag[0] && (integerStringMap.get(9) != null)) {
                            specMapList.add(integerStringMap);
                        }
                    }

                    @Override
                    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
                    }
                })
                .doReadAll();
        int i = 0;
        for (Map<Integer, String> item : specMapList) {
            if (!Objects.equals(item.get(0), "Dim#")) {
                DecimalFormat decimalFormat = new DecimalFormat("#.##");
                double cData = Double.parseDouble(item.get(2));
                double dData = Double.parseDouble(item.get(3));
                double eData = Double.parseDouble(item.get(4));
                double jData = Double.parseDouble(item.get(9));
                double hData = Double.parseDouble(decimalFormat.format(Double.parseDouble(item.get(7))));
                String fcdData = item.get(6);
                if (jData > 0 && Objects.equals(fcdData, "Y")) {
                    i++;
                }

                double ceData = Double.parseDouble(decimalFormat.format(cData - Math.abs(eData)));
                double cdData = Double.parseDouble(decimalFormat.format(cData + Math.abs(dData)));

                String cpStr = "Cavity:" + cavitiesValue + "; "
                        + "Dimension:" + item.get(0) + "; "
                        + "Dimm out of spec; "
                        + "Value:" + hData + "; "
                        + "LSL:" + ceData + "; "
                        + "USL:" + cdData + "; ";
                System.out.println(cpStr);
                resultMapList.add(cpStr);
            }
        }
        if (i > 0) {
            fcdValue = i +" FCD dimension out of spec.";
            resultMapList.add(fcdValue);
        }
        System.out.println(resultMapList);
    }
}
