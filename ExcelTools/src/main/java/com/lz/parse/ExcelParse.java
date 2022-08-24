package com.lz.parse;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.util.*;

/**
 * @author qiliz
 */
public class ExcelParse {
    static final List<String> CAVITIES_LIST = new LinkedList<>();
    static final List<String> DIMENSION_LIST = new LinkedList<>();
    static final List<Integer> CP_INX_LIST = new LinkedList<>();
    static final List<Integer> CPK_INX_LIST = new LinkedList<>();
    static final List<String> CP_LIST = new LinkedList<>();
    static final List<String> CPK_LIST = new LinkedList<>();
    static String cpkName = "Est. Cpk";
    static String cpName = "Est. Cp";
    static String dimensionName = "Dimension #";
    static String cavitiesName = "Cavities :";
    static String cavitiesValue = "";

    public static void main(String[] args) {
        final List<Map<Integer, String>> dimensionDataList = new LinkedList<>();
        final List<Map<Integer, String>> estCpkDataList = new LinkedList<>();
        final List<Map<Integer, String>> estCpDataList = new LinkedList<>();
        final List<Map<Integer, String>> cavitiesNameDataList = new LinkedList<>();
        EasyExcel.read("E:\\_Info\\Work\\HP\\upload\\7QR88-40126_13Jun2022_CPK.xlsm")
                .registerReadListener(new AnalysisEventListener<Map<Integer, String>>() {
                    @Override
                    public void invoke(Map<Integer, String> integerStringMap, AnalysisContext analysisContext) {
                        if (cavitiesName.equals(integerStringMap.get(0))) {
                            cavitiesNameDataList.add(integerStringMap);
                        }
                        if (dimensionName.equals(integerStringMap.get(0))) {
                            if (integerStringMap.size() < 45) {
                                dimensionDataList.add(integerStringMap);
                            }
                        }
                        if (cpkName.equals(integerStringMap.get(0))) {
                            estCpkDataList.add(integerStringMap);
                        }
                        if (cpName.equals(integerStringMap.get(0))) {
                            estCpDataList.add(integerStringMap);
                        }
                    }

                    @Override
                    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
                    }
                })
                .doReadAll();
        getDataProcessing(cavitiesNameDataList, "");
        getDataProcessing(dimensionDataList, "dimension");
        getDataProcessing(estCpDataList, "cp");
        getDataProcessing(estCpkDataList, "cpk");
        getCavitiesName(CAVITIES_LIST);
        getDimensionDataProcessing(CP_INX_LIST, "CP");
        getDimensionDataProcessing(CPK_INX_LIST, "CPK");
    }

    public static void getCavitiesName(List<String> cavitiesList) {
        if (cavitiesList.size() > 2) {
            cavitiesValue = cavitiesList.get(2);
        }
    }

    public static void getDataProcessing(List<Map<Integer, String>> list, String name) {
        for (Map<Integer, String> integerStringMap: list) {
            for (Integer key : integerStringMap.keySet()) {
                if (key > 0 && key < 16) {
                    if (Objects.equals(name, "cpk")) {
                        double num;
                        if (Objects.equals(integerStringMap.get(key), null) ||
                                Objects.equals(integerStringMap.get(key), "N/A")) {
                            num = 2;
                        } else {
                            num = Double.parseDouble(integerStringMap.get(key));
                        }
                        if (num < 1.5) {
                            CPK_INX_LIST.add(key);
                            CPK_LIST.add(integerStringMap.get(key));
                        }

                    } else if (Objects.equals(name, "cp")) {
                        double num;
                        if (Objects.equals(integerStringMap.get(key), null) ||
                                Objects.equals(integerStringMap.get(key), "N/A")) {
                            num = 2;
                        } else {
                            num = Double.parseDouble(integerStringMap.get(key));
                        }
                        if (num < 1.3) {
                            CP_INX_LIST.add(key);
                            CP_LIST.add(integerStringMap.get(key));
                        }
                    } else if (Objects.equals(name, "dimension")){
                        DIMENSION_LIST.add(integerStringMap.get(key));
                    } else {
                        if (!Objects.equals(integerStringMap.get(key), null)) {
                            CAVITIES_LIST.add(integerStringMap.get(key));
                        }
                    }
                }
            }
        }

    }

    public static void getDimensionDataProcessing(List<Integer> list, String name) {
        int i = 0;
        for (int integer : list) {
            String cpStr = "Cavity:"+ cavitiesValue +"; "
                    + "Dimension:"+DIMENSION_LIST.get((integer-1))
                    + "; "+ name +" out spec; "+ name +":"
                    + (Objects.equals(name, "cp") ? CP_LIST.get(i) : CPK_LIST.get(i))
                    + "; " + name +" Target:"
                    + (Objects.equals(name, "cp") ? "1.5" : "1.33") + ";";
            System.out.println(cpStr);
            i++;
        }
    }
}
