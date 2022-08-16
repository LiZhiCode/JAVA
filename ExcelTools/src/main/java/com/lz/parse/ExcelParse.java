package com.lz.parse;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.util.*;

/**
 * @author qiliz
 */
public class ExcelParse {
    static final List<String> CPK_LIST = new LinkedList<>();
    static final List<String> CP_LIST = new LinkedList<>();
    static String cpkName = "Est. Cpk";
    static String cpName = "Est. Cp";
    public static void main(String[] args) {
        final List<Map<Integer, String>> excelDataList = new LinkedList<>();
        final List<Map<Integer, String>> estCpkDataList = new LinkedList<>();
        final List<Map<Integer, String>> estCpDataList = new LinkedList<>();
        EasyExcel.read("E:\\_Info\\Work\\HP\\upload\\7QR88-40126_13Jun2022_CPK.xlsm")
                .sheet("P1")
                .registerReadListener(new AnalysisEventListener<Map<Integer, String>>() {
                    @Override
                    public void invoke(Map<Integer, String> integerStringMap, AnalysisContext analysisContext) {
                        if (cpkName.equals(integerStringMap.get(0))) {
                            estCpkDataList.add(integerStringMap);
                        }
                        if (cpName.equals(integerStringMap.get(0))) {
                            estCpDataList.add(integerStringMap);
                        }
                    }
                    @Override
                    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
                        // System.out.println("数据读取完毕");
                    }
                })
                .doRead();
        getDataProcessing(estCpkDataList,"cpk");
        getDataProcessing(estCpDataList, "cp");
        System.out.println(CPK_LIST);
        System.out.println(CP_LIST);
    }
    public static <name> void getDataProcessing(List<Map<Integer, String>> estCpkDataList, String name) {
        for (Map<Integer, String> integerStringMap : estCpkDataList) {
            for (Integer key : integerStringMap.keySet()) {
                if ( key > 0 && key < 16) {
                    if (Objects.equals(name, "cpk")) {
                        CPK_LIST.add(integerStringMap.get(key));
                    } else {
                        CP_LIST.add(integerStringMap.get(key));
                    }
                }
            }
        }

    }
}
