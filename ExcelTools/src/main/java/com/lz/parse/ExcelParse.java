package com.lz.parse;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * @author qiliz
 */
public class ExcelParse {
    public static void main(String[] args) {
        final List<Map<Integer, String>> excelDataList = new LinkedList<>();
        final List<Map<Integer, String>> dimensionDataList = new LinkedList<>();
        final List<Map<Integer, String>> cpDataList = new LinkedList<>();
        final List<Map<Integer, String>> cpkDataList = new LinkedList<>();
        int count = 0;
        EasyExcel.read("E:\\_Info\\Work\\HP\\upload\\7QR88-40126_13Jun2022_CPK.xlsm")
                .sheet("P1")
                .registerReadListener(new AnalysisEventListener<Map<Integer, String>>() {
                    @Override
                    public void invoke(Map<Integer, String> integerStringMap, AnalysisContext analysisContext) {
                        if ("Dimension #".equals(integerStringMap.get(0))) {
                            dimensionDataList.add(integerStringMap);
                        }
                        excelDataList.add(integerStringMap);
                        System.out.println(integerStringMap);
                    }
                    @Override
                    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
                        System.out.println("数据读取完毕");
                    }
                })
                .doRead();
//        for (Map<Integer, String> integerStringMap : excelDataList) {
//            Set<Integer> keySet = integerStringMap.keySet();
//            count++;
//            for (Integer key : keySet) {
//                System.out.print(key + ":" + integerStringMap.get(key) + ", ");
//            }
//            System.out.println(count);
//            System.out.println("");
//        }
//        System.out.println(excelDataList);
        System.out.println(dimensionDataList);
    }
}
