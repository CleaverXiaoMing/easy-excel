package com.simple.util;

import com.simple.exception.AnnotationsColumReuseException;
import com.simple.exception.NoExcelEntryAnnotationsException;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public interface ExportExcel {
    int WIDTH_PARAM = 256;
    int WIDTH_PARAM_ADD = 184;
    String DEFALUT_SHEET_NAME = "sheet1";
    String DEFALUT_TABLE_NAME = "table";
    int DEFALUT_WIDTH = 15;
    int MAX_ROW_NUM = 60000;
    short EXCEL_2003 = 3;
    short EXCEL_2007 = 7;

    public void setExcelType(short excelType);
    String[] exportMapExcel(String sheetName, String tableHeadName, List<String> titleList, Map<String,String> titleMapper,
                               List<Map<String,Object>> contentList, String fileName, int columWidth)
            throws IOException ;
    String[] exportMapExcel(List<Map<String,Object>> contentList,Class clazz)
            throws IOException, NoExcelEntryAnnotationsException, AnnotationsColumReuseException;
    String[] exportMapExcel(List<Map<String,Object>> contentList,Class clazz,String fileName)
            throws IOException, NoExcelEntryAnnotationsException, AnnotationsColumReuseException;
    String[] exportMapExcel(List<Map<String,Object>> contentList,Class clazz,String fileName,String sheetName)
            throws IOException, NoExcelEntryAnnotationsException, AnnotationsColumReuseException;
    String[] exportMapExcel(List<Map<String,Object>> contentList,Class clazz,String fileName,String sheetName,String tableName)
            throws IOException, NoExcelEntryAnnotationsException, AnnotationsColumReuseException;
    String[] exportMapExcel(List<Map<String,Object>> contentList,Class clazz,String fileName,String sheetName,String tableName,int width)
            throws IOException, NoExcelEntryAnnotationsException, AnnotationsColumReuseException;
    String[] exportClassExcel(List<Object> list)
            throws NoExcelEntryAnnotationsException, AnnotationsColumReuseException, IOException;
    String[] exportClassExcel(List<Object> list,String fileName)
            throws IOException, NoExcelEntryAnnotationsException, AnnotationsColumReuseException;
    String[] exportClassExcel(List<Object> list,String fileName,String sheetName)
            throws IOException, NoExcelEntryAnnotationsException, AnnotationsColumReuseException;
    String[] exportClassExcel(List<Object> list,String fileName,String sheetName,String tableName)
            throws IOException, NoExcelEntryAnnotationsException, AnnotationsColumReuseException;
    String[] exportClassExcel(List<Object> list,String fileName,String sheetName,String tableName,int width)
            throws IOException, NoExcelEntryAnnotationsException, AnnotationsColumReuseException;
}
