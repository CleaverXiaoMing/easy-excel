package com.simple.util;

import com.simple.annotations.ExcelEntry;
import com.simple.exception.AnnotationsColumReuseException;
import com.simple.exception.NoExcelEntryAnnotationsException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.beans.PropertyDescriptor;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;

public class ExportExcelImp implements ExportExcel{


    private short excelType;

    public ExportExcelImp(){
        this.excelType = EXCEL_2003;
    }
    public ExportExcelImp(short excelType){
        this.excelType = excelType;
    }

    public void setExcelType(short excelType) {
        this.excelType = excelType;
    }



    public String[] exportMapExcel(String sheetName, String tableHeadName, List<String> titleList, Map<String,String> titleMapper,
                                   List<Map<String,Object>> contentList, String fileName, int columWidth) throws IOException {
        /**
         * sheetName sheet名称
         * tableHeadName 表名
         * titleList 表头列名
         * titleMapper 表头与contentList/titleList映射 key为中午表头，value为contentList中对应的key值
         * contentList 内容数组
         * columWidth 列宽
         */
        List<String> fileList = new ArrayList<String>();
        int wbs = contentList.size()/MAX_ROW_NUM;
        for(int wbNum=0;wbNum<=wbs;wbNum++) {
            HSSFWorkbook wb = new HSSFWorkbook();
            //建立新的sheet对象（excel的表单）
            HSSFSheet sheet = wb.createSheet(sheetName);

            //在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个
            HSSFRow row1 = sheet.createRow(0);
            //创建单元格（excel的单元格，参数为列索引，可以是0～255之间的任何一个
            HSSFCell cell = row1.createCell(0);

            //设置单元格内容
            cell.setCellValue(tableHeadName);
            cell.getRow().setHeightInPoints(30.2f);
            HSSFCellStyle style = wb.createCellStyle();
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//水平居中
            style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直居中
            HSSFFont font = wb.createFont();
            font.setFontHeightInPoints((short) 15);
            font.setColor(HSSFColor.BLACK.index);
            font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
            font.setFontName("宋体");
            // 把字体 应用到当前样式
            style.setFont(font);
            cell.setCellStyle(style);
            //创建空白行
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, titleList.size() - 1));
            //创建表名行

            //设置标题行
            HSSFRow row2 = sheet.createRow(1);
            for (int i = 0; i < titleList.size(); i++) {
                row2.createCell(i).setCellValue(titleList.get(i));
                sheet.setColumnWidth(i, WIDTH_PARAM * columWidth + WIDTH_PARAM_ADD);
            }
            //设置数据行
            if(wbNum==wbs){
                for (int j = wbNum * MAX_ROW_NUM; j < contentList.size(); j++) {
//                    int rowNum = j -( wbNum * MAX_ROW_NUM) + 2;
                    HSSFRow row3 = sheet.createRow(j -( wbNum * MAX_ROW_NUM) + 2);
                    for (int p = 0; p < titleList.size(); p++) {
                        row3.createCell(p).setCellValue(contentList.get(j).get(titleMapper.get(titleList.get(p))) + "");
                    }
                }
            }else {
                for (int j = 0; j < MAX_ROW_NUM; j++) {
                    HSSFRow row3 = sheet.createRow(j + 2);
                    for (int p = 0; p < titleList.size(); p++) {
                        row3.createCell(p).setCellValue(contentList.get(j + wbNum * MAX_ROW_NUM).get(titleMapper.get(titleList.get(p))) + "");
                    }
                }
            }
            //保存文件
            FileOutputStream fos = new FileOutputStream(fileName+wbNum+".xls");
            wb.write(fos);
            fos.close();
            if(wbNum>0){
                fileList.add(fileName+wbNum+".xls");
            }else{
                fileList.add(fileName+".xls");
            }

        }
        String [] fileArr = fileList.toArray(new String[0]);
        return fileArr;
    }

    public String[] exportTo2007(String sheetName, String tableHeadName, List<String> titleList, Map<String,String> titleMapper,
                               List<Map<String,Object>> contentList, String fileName, int columWidth) throws IOException {
        SXSSFWorkbook sb = new SXSSFWorkbook();
        //建立新的sheet对象（excel的表单）
        Sheet sheet = sb.createSheet(sheetName);

        //在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个
        Row row1 = sheet.createRow(0);
        //创建单元格（excel的单元格，参数为列索引，可以是0～255之间的任何一个
        Cell cell = row1.createCell(0);

        //设置单元格内容
        cell.setCellValue(tableHeadName);
        cell.getRow().setHeightInPoints(30.2f);
        CellStyle style = sb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//水平居中
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直居中
        Font font = sb.createFont();
        font.setFontHeightInPoints((short) 15);
        font.setColor(HSSFColor.BLACK.index);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        font.setFontName("宋体");
        // 把字体 应用到当前样式
        style.setFont(font);
        cell.setCellStyle(style);
        //创建空白行
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, titleList.size() - 1));
        //创建表名行

        //设置标题行
        Row row2 = sheet.createRow(1);
        for (int i = 0; i < titleList.size(); i++) {
            row2.createCell(i).setCellValue(titleList.get(i));
            sheet.setColumnWidth(i, WIDTH_PARAM * columWidth + WIDTH_PARAM_ADD);
        }
        //设置数据行
        for (int j = 0; j < contentList.size(); j++) {
            Row row3 = sheet.createRow(j + 2);
            for (int p = 0; p < titleList.size(); p++) {
                row3.createCell(p).setCellValue(contentList.get(j).get(titleMapper.get(titleList.get(p))) + "");
            }
        }
        //保存文件
        FileOutputStream fos = new FileOutputStream(fileName+".xlsx");
        sb.write(fos);
        fos.close();
        String[] fileNames = new String[1];
        fileNames[0] = fileName+".xlsx";
        return fileNames;
    }
    public String[] exportMapExcel(List<Map<String,Object>> contentList,Class clazz) throws IOException, NoExcelEntryAnnotationsException, AnnotationsColumReuseException {
        String fileName = new Date().getTime()+"";
        return exportMapExcel(contentList,clazz,fileName);

    }

    public String[] exportMapExcel(List<Map<String,Object>> contentList,Class clazz,String fileName) throws IOException, NoExcelEntryAnnotationsException, AnnotationsColumReuseException {
        return exportMapExcel(contentList,clazz,fileName,DEFALUT_SHEET_NAME);
    }

    public String[] exportMapExcel(List<Map<String,Object>> contentList,Class clazz,String fileName,String sheetName) throws IOException, NoExcelEntryAnnotationsException, AnnotationsColumReuseException {
        return exportMapExcel(contentList,clazz,fileName,sheetName,DEFALUT_TABLE_NAME);
    }

    public String[] exportMapExcel(List<Map<String,Object>> contentList,Class clazz,String fileName,String sheetName,String tableName) throws IOException, NoExcelEntryAnnotationsException, AnnotationsColumReuseException {
        return exportMapExcel(contentList,clazz,fileName,sheetName,tableName,DEFALUT_WIDTH);
    }

    public String[] exportMapExcel(List<Map<String,Object>> contentList,Class clazz,String fileName,String sheetName,String tableName,int width) throws IOException, NoExcelEntryAnnotationsException, AnnotationsColumReuseException {
        Map<Integer,String> titleMapper = new HashMap<Integer,String>();
        Map<String,String> mapper = new HashMap<String, String>();
        List tiTleList = new ArrayList();
        int size=0;
        for (Field field : clazz.getDeclaredFields()){
            if (field.isAnnotationPresent(com.simple.annotations.ExcelEntry.class)) {
                size++;
                try {
                    ExcelEntry excelEntry = field.getAnnotation(ExcelEntry.class);
                    mapper.put(excelEntry.titleName(), field.getName());
                    titleMapper.put(excelEntry.columNum(), excelEntry.titleName());
                } catch (Throwable ex) {
                    ex.printStackTrace();
                }
            }
        }
        if(size==0){
            throw new NoExcelEntryAnnotationsException();
        }
        if(titleMapper.size()<size){
            throw new AnnotationsColumReuseException();
        }
        for(int i=0;i<titleMapper.size();i++){
            tiTleList.add(i,titleMapper.get(i));
        }
        if(excelType==EXCEL_2003) {
            return exportMapExcel(sheetName, tableName, tiTleList, mapper, contentList, fileName, width);
        }else{
            return exportTo2007(sheetName, tableName, tiTleList, mapper, contentList, fileName, width);
        }
    }

    /**
     * defalut sheet:sheet1
     * return fileName
     *
     */
    public String[] exportClassExcel(List<Object> list) throws NoExcelEntryAnnotationsException, AnnotationsColumReuseException, IOException {
        String fileName = new Date().getTime()+"";
        return  exportClassExcel(list,fileName);
    }

    public String[] exportClassExcel(List<Object> list,String fileName) throws NoExcelEntryAnnotationsException, AnnotationsColumReuseException, IOException {
        return exportClassExcel(list,fileName,DEFALUT_SHEET_NAME);
    }

    public String[] exportClassExcel(List<Object> list,String fileName,String sheetName) throws NoExcelEntryAnnotationsException, AnnotationsColumReuseException, IOException {
        return exportClassExcel(list,fileName,sheetName,DEFALUT_TABLE_NAME);
    }

    public String[] exportClassExcel(List<Object> list,String fileName,String sheetName,String tableName) throws NoExcelEntryAnnotationsException, AnnotationsColumReuseException, IOException {
        return exportClassExcel(list,fileName,sheetName,tableName,DEFALUT_WIDTH);
    }

    public String[] exportClassExcel(List<Object> list,String fileName,String sheetName,String tableName,int width) throws NoExcelEntryAnnotationsException, AnnotationsColumReuseException, IOException {
        //通过反射获取注解信息
        int size = 0;
        List tiTleList = new ArrayList();
        List<Map<String,Object>> contentList = new ArrayList<Map<String, Object>>();
        Map<Integer,String> titleMapper = new HashMap<Integer,String>();
        Map<String,String> mapper = new HashMap<String, String>();
        for(int i=0;i<list.size();i++){
            Object object = list.get(i);
            Map<String,Object> objMap = new HashMap<String, Object>();
            for (Field field : object.getClass().getDeclaredFields()){
                if (field.isAnnotationPresent(com.simple.annotations.ExcelEntry.class)) {
                    size++;
                    try {
                        if(i==0){
                            ExcelEntry excelEntry = field.getAnnotation(ExcelEntry.class);
                            mapper.put(excelEntry.titleName(), field.getName());
                            titleMapper.put(excelEntry.columNum(), excelEntry.titleName());
                        }
                        PropertyDescriptor pd=new PropertyDescriptor(field.getName(),object.getClass());
                        Method getMethod=pd.getReadMethod();
                        if(getMethod.invoke(object)!=null){
                            objMap.put(field.getName(),getMethod.invoke(object));
                        }
                    } catch (Throwable ex) {
                        ex.printStackTrace();
                    }
                }
            }
            if(i!=list.size()-1){
                size=0;
            }
            contentList.add(objMap);
        }

        if(size==0)
            throw new NoExcelEntryAnnotationsException();

        if(titleMapper.size()<size)
            throw new AnnotationsColumReuseException();

        for(int i=0;i<titleMapper.size();i++){
            tiTleList.add(i,titleMapper.get(i));
        }

        if(excelType==EXCEL_2003) {
            return exportMapExcel(sheetName, tableName, tiTleList, mapper, contentList, fileName, width);
        }else{
            return exportTo2007(sheetName, tableName, tiTleList, mapper, contentList, fileName, width);
        }
    }

}
