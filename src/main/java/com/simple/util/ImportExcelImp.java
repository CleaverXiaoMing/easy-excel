package com.simple.util;

import com.simple.annotations.ExcelEntry;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.beans.IntrospectionException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.*;


public class ImportExcelImp implements  ImportExcel{

    public List<Object> toList(Class clazz, String fileName,int beginRowNum) throws IOException, IllegalAccessException, InstantiationException,InvocationTargetException {
        if(fileName.endsWith(".xlsx")){
            return toListExcel2007(clazz,fileName,beginRowNum);
        }
        List<Object> objList = new ArrayList<Object>();
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(new File(fileName)));
        HSSFSheet sheet = null;
        Map<String,Object> beanMap = new HashMap<String,Object>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {// 获取每个Sheet表
            sheet = workbook.getSheetAt(i);
            for (int j = beginRowNum; j < sheet.getLastRowNum() + 1; j++) {// getLastRowNum，获取最后一行的行标
                Object obj = clazz.newInstance();
                HSSFRow row = sheet.getRow(j);
                if (row != null) {
                    BeanUtils.populate(obj,getBeanMap(row,clazz.getDeclaredFields()));
                }
                objList.add(obj);
            }
        }
        return objList;
    }
    public List<Object> toListExcel2007(Class clazz, String fileName,int beginRowNum) throws IOException, IllegalAccessException, InstantiationException,InvocationTargetException {
        if(fileName.endsWith(".xls")){
            return toList(clazz,fileName,beginRowNum);
        }
        List<Object> objList = new ArrayList<Object>();
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(new File(fileName)));
        Sheet sheet = null;
        Map<String,Object> beanMap = new HashMap<String,Object>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {// 获取每个Sheet表
            sheet = workbook.getSheetAt(i);
            for (int j = beginRowNum; j < sheet.getLastRowNum() + 1; j++) {// getLastRowNum，获取最后一行的行标
                Object obj = clazz.newInstance();
                Row row = sheet.getRow(j);
                if (row != null) {
                    BeanUtils.populate(obj,getBeanMap(row,clazz.getDeclaredFields()));
                }
                objList.add(obj);
            }
        }
        return objList;
    }

    public List<Map<String, Object>> toListMap(Class clazz, String fileName,int beginRowNum) throws IOException {
        if(fileName.endsWith(".xlsx")){
            return toListMapExcel2007(clazz,fileName,beginRowNum);
        }
        List<Map<String, Object>> mapList = new ArrayList<Map<String, Object>>();
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(new File(fileName)));
        HSSFSheet sheet = null;
        Map<String,Object> beanMap = new HashMap<String,Object>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {// 获取每个Sheet表
            sheet = workbook.getSheetAt(i);
            for (int j = beginRowNum; j < sheet.getLastRowNum() + 1; j++) {// getLastRowNum，获取最后一行的行标
                HSSFRow row = sheet.getRow(j);
                if (row != null) {
                    mapList.add(getBeanMap(row,clazz.getDeclaredFields()));
                }
            }
        }
        return mapList;
    }

    public  List<Map<String, Object>> toListMapExcel2007(Class clazz, String fileName,int beginRowNum) throws IOException {
        if(fileName.endsWith(".xls")){
            return toListMap(clazz,fileName,beginRowNum);
        }
        List<Map<String, Object>> mapList = new ArrayList<Map<String, Object>>();
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(new File(fileName)));
        Sheet sheet = null;
        Map<String,Object> beanMap = new HashMap<String,Object>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {// 获取每个Sheet表
            sheet = workbook.getSheetAt(i);
            for (int j = beginRowNum; j < sheet.getLastRowNum() + 1; j++) {// getLastRowNum，获取最后一行的行标
                Row row = sheet.getRow(j);
                if (row != null) {
                    mapList.add(getBeanMap(row,clazz.getDeclaredFields()));
                }
            }
        }
        return mapList;
    }

    public Map<String,Object> getBeanMap(Row row, Field[] fields){
        Map<String,Object> beanMap = new HashMap<String,Object>();
        for(Field f : fields) {
            if (f.isAnnotationPresent(com.simple.annotations.ExcelEntry.class)) {
                ExcelEntry excelEntry = f.getAnnotation(ExcelEntry.class);
                beanMap.put(f.getName(),row.getCell(excelEntry.columNum()));
            }
        }
        return beanMap;
    }
}
