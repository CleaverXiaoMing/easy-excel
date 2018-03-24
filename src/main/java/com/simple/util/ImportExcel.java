package com.simple.util;


import java.beans.IntrospectionException;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.List;
import java.util.Map;

public interface ImportExcel {
    //默认从第三行开始导入
    int BEGIN_ROW_NUM = 1;
    int BEGIN_ROW_NUM_WITH_HEAD = 2;

    List<Object> toList(Class clazz,String fileName,int beginRowNum) throws IOException, IllegalAccessException, InstantiationException, InvocationTargetException;

    List<Object> toListExcel2007(Class clazz, String fileName,int beginRowNum) throws IOException, IllegalAccessException, InstantiationException,InvocationTargetException;

    List<Map<String, Object>> toListMap(Class clazz, String fileName,int beginRowNum) throws IOException;

    List<Map<String, Object>> toListMapExcel2007(Class clazz, String fileName,int beginRowNum) throws IOException;

}
