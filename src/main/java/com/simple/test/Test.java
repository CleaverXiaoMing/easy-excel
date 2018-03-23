package com.simple.test;


import com.simple.entry.Student;
import com.simple.exception.AnnotationsColumReuseException;
import com.simple.exception.NoExcelEntryAnnotationsException;
import com.simple.util.ExportExcel;
import com.simple.util.ExportExcelImp;
import com.simple.util.ImportExcel;
import com.simple.util.ImportExcelImp;

import java.beans.IntrospectionException;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

public class Test {
    public static void main(String[] args){

        //导入测试

        //模拟从数据库取出list
        List<Object> stuList = new ArrayList<Object>();
        int i=0;
        while(i<2030) {
            i++;
            stuList.add(new Student("xiaoming", 1+(i+""),12));
            stuList.add(new Student("xiaoing", 2+(i+""),13));
            stuList.add(new Student("xiaomng", 3+(i+""),14));
        }
        //开始导出
        try {
            //创建导出工具类
            ExportExcel exportExcel = new ExportExcelImp();
            //设置生成excel格式
            exportExcel.setExcelType(ExportExcel.EXCEL_2007);
            //输入参数，开始导出，返回所有生成的文件名称
            String[] fileString = exportExcel.exportClassExcel(stuList,"stu");
            for (String str:fileString) {
                System.out.println(str);
            }
        } catch (NoExcelEntryAnnotationsException e) {
            e.printStackTrace();
        } catch (AnnotationsColumReuseException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        //导入测试
        try{
            //穿件导出工具类
            ImportExcel importExcel = new ImportExcelImp();

            //输入文件所在位置和起始行号开始导入
            List<Object> stuLis =  importExcel.toList(Student.class,"E:\\wokkspaceidea\\easy-excel\\stu.xlsx",ImportExcel.BEGIN_ROW_NUM);
            for(Object obj:stuLis){
                if(obj instanceof Student){
                    Student stu = (Student) obj;
                    System.out.println(stu.toString());
                }
            }
            System.out.println("共"+stuLis.size()+"条");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        }catch (IllegalAccessException e) {
            e.printStackTrace();
        }
    }
}
