package com.simple.exception;

/**
 * 实体中不存在ExcelEntry注解
 * 无法使用注解加载
 */
public class NoExcelEntryAnnotationsException extends  Exception {
    public NoExcelEntryAnnotationsException(){
        super();
    }
    public NoExcelEntryAnnotationsException(String args){
        super(args);
    }
}
