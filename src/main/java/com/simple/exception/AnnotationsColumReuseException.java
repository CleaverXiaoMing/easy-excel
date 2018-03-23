package com.simple.exception;

/**
 * ExcelEntry注解中colum值重复使用
 * colum注解存在重复的int值，titleName无法正确加载到对应列
 */
public class AnnotationsColumReuseException extends  Exception{
    public AnnotationsColumReuseException(){
        super();
    }
    public AnnotationsColumReuseException(String args){
        super(args);
    }
}
