package com.simple.annotations;


import java.lang.annotation.*;

@Documented
@Target(ElementType.FIELD)
@Inherited
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelEntry {
    /**
     * 列名
     * 导出到excel对应的列名
     */
    String titleName() default "A列";

    /**
     * 列号
     * 从0开始
     */
    int columNum() default 0;
}
