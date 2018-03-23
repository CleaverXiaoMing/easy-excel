package com.simple.entry;

import com.simple.annotations.ExcelEntry;

public class Student {

    @ExcelEntry(titleName="学号",columNum = 1)
    String sno;
    @ExcelEntry(titleName="姓名",columNum = 0)
    String name;

    String sex;

    String grade;

    @ExcelEntry(titleName="年龄",columNum = 2)
    int age;

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public Student(String name, String sno,int age) {
        this.name = name;
        this.sno = sno;
        this.age = age;
    }
    public Student(){

    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public String getSno() {
        return sno;
    }

    public void setSno(String sno) {
        this.sno = sno;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    @Override
    public String toString() {
        return "Student{" +
                "sno='" + sno + '\'' +
                ", name='" + name + '\'' +
                ", sex='" + sex + '\'' +
                ", grade='" + grade + '\'' +
                ", age=" + age +
                '}';
    }
}
