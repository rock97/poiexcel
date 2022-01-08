package com.rock97;

import com.rock97.annotation.Excel;
import com.rock97.annotation.ExcelEnum;

public class UserBean {
    @Excel(value = "姓名",index = 2)
    private String name;
    @Excel(value = "年龄",index = 1)
    private int age;
    @Excel(value = "性别",index = 0)
    @ExcelEnum(code = {"0","1"},name = {"女","男"})
    private String sex;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }
}
