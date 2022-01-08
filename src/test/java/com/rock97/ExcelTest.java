package com.rock97;

import java.util.ArrayList;
import java.util.List;

public class ExcelTest {
    public static void main(String[] args) {

        List<UserBean> userBeanList = new ArrayList<>();

        for (int i = 0; i < 100000; i++) {
            UserBean userBean = new UserBean();
            userBean.setAge(i%100);
            userBean.setName("testname"+i);
            userBean.setSex(i%2 + "");
            userBeanList.add(userBean);
        }

        ExcelUtil excelUtil = ExcelUtil.newInstance(UserBean.class, "测试导出");
        excelUtil.addList(userBeanList);
        excelUtil.write();
    }
}
