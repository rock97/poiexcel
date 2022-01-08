package com.rock97.annotation;

import java.lang.annotation.*;

/**
 * @Description: 导出工具类枚举
 * 使用方法：@ExcelEnum(code = {"0","1"},name={"女","男"})
 * @Author: lizhihua16
 * @Email: lizhihua6@jd.com
 * @Create: 2018-04-19 10:08
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelEnum {
    String[] code();
    String[] name();
}
