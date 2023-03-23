package com.hj.excellibrary.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 标记field ,sheet 未注解的列将以 JsonArray String 写入读入该变量中。
 * 之作用在读取 excel
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelReadAggregate {
}
