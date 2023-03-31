package com.hj.excellibrary.annotation;



import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface CellStyle {

    HorizontalAlignment horizontalAlign() default HorizontalAlignment.CENTER;

    VerticalAlignment verticalAlign() default VerticalAlignment.CENTER;

    boolean bold() default false;

    boolean italic() default false;

    FontUnderline underline() default FontUnderline.NONE;

}
