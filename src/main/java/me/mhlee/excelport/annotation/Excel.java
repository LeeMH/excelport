package me.mhlee.excelport.annotation;

import me.mhlee.excelport.cellstyle.Align;
import me.mhlee.excelport.cellstyle.DateFormat;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;


@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface Excel {
    String name() default "";
    int order() default 999;
    Align align() default Align.NONE;
    DateFormat format() default DateFormat.NONE;
}
