package com.sanisy.excel;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * announce the field of the data bean that should be exporter to excel file
 * Created by Sanisy on 2018/4/18.
 */
@Retention( RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelColumn {

    String index();

    String title();

    Class<?> converter() default DefaultColumnDataConverter.class;

    Class<?> dataPipeline() default Void.class;
}
