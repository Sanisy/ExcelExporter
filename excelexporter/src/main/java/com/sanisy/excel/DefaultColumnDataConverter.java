package com.sanisy.excel;

/**
 * the cell data's default data converter
 * Created by Sanisy on 2018/4/18.
 */
public class DefaultColumnDataConverter implements ColumnDataConverter{


    public String convert(Object data) {
        if (data == null) {
            return "";
        }

        return data.toString();
    }
}
