package com.sanisy.excel;

/**
 * data converter interface, first you should be know the cell data's type is always String.
 * When the bean's field's data type is int or other primitive type, you should convert the data to String.
 * It can help us avoid of the data type error when we invoke method cell.setCellValue(). for example, the cell require int value,
 * but the real data is a String or Object,
 * Created by Sanisy on 2018/4/18.
 */
public interface ColumnDataConverter {

    String convert(Object data);
}
