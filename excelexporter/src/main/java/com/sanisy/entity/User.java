package com.sanisy.entity;

import com.sanisy.excel.ExcelColumn;
import lombok.Data;

@Data
public class User {

    @ExcelColumn(index="0", title = "姓名")
    private String userName;

    @ExcelColumn(index="1", title = "手机号码", dataPipeline = TestCellDataPipeline.class)
    private String mobile;

    @ExcelColumn(index="2", title = "地址")
    private String address;
}
