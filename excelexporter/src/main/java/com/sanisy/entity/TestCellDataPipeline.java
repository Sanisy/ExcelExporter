package com.sanisy.entity;

import com.sanisy.excel.CellDataPipeline;

import java.util.Random;

public class TestCellDataPipeline implements CellDataPipeline {

    public String dataOfCell(int index) {
        Random random = new Random();
        int i = random.nextInt(1000);
        return "testCell" + index + ":" + i;
    }
}
