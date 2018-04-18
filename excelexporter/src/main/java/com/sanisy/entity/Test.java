package com.sanisy.entity;

import com.sanisy.excel.ExcelExporter;

import java.util.ArrayList;
import java.util.List;

public class Test {

    public static void main(String[] args) {
        List<User> userList = new ArrayList<User>();
        User user0 = new User();
        user0.setUserName("SteveJobs");
        user0.setAddress("City0");

        User user1 = new User();
        user1.setUserName("Wozniak");
        user1.setAddress("City1");

        User user2 = new User();
        user2.setUserName("JeffBezos");
        user2.setAddress("City3");

        userList.add(user0);
        userList.add(user1);
        userList.add(user2);

        ExcelExporter excelExporter = new ExcelExporter.Builder().exportTo("D:/exporter.xlsx", "mySheet0").build();
        excelExporter.doExport(userList);
    }
}
