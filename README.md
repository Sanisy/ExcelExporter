# ExcelExporter
making the operation exporter bean data to excel more easier

###Example

```
@Data
public class User {

    @ExcelColumn(index="0", title = "姓名")
    private String userName;

    @ExcelColumn(index="1", title = "手机号码", dataPipeline = TestCellDataPipeline.class)
    private String mobile;

    @ExcelColumn(index="2", title = "地址")
    private String address;
}


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
```
