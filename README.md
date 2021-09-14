# 高性能事件驱动  eventpoi 目前开源版本只支持poi4.x以上，不支持3.x

#### 功能介绍
- 支持 【功能】 【读取excel/导入】excel文件高性能读取 并自动转java对象，只需一行代码
- 支持 【功能】 支持导出表格形式的excel文件，和列表excel文件
- 支持 【功能】 【图片导入】支持图片列和文本混合列形式的excel导入
- 支持 【功能】 【行回调异步处理】边读边处理，无须等待excel全部解析。示例文件查看ExcelEventStreamTest.java文件
- 支持 【功能】 自动识别xls文件和xlsx文件格式
- 支持 【功能】 自动识别date列时间格式，无需像别的工具那样，时间字段上设置yyyyMMdd格式
- 支持 【功能】 使用事件驱动方式读取excel，无行数和内存限制。
- 支持 【功能】 复杂的表格导出，既包含列表形式的数据 也包含列表形式数据的导出
- 支持 【案例】 查看src/test/java目录

#### 使用案例见/eventpoi/src/test/java目录
#### 使用一行代码读取excel
```java
public class 使用一行代码读取excel {
    public static void main(String[] args) throws Exception {
        List<Person> objList = ExcelHelper.parseExcelToObject(Test1.class.getResourceAsStream("demo1.xlsx"), Test1.class.getResourceAsStream("demo1Templete.xlsx"), Person.class);
        System.err.println(JSON.toJSONString(objList,true));
    }
}
```

#### 使用一行代码导出excel
```java

public class 使用一行代码导出excel {
    public static void main(String[] args) throws Exception {
        List<Person> personList = new ArrayList<Person>();
        Person p1 = new Person();
        p1.setAge(11);
        p1.setName("ssssss111");
        p1.setRemark("测试测试啊remark111");
        
        Person p2 = new Person();
        p2.setAge(22);
        p2.setName("ssssss222");
        p2.setRemark("测试测试啊remark2222");
        personList.add(p1);
        personList.add(p2);
        
        //第三个参数表示，导出时，删除那些列（按模板文件中的key删除，可不传）
        byte[] exportByteData = ExcelHelper.exportExcel(Test1.class.getResourceAsStream("demo1Templete.xlsx"), personList, "${salary}");
        Files.write(Paths.get("./my_test_temp/测试导出指定对象并删除指定列.xlsx"), exportByteData);
    }
}

```