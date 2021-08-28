# 高性能事件驱动  eventpoi 目前开源版本只支持poi4.x以上，不支持3.x

#### 功能介绍
- 支持 excel文件高性能读取 并自动转java对象，只需一行代码即可
- 特点：性能高、支持excel大数据量读取、LruCache进一步提高excel读取性能
- 支持 指定工作簿sheetIndex读取
- 支持 自动识别xls文件  和 xlsx 文件格式
- 支持 自动识别date列时间格式
- 支持 事件流ExcelEventStream行回调异步实时读取文件（无论文件有多大，都不会直接占用内存，而是异步小批量缓冲流的方式抓取）
- 支持 java List集合对象转excel模板数据导出

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