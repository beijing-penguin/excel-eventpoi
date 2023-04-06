# 高性能事件驱动  eventpoi 支持poi4.x以上

#### 功能介绍
- 支持 【功能】 支持占位符公式，如${money/100}
- 支持 【功能】 支持四舍五入公式，如${round(money/100,2)}对计算后的数值，按四舍五入保留两位小数
- 支持 【功能】 支持保留小数（不四舍五入），如${truncate(money/100,2)}对计算后的数值，按保留两位小数，不四舍五入
- 支持 【功能】 【读取excel/导入】excel文件高性能读取 并自动转java对象，只需一行代码
- 支持 【功能】 支持导出表格形式的excel文件，和列表excel文件
- 支持 【功能】 【图片导入】支持图片列和文本混合列形式的excel导入
- 支持 【功能】 【行回调异步处理】边读边处理，无须等待excel全部解析。
- 支持 【功能】 自动识别xls文件和xlsx文件格式
- 支持 【功能】 自动识别date列时间格式，无需像别的工具那样，时间字段上设置yyyyMMdd格式
- 支持 【功能】 复杂的表格导出，既包含【列表形式】的数据 也包含【表单形式】的数据导出（查看 "测试包含表格和列表数据的复杂导出.java"  文件）
- 支持 【特性】 使用事件驱动方式读取excel，无行数和内存限制。
- 支持 【特性】 高性能SXSSFWorkbook导出大文件，默认也是使用SXSSFWorkbook导出（导出20w行数据，仅需要3s）
- 支持 【案例】 查看src/test/java目录

#### 使用注意事项
- 导出复杂结构的excel时，**模板内容** 除了${xxxx}占位符之外，禁止${符号的内容再出现，原因是程序目前只按${解析模板
   - 后期可优化为用户指定占位符
- 所有占位符不可重复
- 导入图片，或者 导出图片 时，java对象中的属性要是byte[]类型，如``` private byte[] headIamge; ```

#### 有使用问题，请提交issues，或者加QQ群591170125反馈

#### 使用案例见/eventpoi/src/test/java目录
#### 新建模板excel文件
![image](https://user-images.githubusercontent.com/10703753/142003702-2c8b09a5-84e4-4025-bfc9-96d7d8b34d79.png)

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
        
        byte[] exportByteData = ExcelHelper.exportExcel(Test1.class.getResourceAsStream("demo1Templete.xlsx"), personList, "${salary}");
        Files.write(Paths.get("./my_test_temp/测试导出指定对象并删除指定列.xlsx"), exportByteData);
    }
}

```
#### 复杂结构的excel导出(都在src/test/java中 ` 测试包含表格和列表数据的复杂导出.java `)，如图
![image](https://user-images.githubusercontent.com/10703753/142004207-2e863b7a-a4e1-49cc-8295-89de6028b89c.png)

### excel模板 ` ./my_test_temp/测试包含表格和列表数据的复杂导出.xlsx `
```java
public class 测试包含表格和列表数据的复杂导出 {
    public static void main(String[] args) throws Exception {
    	//构造表格形式的数据
    	OrderInfo orderInfo = new OrderInfo();
    	orderInfo.setKehu("ddddcccc");
    	orderInfo.setOrderName("进口海鲜");
    	orderInfo.setTotalMoney("15.66");
    	orderInfo.setBuyer("张三");
    	orderInfo.setSaller("李四");
    	
    	//构造列表形式的数据
        List<Object>  excelDataList = new ArrayList<Object>();
        List<ProductInfo> productList = new ArrayList<ProductInfo>();
        ProductInfo p1 = new ProductInfo();
        p1.setNo("NO_1");
        p1.setName("ssssss111");
        
        ProductInfo p2 = new ProductInfo();
        p2.setNo("NO_2");
        p2.setName("ssssss222");
        
        ProductInfo p3 = new ProductInfo();
        p3.setNo("NO_3");
        p3.setName("ssssss333");
        
        productList.add(p1);
        productList.add(p2);
        productList.add(p3);
        
        excelDataList.add(productList);
        excelDataList.add(orderInfo);

        byte[] exportByteData = ExcelHelper.exportTableExcel(Test1.class.getResourceAsStream("订单_templete.xlsx"), excelDataList);
        Files.write(Paths.get("./my_test_temp/测试包含表格和列表数据的复杂导出.xlsx"), exportByteData);
    }
}

```
