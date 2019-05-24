# eventpoi 目前开源版本只支持poi4.x以上，不支持3.x，如果需要3.x的代码，私信我qq号码。
# 联系方式429544557@qq.com

# 支持 指定工作簿sheetIndex读取
# 支持 模板映射成实体对象读取文件
# 支持 兼容xls文件  和 xlsx 文件格式，代码编写上屏蔽了文件区别，只需一行代码即可读取文件，无序考虑文件类型
# 支持 事件流ExcelEventStream直接读取文件

# 使用案例 查阅测试类 com.dc.eventpoi.Test1.java如下
```java
public class Test1 {
	public static void main(String[] args) {
		/***********返回List<ExcelRow>类型数据*********************************/
		InputStream excelInput = Test1.class.getResourceAsStream("demo1.xlsx");
		System.out.println(excelInput);
		try {
			List<ExcelRow> dataList1 = ExcelHelper.parseExcelRowList(excelInput);//默认只读取所有工作簿数据，第二个参数指定工作簿
			//指定工作薄eg:
			//List<ExcelRow> dataList2 = ExcelHelper.parseExcelRowList(excelInput,0);//默认只读取sheetIndex=0的工作簿数据，第二个参数指定工作簿
			System.out.println("---------------------------excel转系统自带ExcelRow对象----------------------");
			System.out.println(JSON.toJSONString(dataList1,true));
			
			/***********返回一个自定义对象List<Person>类型数据，需要提前定义excel模板文件，如测试中的demo1Templete.xlsx*********************************/
			InputStream templeteInput = Test1.class.getResourceAsStream("demo1Templete.xlsx");
			List<ExcelRow> templeteList1 = ExcelHelper.parseExcelRowList(templeteInput);
			List<Person> objList = ExcelHelper.parseExcelToObject(dataList1, templeteList1, Person.class);
			System.out.println("---------------------------excel转对象----------------------");
			System.out.println(JSON.toJSONString(objList,true)); 
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
```
# 程序运行结果
```java
java.io.BufferedInputStream@15db9742
---------------------------excel转系统自带ExcelRow对象----------------------
[
	{
		"cellList":[
			{
				"index":0,
				"value":"编号"
			},
			{
				"index":1,
				"value":"姓名"
			},
			{
				"index":2,
				"value":"年龄"
			},
			{
				"index":3,
				"value":"工资"
			}
		],
		"rowIndex":0
	},
	{
		"cellList":[
			{
				"index":0,
				"value":"NO1"
			},
			{
				"index":1,
				"value":"dc1"
			},
			{
				"index":2,
				"value":"18"
			},
			{
				"index":3,
				"value":"1000.11"
			}
		],
		"rowIndex":1
	},
	{
		"cellList":[
			{
				"index":0,
				"value":"NO2"
			},
			{
				"index":1,
				"value":"dc2"
			},
			{
				"index":2,
				"value":"19"
			},
			{
				"index":3,
				"value":"1001.11"
			}
		],
		"rowIndex":2
	},
	{
		"cellList":[
			{
				"index":0,
				"value":"NO3"
			},
			{
				"index":1,
				"value":"dc3"
			},
			{
				"index":2,
				"value":"20"
			},
			{
				"index":3,
				"value":"1002.11"
			}
		],
		"rowIndex":3
	},
	{
		"cellList":[
			{
				"index":0,
				"value":"NO4"
			},
			{
				"index":1,
				"value":"dc4"
			},
			{
				"index":2,
				"value":"21"
			},
			{
				"index":3,
				"value":"1003.11"
			}
		],
		"rowIndex":4
	}
]
---------------------------excel转对象----------------------
[
	{
		"age":"18",
		"name":"dc1",
		"no":"NO1",
		"salary":"1000.11"
	},
	{
		"age":"19",
		"name":"dc2",
		"no":"NO2",
		"salary":"1001.11"
	},
	{
		"age":"20",
		"name":"dc3",
		"no":"NO3",
		"salary":"1002.11"
	},
	{
		"age":"21",
		"name":"dc4",
		"no":"NO4",
		"salary":"1003.11"
	}
]

```