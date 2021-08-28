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
#### 程序运行结果
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
		"age":18,
		"name":"dc1",
		"no":"NO1",
		"salary":1000.11
	},
	{
		"age":19,
		"name":"dc2",
		"no":"NO2",
		"salary":1001.11
	},
	{
		"age":20,
		"name":"dc3",
		"no":"NO3",
		"salary":1002.11
	},
	{
		"age":21,
		"name":"dc4",
		"no":"NO4",
		"salary":1003.11
	}
]

```
