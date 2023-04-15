package com.dc.eventpoi.test;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import com.dc.eventpoi.ExcelHelper;
import com.dc.eventpoi.core.PoiUtils;
import com.dc.eventpoi.core.entity.ListAndTableEntity;
import com.dc.eventpoi.core.inter.ExcelFunction;

public class 使用一行代码导出并使用function自定义公式 {
	public static void main(String[] args) throws Throwable {
		//注册你自定义的function，目前funclist为公开集合，方便注册操作
		ExcelHelper.funcList.add(new MyAddFunction());
		
		
		String img_file_path = new File(Me.class.getResource("unnamed.jpg").getPath()).getAbsolutePath();
		byte[] imageData = Files.readAllBytes(Paths.get(img_file_path));
		//写到本地，并查看图片
		Files.write(Paths.get("./my_test_temp/unnamed_temp.jpg"), imageData);

		List<Person> personList = new ArrayList<Person>();
		//构造导出时的数据
		for (int i = 0; i < 10000; i++) {

			Person p1 = new Person();
			p1.setNo("NO_"+i);
			p1.setAge(i);
			p1.setName("ssssss_"+i);
			p1.setRemark("测试测试啊remar_"+i);
			if(i==0) {
				p1.setHeadImage(imageData);
			}
			personList.add(p1);
		}

		Person p2 = new Person();
		p2.setName("李四");
		p2.setAge(4);
		
		//模拟一些需要删除某些列 的业务场景
		//byte[] newTempFile = PoiUtils.deleteCol(Me.class.getResourceAsStream("demo1Templete.xlsx"), "${salary}");
		
		byte[] tempData = PoiUtils.inputStreamToByte(Me.class.getResourceAsStream("demo1Templete.xlsx"));
		ListAndTableEntity dataEntity = ListAndTableEntity.build().setDataList(personList).setTable(p2);
		
		long t1 = System.currentTimeMillis();
		byte[] exportByteData = ExcelHelper.exportExcel(tempData, dataEntity,null,null,null,null);
		System.out.println("cost="+(System.currentTimeMillis()-t1));
		Files.write(Paths.get("./my_test_temp/测试导出指定对象并删除指定列.xlsx"), exportByteData);
	}
}

/**
 * 自定义一个加法函数，并自动讲null得数据变为0再相加
 * 占位符${myAdd(list.aaa,1)}
 * @author beijing-penguin
 *
 */
class MyAddFunction implements ExcelFunction{

	@Override
	public String getName() {
		return "myAdd";
	}

	@Override
	public Object execute(List<Object> paramValueList) {
		Object k1 = paramValueList.get(0);
		if(k1 == null) {
			k1 = 0;
		}
		Object k2 = paramValueList.get(1);
		return Long.valueOf(k1.toString())+Long.valueOf(k2.toString());
	}

}