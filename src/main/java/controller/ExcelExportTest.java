package controller;

import Utils.ExcelUtil;
import com.alibaba.fastjson.JSONArray;
import domain.Student;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;

/**
 * Created by Cheung on 2017/12/19.
 */
public class ExcelExportTest {

	public static void main(String[] args) throws IOException {
		// 模拟10W条数据
		int count = 100000;
	 	JSONArray studentArray = new JSONArray();
		for (int i = 0; i < count; i++) {
			Student s = new Student();
			s.setName("POI" + i);
			s.setAge(i);
			s.setBirthday(new Date());
			s.setHeight(i);
			s.setWeight(i);
			s.setSex(i % 2 == 0 ? false : true);
			studentArray.add(s);
		}

		ArrayList<LinkedHashMap> titleList = new ArrayList<LinkedHashMap>();
		LinkedHashMap<String, String> titleMap = new LinkedHashMap<String, String>();
		titleMap.put("title1","POI导出大数据量Excel Demo");
		titleMap.put("title2","https://github.com/550690513");
		LinkedHashMap<String, String> headMap = new LinkedHashMap<String, String>();
		headMap.put("name", "姓名");
		headMap.put("age", "年龄");
		headMap.put("birthday", "生日");
		headMap.put("height", "身高");
		headMap.put("weight", "体重");
		headMap.put("sex", "性别");
		titleList.add(titleMap);
		titleList.add(headMap);



		File file = new File("D://ExcelExportDemo/");
		if (!file.exists()) file.mkdir();// 创建该文件夹目录
		OutputStream os = null;
		Date date = new Date();
		try {
			// .xlsx格式
			os = new FileOutputStream(file.getAbsolutePath() + "/" + date.getTime() + ".xlsx");
			System.out.println("正在导出xlsx...");
			ExcelUtil.exportExcel(titleList, studentArray, os);
			System.out.println("导出完成...共" + count + "条数据,用时" + (System.currentTimeMillis() - date.getTime()) + "ms");
			System.out.println("文件路径：" + file.getAbsolutePath() + "/" + date.getTime() + ".xlsx");
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			os.close();
		}

	}
}
