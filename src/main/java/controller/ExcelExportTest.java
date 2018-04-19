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
 *
 * @author Cheung
 * @version 2.0.0
 * @date 2018/4/19
 */
public class ExcelExportTest {

	public static void main(String[] args) throws IOException {
		// 模拟100W条数据,存入JsonArray,此处使用fastJson(号称第一快json解析)快速解析大数据量数据
		// 至于容量问题,Java数组的length必须是非负的int，所以它的理论最大值就是java.lang.Integer.MAX_VALUE = 2^31-1 = 2147483647。
		// 由于xlsx最大支持行数为1048576行,此处模拟了1048573调数据,剩下的3条占用留给自定义的excel的头信息和列项.
		int count = 1048573;
		JSONArray studentArray = new JSONArray();
		for (int i = 0; i < count; i++) {
			Student s = new Student();
			s.setName("POI-" + i);
			s.setAge(i);
			s.setBirthday(new Date());
			s.setHeight(i);
			s.setWeight(i);
			s.setSex(i % 2 == 0 ? false : true);
			studentArray.add(s);
		}

		/*
		 * titleList存放了2个元素,分别为titleMap和headMap
		 */
		ArrayList<LinkedHashMap> titleList = new ArrayList<LinkedHashMap>();
		// 1.titleMap存放了该excel的头信息
		LinkedHashMap<String, String> titleMap = new LinkedHashMap<String, String>();
		titleMap.put("title1", "POI导出大数据量Excel Demo");
		titleMap.put("title2", "https://github.com/550690513");
		// 2.headMap存放了该excel的列项
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
		if (!file.exists()) file.mkdirs();// 创建该文件夹目录
		OutputStream os = null;
		try {
			System.out.println("正在导出xlsx...");
			long start = System.currentTimeMillis();
			// .xlsx格式
			os = new FileOutputStream(file.getAbsolutePath() + File.separator + start + ".xlsx");
			ExcelUtil.exportExcel(titleList, studentArray, os);
			System.out.println("导出完成...共" + count + "条数据,用时" + (System.currentTimeMillis() - start) + "毫秒");
			System.out.println("文件路径：" + file.getAbsolutePath() + File.separator + start + ".xlsx");
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			os.close();
		}

	}
}
