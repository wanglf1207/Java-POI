package com.demo.poi;

import java.io.FileOutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelUtils {

    public static void main(String[] args) throws ParseException {

        /**
         * 创建Excel的实例
         * @throws ParseException
         */
        List list = new ArrayList();


        Map map = new HashMap();
        map.put("name", "王利峰");
        map.put("age", "32");
        map.put("desc", "失败者");

        Map map1 = new HashMap();
        map1.put("name", "王利峰");
        map1.put("age", "32");
        map1.put("desc", "失败者");

        list.add(map);
        list.add(map1);

        System.out.println(list);

        createExcel(list);

    }

    @SuppressWarnings("deprecation")
    public static void createExcel(List list) {

        // 第一步，创建一个webbook，对应一个Excel文件
        HSSFWorkbook wb = new HSSFWorkbook();
        // 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet = wb.createSheet("制卡费用核算报表");
        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
        HSSFRow row = sheet.createRow(0);
        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_LEFT); // 创建一个居中格式

        HSSFCell cell = row.createCell(0);

        cell.setCellValue("姓名");
        cell.setCellStyle(style);

        cell = row.createCell(1);
        cell.setCellValue("年龄");
        cell.setCellStyle(style);
        cell = row.createCell(2);
        cell.setCellValue("描述");
        cell.setCellStyle(style);


        // 第五步，写入实体数据 实际应用中这些数据从数据库得到，

        for (int i = 0; i < list.size(); i++)
        {
            HSSFRow rowContent = sheet.createRow(i+1);
            Map map = new HashMap();
            map = (Map) list.get(i);
            rowContent.createCell((short) 0).setCellValue(map.get("name").toString());
            rowContent.createCell((short) 1).setCellValue(map.get("age").toString());
            rowContent.createCell((short) 2).setCellValue(map.get("desc").toString());

        }
        // 第六步，将文件存到指定位置
        try
        {
            FileOutputStream fout = new FileOutputStream("E:/students.xls");
            wb.write(fout);
            fout.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }

    }


}
