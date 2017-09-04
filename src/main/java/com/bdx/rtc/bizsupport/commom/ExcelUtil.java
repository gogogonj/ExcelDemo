package com.bdx.rtc.bizsupport.commom;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 导出Excel文档工具类
 */
public class ExcelUtil {

    public static void main(String[] args) {
        String path = "D:/日常记账.xls";
        String sheetName = "2017-09";
        String[] columnNames = {"日期", "金额", "事由","支付方式","对应卡片","类别1","类别2"};
        List<String[]> dataList = new ArrayList<String[]>();
        String[] b = {"2017-09-04", "15", "午饭-万福面","微信","招商银行储蓄卡","食",""};
        String[] c = {"2017-09-04", "8", "饼干+酸奶","微信","招商银行储蓄卡","食",""};
        dataList.add(b);
        dataList.add(c);

        excelAdd(path, sheetName, columnNames, dataList);
    }

    public static void excelAdd(String path, String sheetName, String[] columnNames, List<String[]> list) {
        Workbook wb = null;

        File file = new File(path);
        if (!file.exists()) {
            //文件不存在，创建文件、sheet、第一行
            wb = saveAsExcel(path, sheetName, columnNames);
        } else {
            //文件存在，获取Workbook
            try {
                wb = new HSSFWorkbook(new FileInputStream(path));
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        //获取sheet
        Sheet sheet = wb.getSheet(sheetName);
        if (null == sheet) {
            //sheet不存在
            sheet = createSheet(wb, sheetName, columnNames);
        }
        //追加数据
        excelDataAdd(wb, sheet, path, list);
    }

    /**
     * 向excel追加数据
     *
     * @param path
     * @param list
     * @throws Exception
     */
    public static void excelDataAdd(Workbook wb, Sheet sheet, String path, List<String[]> list) {
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(path);
            for (String[] contents : list) {
                Row row = sheet.createRow((short) (sheet.getLastRowNum() + 1)); //在现有行号后追加数据
                for (int i = 0; i < contents.length; i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(contents[i]);
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            try {
                out.flush();
                wb.write(out);
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 新建一个excel，包括sheet和标题
     *
     * @param path
     * @param sheetName
     * @param columnNames
     * @return
     */
    public static Workbook saveAsExcel(String path, String sheetName, String[] columnNames) {
        Workbook wb = null;
        try {
            FileOutputStream os = new FileOutputStream(path);
            wb = createWorkBook(sheetName, columnNames);
            wb.write(os);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }

    /**
     * 创建sheet
     *
     * @param wb
     * @param sheetName
     * @param columnNames
     * @return
     */
    public static Sheet createSheet(Workbook wb, String sheetName, String[] columnNames) {
        Sheet sheet = wb.createSheet(sheetName);
        // 创建第一行: 列名行
        Row row1 = sheet.createRow((short) 0); //行数row序号也是从0开始的
        // 手动设置列宽。第一个参数表示要为第几列设，第二个参数表示列的宽度，n为列高的像素数。
        for (int i = 0; i < columnNames.length; i++) {
            sheet.setColumnWidth((short) i, (short) (35.7 * 150));
            Cell cell = row1.createCell(i);
            cell.setCellValue(columnNames[i]);
        }
        return sheet;
    }

    /**
     * 创建excel文档
     *
     * @param sheetName   sheet名称
     * @param columnNames excel的列名
     */
    public static Workbook createWorkBook(String sheetName, String[] columnNames) {
        Workbook wb = new HSSFWorkbook();
        // 创建第一个sheet（页），并命名
        Sheet sheet = wb.createSheet(sheetName);
        // 手动设置列宽。第一个参数表示要为第几列设，第二个参数表示列的宽度，n为列高的像素数。
        for (int i = 0; i < columnNames.length; i++) {
            sheet.setColumnWidth((short) i, (short) (35.7 * 150));
        }

        // 创建两种单元格格式
        CellStyle cs = wb.createCellStyle();
        CellStyle cs2 = wb.createCellStyle();

        // 创建两种字体
        Font f = wb.createFont();
        Font f2 = wb.createFont();

        // 创建第一种字体样式（用于列名）
        f.setFontHeightInPoints((short) 10);
        f.setColor(IndexedColors.BLACK.getIndex());
        f.setBoldweight(Font.BOLDWEIGHT_BOLD);

        // 创建第二种字体样式（用于值）
        f2.setFontHeightInPoints((short) 10);
        f2.setColor(IndexedColors.BLACK.getIndex());

        // 设置第一种单元格的样式（用于列名）
        cs.setFont(f);
        cs.setBorderLeft(CellStyle.BORDER_THIN);
        cs.setBorderRight(CellStyle.BORDER_THIN);
        cs.setBorderTop(CellStyle.BORDER_THIN);
        cs.setBorderBottom(CellStyle.BORDER_THIN);
        cs.setAlignment(CellStyle.ALIGN_CENTER);

        // 设置第二种单元格的样式（用于值）
        cs2.setFont(f2);
        cs2.setBorderLeft(CellStyle.BORDER_THIN);
        cs2.setBorderRight(CellStyle.BORDER_THIN);
        cs2.setBorderTop(CellStyle.BORDER_THIN);
        cs2.setBorderBottom(CellStyle.BORDER_THIN);
        cs2.setAlignment(CellStyle.ALIGN_CENTER);

        // 创建第一行: 列名行
        Row row1 = sheet.createRow((short) 0); //行数row序号也是从0开始的
        for (int i = 0; i < columnNames.length; i++) {
            Cell cell = row1.createCell(i);
            cell.setCellValue(columnNames[i]);
            cell.setCellStyle(cs);
        }
        return wb;
    }

}
