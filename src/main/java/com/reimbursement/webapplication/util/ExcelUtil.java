package com.reimbursement.webapplication.util;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

public class ExcelUtil {

    /**
     * 轉換數據格式
     * @param cell 單元格
     * @return Object
     */
    public static Object getValue(Cell cell) {
        Object cellValue ;
        //判断数据的类型
        if(cell == null){
            return "";
        }
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC: //数字
                // 判断是否是日期类型
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
                    Date dateCellValue = cell.getDateCellValue();
                    cellValue = formatter.format(dateCellValue);
                } else {
                    cellValue = cell.getNumericCellValue();
                }
                break;
            case Cell.CELL_TYPE_STRING: //字符串
                cellValue = cell.getStringCellValue();
                break;
            case Cell.CELL_TYPE_BOOLEAN: //Boolean
                cellValue = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_FORMULA: //公式
               try {
                   cellValue = Double.valueOf(cell.getCellFormula());
               }catch (Exception e){
                   cellValue = Double.valueOf(cell.getNumericCellValue());
               }
               System.out.println(cellValue);
                break;
            case Cell.CELL_TYPE_BLANK: //空值
                cellValue = "";
                break;
            case Cell.CELL_TYPE_ERROR: //故障
                cellValue = "非法字符";
                break;
            default:
                cellValue = "未知类型";
                break;
        }

        return cellValue;
    }


    /**
     * 根据excel里面的内容读取客户信息
     * @param is 输入流
     * @param suffix excel是2003还是2007版本
     * @return
     * @throws IOException
     */
    public static Workbook getWorkbook(InputStream is, String suffix){
        /** 根据版本选择创建Workbook的方式 */
        Workbook wb = null;
        try{
            if(suffix.equals("xls")){
                //当excel是2003时
                wb = new HSSFWorkbook(is);
            }else{
                //当excel是2007时
                wb = new XSSFWorkbook(is);
            }
        }
        catch (IOException e)  {
            e.printStackTrace();
        }
        return wb;
    }


    /**
     * 通過標題列表返回實體對象
     * @param cls class反射
     * @param row 列
     * @param titles 標題列表
     * @return Object 返回實體對象
     */
    public static Object getDataList(Class<?> cls, Row row, List<String> titles) throws ClassNotFoundException, IllegalAccessException, InstantiationException {
        Object bean = Class.forName(cls.getName()).newInstance();

        for(int i = 1 ;i<row.getLastCellNum(); i++){

            if(row.getCell(i) == null || "".equals(row.getCell(i).toString())){
                break;
            }

            try {
                if(titles.get(i-1).equals("")){
                    continue;
                }
                Field field = bean.getClass().getDeclaredField(titles.get(i-1));
                field.setAccessible(true);
                field.set(bean, ExcelUtil.getValue(row.getCell(i)));
            } catch (Exception e) {
                System.out.println("没有对应的方法：" + e);
                break;

            }
        }

        return bean;

    }
}