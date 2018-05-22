package com.wdocode.document.excel;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

/**
 * excel文件处理公用方法
 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
 * @since 2016年4月5日 下午5:34:44
 *
 */
public class EXCELDocumentCommon {
	private static final String def_dateFormate = "yyyy-MM-dd HH:mm:ss:SSS";
	private static final String def_DecimalFormat = "0.00";
	public static final int SHEETMAXROWS_2003 = 65535;
	public static final int SHEETMAXROWS_2007 = 1048575;
	public static final String EXCEL_SUFFIX_XLSX = ".xlsx";
	public static final String EXCEL_SUFFIX_XLS = ".xls";
	public static final String EXCEL_SUFFIX_CSV = ".csv";
	
	/**
	 * 获取每行数据
	 *
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param row
	 * @return
	 */
	public static Object[] getRowDate(Row row) {
		if (row == null)
			return null;
		int tempRowSize = row.getLastCellNum();
		Object[] values = new Object[tempRowSize];
//		Arrays.fill(values, null);
		for (int columnIndex = 0; columnIndex < tempRowSize; columnIndex++) {
			values[columnIndex] = getCellData(row.getCell(columnIndex));
		}
		return values;
	}
	
	/**
	 * 获取单行数据
	 *
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param row
	 * @return
	 */
	public static String[] getRowDateToString(Row row) {
		return getRowDateToString(row,null);
	}
	
	/**
	 * 获取行数据
	 *
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param row
	 * @param defaultValue
	 * @return
	 */
	public static String[] getRowDateToString(Row row,String defaultValue) {
		if (row == null)
			return null;
		int maxSize = row.getLastCellNum();
		String[] values = new String[maxSize];
//		数组设置默认值
		Arrays.fill(values, defaultValue);
		for (int columnIndex = 0; columnIndex < maxSize; columnIndex++) {
			Cell cell = row.getCell(columnIndex);
			values[columnIndex] = getCellDataToString(cell,null);
		}
		return values;
	}
	
	/**
	 * 获取单元格中的数据
	 * 
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param cell
	 * @return
	 */
	public static String getCellDataToString(Cell cell,String formate) {
		Object value = getCellData(cell);
		if(value == null)
			return null;
		return formateData(value,formate);
	}

	
	
	/**
	 * 将Object对象格式化为字符串
	 * 
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param value
	 * @param formate
	 * @return
	 */
	public static String formateData(Object value, final String formate) {
		String f = "";
		if(value instanceof Date){
			f = StringUtils.isNotBlank(formate)?formate:def_dateFormate;
			return value == null?null:new SimpleDateFormat(f).format(value);
		}else if(value instanceof Number){
			f = StringUtils.isNotBlank(formate)?formate:def_DecimalFormat;
			return value == null?"":new DecimalFormat(f).format(value);
		}else{
			return value == null?"":String.valueOf(value);
		}
	}

	/**
	 * 获取单元格中的数据
	 * 
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param cell
	 * @return
	 */
	@SuppressWarnings("deprecation")
	public static Object getCellData(Cell cell) {
		if (cell == null)
			return null;
		// 注意：一定要设成这个，否则可能会出现乱码
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			// 字符型可直接返回
			return cell.getStringCellValue();
		case Cell.CELL_TYPE_NUMERIC:
			// 数值性判断是否为时间格式
			if (DateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue();
			} else {
				// 普通数字则返回double型数据
				return cell.getNumericCellValue();
				// value = new DecimalFormat("0").format(data); //
			}
		case Cell.CELL_TYPE_FORMULA:
			// 导入时如果为公式生成的数据则无值
			try {
				return cell.getStringCellValue();
			} catch (Exception e) {
				return String.valueOf(cell.getNumericCellValue());
			}

		case Cell.CELL_TYPE_BLANK:
			// 空值返回null
			return null;
		case Cell.CELL_TYPE_BOOLEAN:
			// 布尔型
			return cell.getBooleanCellValue();
		case Cell.CELL_TYPE_ERROR:
			return null;
		default:
			return null;
		}
	}
	
}
