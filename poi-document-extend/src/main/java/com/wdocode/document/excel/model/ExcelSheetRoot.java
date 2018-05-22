package com.wdocode.document.excel.model;

import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * excel 尾部模板数据
 * 注意模板数据全部当作文本处理
 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
 * @since 2017年9月22日 14:05:45
 *
 */
public final class ExcelSheetRoot{
	
	/**单元格数据 */
	private List<String []> cellDate;
	/**单元格样式 */
	private List<CellStyle []> cellStyles;
	/**合并单元格坐标 */
	private List<CellRangeAddress> ranges;
	
	
	public ExcelSheetRoot() {
	}
	
	public ExcelSheetRoot(List<String []> cellDate,
			List<CellStyle []> cellStyles,List<CellRangeAddress> ranges) {
		this.cellDate =cellDate; 
		this.cellStyles = cellStyles;
		this.ranges = ranges;
	}
	
	
	
	public List<String[]> getCellDate() {
		return cellDate;
	}
	public void setCellDate(List<String[]> cellDate) {
		this.cellDate = cellDate;
	}
	public List<CellStyle[]> getCellStyles() {
		return cellStyles;
	}
	public void setCellStyles(List<CellStyle[]> cellStyles) {
		this.cellStyles = cellStyles;
	}
	public List<CellRangeAddress> getRanges() {
		return ranges;
	}
	public void setRanges(List<CellRangeAddress> ranges) {
		this.ranges = ranges;
	}
	
	
	
	
}