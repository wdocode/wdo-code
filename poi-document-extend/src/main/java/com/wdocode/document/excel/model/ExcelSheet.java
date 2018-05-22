package com.wdocode.document.excel.model;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.dom4j.Element;

/**
 * excel映射文件 表头映射关系
 * @author zhangzx
 * 
 * @since 2016年4月5日 下午5:34:44
 */
public final class ExcelSheet{

	/**导出文件中-对应单个sheet表描述根节点 */
	public static final String sheet_mapper = "sheet_mapper";
	/**导出文件中-对应单个sheet表描述根节点 -属性-表名 */
	public static final String sheetName_key = "sheet_name";
	/**导出文件中-对应单个sheet表描述根节点 -属性-对应数据源map中的key*/
	public static final String map_key = "map_key";
	/**导出文件中-对应单个sheet表每行描述根节点*/
	public static final String rows = "rows";
	public static final String attr_rowHeight = "rowHeight";
	public static final String attr_cellWidth = "cellWidth";
	
	/**表格名 */
	private String sheetName;
	/**数据源map中的key */
	private String mapKey;
	/**sheet对应数据源map中的行数据key */
	private String rowMapKey;
	
	/**行高 */
	private int rowHeight; 
	private int cellWidth;
	
	private List<ExcelColumn> columns;
	/**excel模板文件 */
	private ExcelTemplate template;
	
	public ExcelSheet() {
	}
	
	public ExcelSheet(String sheetName,String mapKey) {
		setSheetName(sheetName); 
		setMapKey(mapKey);
	}

	public ExcelSheet(String sheetName,String mapKey,String rowMapKey) {
		setSheetName(sheetName); 
		setMapKey(mapKey);
		this.rowMapKey = rowMapKey;
	}
	
	public String getSheetName() {
		return sheetName;
	}
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName==null?null:sheetName.trim();
	}
	public String getMapKey() {
		return mapKey;
	}
	public void setMapKey(String mapKey) {
		this.mapKey = mapKey;
	}
	public String getRowMapKey() {
		return rowMapKey;
	}
	public void setRowMapKey(String rowMapKey) {
		this.rowMapKey = rowMapKey;
	}

	public List<ExcelColumn> getColumns() {
		return columns;
	}

	public void setColumns(List<ExcelColumn> columns) {
		this.columns = columns;
	}

	
	public int getCellWidth() {
		return cellWidth;
	}

	public void setCellWidth(int cellWidth) {
		this.cellWidth = cellWidth;
	}

	public int getRowHeight() {
		return rowHeight;
	}

	public void setRowHeight(int rowHeight) {
		this.rowHeight = rowHeight;
	}
	
	

	public ExcelTemplate getTemplate() {
		return template;
	}

	public void setTemplate(ExcelTemplate template) {
		this.template = template;
	}

	/**
	 * 在xml中读取sheet表格映射配置
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param sheetNode
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public static ExcelSheet readFromElement(Element sheetNode) {
		if(sheetNode == null 
				|| StringUtils.isBlank(sheetNode.attributeValue(sheetName_key))
				|| StringUtils.isBlank(sheetNode.attributeValue(map_key)))
			return null;
		ExcelSheet sheet = new ExcelSheet(
				sheetNode.attributeValue(sheetName_key),sheetNode.attributeValue(map_key));
		
		ExcelTemplate temp = ExcelTemplate.readFromElement(sheetNode.element(ExcelTemplate.template));
		sheet.setTemplate(temp);
		Element row = sheetNode.element(rows);
		if(row != null){
			String h =  row.attributeValue(attr_rowHeight,"16");
			sheet.setRowHeight(Integer.parseInt(h));
			String w =  row.attributeValue(attr_cellWidth);
			if(StringUtils.isNumeric(w))
				sheet.setCellWidth(Integer.parseInt(w));
			sheet.setRowMapKey(row.attributeValue(map_key));
			sheet.setColumns(ExcelColumn.readFromElement(row.elements(ExcelColumn.sheetcolumn_key)));
		}
		
		return sheet;
	}
	
	/**
	 * 在xml中读取sheet表格映射配置
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param elements
	 * @return
	 */
	public static List<ExcelSheet> readFromElement(List<Element> elements) {
		if(elements == null)
			return null;
		List<ExcelSheet> list = new ArrayList<ExcelSheet>();
		for(Element e:elements){
			list.add(readFromElement(e));
		}
		
		return list;
	}
	
	
}