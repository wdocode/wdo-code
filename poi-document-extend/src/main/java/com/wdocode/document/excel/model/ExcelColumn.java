package com.wdocode.document.excel.model;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.dom4j.Element;

/**
 * excel映射文件 表头映射关系
 * @author zhangzx
 *
 */
public final class ExcelColumn{
	/**导出文件中-对应单个sheet表每行-各列描述*/
	public static final String sheetcolumn_key = "column";
	/**导出文件中-对应单个sheet表每行中各列标题 */
	public static final String sheettitle_key = "title";
	/**导出文件中-对应单个sheet表每行中各列标题对应的数据源map中的key */
	public static final String cl_name_key = "cl_name";
	/**导出文件中-对应单个sheet表每行中各列时间格式化字符(数据源类型为java.util.Date时生效)*/
	public static final String formate_key = "dataformate";
	/**DATA TYPE [NUMBER,DATE] */
	public static final String data_type_key = "data_type";
	/** 单列中要映射的字典项*/
	public static final String dataMapper_key = "data_mapper";
	/** 单列中要映射的字典项-数据源中对应列的数据*/
	public static final String dataMapperKey_name = "key";
	/** 单列中要映射的字典项-转换的目标文本*/
	public static final String dataMapperKey_value = "value";
	public static final String attr_width = "width";
	public static final String attr_horizontalAlign = "horizontalAlign";
	public static final String attr_verticalAlign = "verticalAlign";
	public static final String attr_mergeRow = "mergeRow";
	public static final String attr_mergeRelyLeft = "mergeRelyLeft";
	
	
	private String tilte;
	private String column;
	private Map<String, String> dataMapper = new HashMap<String, String>();
	private String dataformat;
	private int width;
	/**水平对齐方式 */
	private String horizontalAlign;
	/**垂直对其方式 */
	private String verticalAlign;
	/**是否需要合并行 */
	private boolean mergeRow;
	/**合并行时 是否依赖左侧单元格 */
	private boolean mergeRelyLeft;
	
	private SimpleDateFormat format_date = null;
	private DecimalFormat format_number = null;
	private String dataType ;
	
	public ExcelColumn() {
	}
	
	public ExcelColumn(String title,String column) {
		this.tilte = title;
		this.column = column;
	}
	
	public String getTilte() {
		return tilte;
	}
	public void setTilte(String tilte) {
		this.tilte = tilte;
	}
	public String getColumn() {
		return column;
	}
	public void setColumn(String column) {
		this.column = column;
	}
	public Map<String, String> getDataMapper() {
		return dataMapper;
	}
	public void setDataMapper(Map<String, String> dataMapper) {
		this.dataMapper = dataMapper;
	}
	public String getDataformat() {
		return dataformat;
	}
	public void setDataformat(String dataformat) {
		this.dataformat = dataformat;
	}
	
	
	
	
	public int getWidth() {
		return width;
	}

	public void setWidth(int width) {
		this.width = width;
	}
	
	public void setWidth(String width) {
		if (StringUtils.isNotBlank(width) && StringUtils.isNumeric(width)) {
			this.width = Integer.parseInt(width);
		}
	}
	
	

	public String getHorizontalAlign() {
		return horizontalAlign;
	}

	public void setHorizontalAlign(String horizontalAlign) {
		this.horizontalAlign = horizontalAlign;
	}

	public String getVerticalAlign() {
		return verticalAlign;
	}

	public void setVerticalAlign(String verticalAlign) {
		this.verticalAlign = verticalAlign;
	}
	

	public String getDataType() {
		return dataType;
	}

	public void setDataType(String dataType) {
		this.dataType = dataType;
	}
	

	public boolean isMergeRow() {
		return mergeRow;
	}

	public void setMergeRow(boolean mergeRow) {
		this.mergeRow = mergeRow;
	}

	/**
	 * 在xml中读取sheet表格映射配置
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param elements
	 * @return
	 */
	public static ExcelColumn readFromElement(Element element) {
		if(element == null)
			return null;
		String title = element.attributeValue(sheettitle_key);
		String colname = element.attributeValue(cl_name_key);
		String formate_str = element.attributeValue(formate_key); 
		String width = element.attributeValue(attr_width);
		String data_type = element.attributeValue(data_type_key);
		if(StringUtils.isBlank(title) || StringUtils.isBlank(colname))
			return null;
		ExcelColumn col = new ExcelColumn(title,colname);
		col.setWidth(width);
		col.setHorizontalAlign(element.attributeValue(attr_horizontalAlign));
		col.setVerticalAlign(element.attributeValue(attr_verticalAlign));
		col.setDataMapper(getcolumndataMapper(element));
		col.setDataType(data_type);
		col.setMergeRow("true".equalsIgnoreCase(element.attributeValue(attr_mergeRow)));
		col.setMergeRelyLeft("true".equalsIgnoreCase(element.attributeValue(attr_mergeRelyLeft)));
		if(StringUtils.isNotBlank(formate_str))
			col.setDataformat(formate_str);
		return col;
	}
	
	/**
	 * 在xml中读取sheet表格映射配置
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param elements
	 * @return
	 */
	public static List<ExcelColumn> readFromElement(List<Element> elements) {
		if(elements == null)
			return null;
		List<ExcelColumn> list = new ArrayList<ExcelColumn>();
		for(Element e:elements){
			list.add(readFromElement(e));
		}
		return list;
	}
	
	

	public SimpleDateFormat getFormat_date() {
		if(format_date != null){
			return format_date;
		}
//		NUMBER,DATE
		if("DATE".equalsIgnoreCase(dataType) && StringUtils.isNotBlank(dataformat)){
			format_date = new SimpleDateFormat(dataformat);
		}
		return format_date;
	}

	public void setFormat_date(SimpleDateFormat format_date) {
		this.format_date = format_date;
	}

	public DecimalFormat getFormat_number() {
		if(format_number != null){
			return format_number;
		}
		
		if("NUMBER".equalsIgnoreCase(dataType) && StringUtils.isNotBlank(dataformat)){
			format_number = new DecimalFormat(dataformat);
		}
		
		return format_number;
	}

	public void setFormat_number(DecimalFormat format_number) {
		this.format_number = format_number;
	}
	

	public boolean isMergeRelyLeft() {
		return mergeRelyLeft;
	}

	public void setMergeRelyLeft(boolean mergeRelyLeft) {
		this.mergeRelyLeft = mergeRelyLeft;
	}

	/**
	 * 获取没列数据描述映射（字典值映射）
	 * @param c
	 * @return
	 */
	@SuppressWarnings("unchecked")
	private static Map<String, String> getcolumndataMapper(Element c) {
		if(c == null)
			return null;
		List<Element> datamapper = c.elements(dataMapper_key);
		if(datamapper == null || datamapper.isEmpty())
			return null;
		Map<String, String> result = new HashMap<String, String>();
		for(Element e:datamapper){
			String val = e.attributeValue(dataMapperKey_value);
			String key = e.attributeValue(dataMapperKey_name);
			result.put(key == null?null:key.trim().toUpperCase(),
					StringUtils.isBlank(val)? e.getTextTrim():val.trim());
		}
		return result;
	}
	
}