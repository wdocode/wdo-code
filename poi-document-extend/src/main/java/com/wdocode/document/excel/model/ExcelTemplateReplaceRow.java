package com.wdocode.document.excel.model;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.dom4j.Element;

/**
 * EXCEL 模板文件
 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
 * @since 2016年4月5日 下午5:34:44
 *
 */
public final class ExcelTemplateReplaceRow{
//	<mapper id='exportload_limit_notice' file_name="载重平衡中心飞机临时减载">
//  <!-- 模板路径只适用于导出时，模板文件放置在相对与本xml文件的下级templates目录中的 -->
//  	<template path = "flight_load_limit_notice.xls" >
//	    	<replace_rows>
//		    	<row index='2'>
//		    		<replaceData placeholder="NEW_DATE" key="$TIME_YYYYMMDD_CN"/>
//		    		<replaceData placeholder="NEW_VERSION" key="NEW_VERSION"/>
//		    	</row>
//	    	</replace_rows>
//  	</template>
//  </mapper>
	
	/**导出描述中模板文件配置-模板中需要替换的行父节点 */
	public static final String replace_rows = "replace_rows";
	/**导出描述中模板文件配置-模板中需要替换的行 内容描述节点 */
	public static final String row_key = "row";
	/**导出描述中模板文件配置-模板中需要替换的行 -行号 */
	public static final String row_index = "index";
	
	
	/**要替换文本所在行号 */
	private int rowIndex;
	
	private List<ExcelReplaceHolder> replaceHolder;
	
	public ExcelTemplateReplaceRow() {
	}
	
	public ExcelTemplateReplaceRow(int rowIndex) {
		setRowIndex(rowIndex);
	}

	public int getRowIndex() {
		return rowIndex;
	}

	public void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}


	public List<ExcelReplaceHolder> getReplaceHolder() {
		return replaceHolder;
	}

	public void setReplaceHolder(List<ExcelReplaceHolder> replaceHolder) {
		this.replaceHolder = replaceHolder;
	}

	/**
	 * 在xml中读取模板中要替换的行数据
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param elements
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public static List<ExcelTemplateReplaceRow> readFromElement(List<Element> rows) {
		if(rows == null)
			return null;
		List<ExcelTemplateReplaceRow> list = new ArrayList<ExcelTemplateReplaceRow>();
		for(Element r: rows){
			String index = r.attributeValue(row_index);
			if(!StringUtils.isNumeric(index))
				continue;
			ExcelTemplateReplaceRow row = new ExcelTemplateReplaceRow(Integer.parseInt(index));
			row.setReplaceHolder(ExcelReplaceHolder.readFromElement(r.elements(ExcelReplaceHolder.replaceData)));
			list.add(row);
		}
		return list;
	}

	/**
	 * 在xml中读取模板中要替换的行数据
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param e
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public static List<ExcelTemplateReplaceRow> readFromElement(Element e) {
		if(e == null)
			return null;
		return readFromElement(e.elements(row_key));
	}

	
}