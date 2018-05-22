package com.wdocode.document.excel.model;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.dom4j.Element;

/**
 * 
 * 模板中要替换的数据
 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
 * @since 2016年4月5日 下午5:34:44
 *
 */
public class ExcelReplaceHolder {
	/**导出描述中模板文件配置-模板中需要替换的行 -需要替换的数据节点描述 */
	public static final String replaceData  = "replaceData";
	/**导出描述中模板文件配置-模板中需要替换的行 -要替换内容占位符 */
	public static final String placeholder_key = "placeholder";
	/**导出描述中模板文件配置-模板中需要替换的行 -替换内容对应的数据源map中的key */
	public static final String replaceData_key = "key";
	
	/**占位符 */
	private String placeholder;
	/**数据源map中的key */
	private String key;
	
	
	public ExcelReplaceHolder() {
	}

	public ExcelReplaceHolder(String placeholder,String key) {
		setKey(key);
		setPlaceholder(placeholder);
	}
	
	public String getPlaceholder() {
		return placeholder;
	}
	public void setPlaceholder(String placeholder) {
		this.placeholder = placeholder;
	}
	public String getKey() {
		return key;
	}
	public void setKey(String key) {
		this.key = key;
	}
	
	/**
	 * 在xml中读取模板中要替换的行数据
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param elements
	 * @return
	 */
	public static List<ExcelReplaceHolder> readFromElement(List<Element> elements) {
		if(elements == null)
			return null;
		List<ExcelReplaceHolder> list = new ArrayList<ExcelReplaceHolder>();
		for(Element e:elements)
		list.add(readFromElement(e));
		return list;
	}
	
	/**
	 * 在xml中读取模板中要替换的行数据
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param element
	 * @return
	 */
	public static ExcelReplaceHolder readFromElement(Element element) {
		if(element == null || StringUtils.isBlank(element.attributeValue(placeholder_key)))
			return null;
		return new ExcelReplaceHolder(element.attributeValue(placeholder_key),
				element.attributeValue(replaceData_key,""));
	}
	
}
