package com.wdocode.document.excel.model;

import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.dom4j.Element;

/**
 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
 * @since 2016年4月5日 下午5:34:44
 *
 */
public class ExcelDataMapper {
	/**数据映射描述id xml中要根据该值判断要导出的excle文件内容*/
	public static final String attr_dataMapperId = "id";
	/**导出显示的文件名 */
	public static final String attr_fileName = "file_name";
	public static final String attr_template_fileName = "template_fileName";
	public static final String attr_DEF_DATAFORMATE = "def_dataformate";
	/**模板文件绝对路径 */
	private String templateFilePath;
	private String xmlBasePath;
	
	/**导出要生成的文件名,不包含后缀 */
	private String fileName;
	
	/**导出要生成的文件类型（后缀[如.xml]） */
	private String fileSuffix;
	/**MAPPER id */
	private String id;
	
	/**excel中sheet列表 */
	private List<ExcelSheet> sheets;
	/**模板文件名 */
	private String template_fileName;
	/**
	 * 日期类型默认格式
	 */
	private String defDataformate;
	
	public ExcelDataMapper() {
	}
	
	/**导出要生成的文件名,不包含后缀 */
	public String getFileName() {
		return fileName;
	}
	
	/**
	 * 获取完整文件名
	 * @return
	 */
	public String getFullFileName() {
		return fileName + fileSuffix;
	}
	
	public void setFileName(String fileName) {
		this.fileName = fileName ==null?null:fileName.trim();
	}
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}
	
	public List<ExcelSheet> getSheets() {
		return sheets;
	}
	public void setSheets(List<ExcelSheet> sheet) {
		this.sheets = sheet;
	}
	
	public String getTemplate_fileName() {
		return template_fileName;
	}

	public void setTemplate_fileName(String template_fileName) {
		this.template_fileName = template_fileName;
	}
	

	public String getXmlBasePath() {
		return xmlBasePath;
	}

	public void setXmlBasePath(String xmlBasePath) {
		this.xmlBasePath = xmlBasePath;
	}
	
	/**导出要生成的文件类型（后缀[如.xml]） */
	public String getFileSuffix() {
		return fileSuffix;
	}

	public void setFileSuffix(String fileSuffix) {
		this.fileSuffix = fileSuffix;
	}

	/**
	 * 在配置文件xml片段中读取映射对象
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param e
	 * @return
	 */
	public static ExcelDataMapper readFromElement(Element e) {
		return readFromElement(null, e);
	}

	/**
	 * 在配置文件xml片段中读取映射对象
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param e
	 * @return
	 */
	public static ExcelDataMapper readFromElement(String xmlBasePath, Element e) {
		if(StringUtils.isBlank(e.attributeValue(attr_dataMapperId)))
			return null;
		ExcelDataMapper mapper = new ExcelDataMapper();
		mapper.setXmlBasePath(xmlBasePath);
		mapper.setId(e.attributeValue(attr_dataMapperId));
		String fileName = e.attributeValue(attr_fileName, "download.xls");
		mapper.setFileName(fileName);
		if(fileName.indexOf(".")>0){
			mapper.setFileName(fileName.substring(0, fileName.lastIndexOf(".")));
			mapper.setFileSuffix(fileName.substring(fileName.lastIndexOf(".")));
		}
		mapper.setTemplate_fileName(e.attributeValue(attr_template_fileName));
		mapper.setDefDataformate(e.attributeValue(attr_DEF_DATAFORMATE));
		@SuppressWarnings("unchecked")
		List<Element> sheet = e.elements(ExcelSheet.sheet_mapper);
		mapper.setSheets(ExcelSheet.readFromElement(sheet));
		return mapper;
	}
	
	/**
	 * 根据数据源map中key获取sheet描述对象
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param key
	 * @return
	 */
	public ExcelSheet getSheetByMapKey(String key) {
		if(sheets == null || StringUtils.isBlank(key))
			return null;
		for(ExcelSheet st :sheets){
			if(key.equals(st.getMapKey()))
				return st;
		}
		return null;
	}
	
	/**
	 * 获取模板绝对路径
	 * @return
	 */
	public String getTemplateFilePath(){
		if(StringUtils.isBlank(template_fileName))
			return null;
		if(templateFilePath != null){
			return templateFilePath;
		}
		// 获取当前实例 部署源码的根目录
		String webroot =  System.getProperty("webapp.root");
		if(webroot == null){
			webroot = xmlBasePath == null?"":xmlBasePath;
		}
		webroot = webroot.replace("\\", "/");
		
		StringBuffer sb = new StringBuffer(webroot);
		sb.append(webroot.endsWith("/")?"":"/")
		.append("resource/templates/")
		.append(template_fileName);
		templateFilePath = sb.toString();
		return templateFilePath;
	}

	public String getDefDataformate() {
		return defDataformate;
	}

	public void setDefDataformate(String defDataformate) {
		this.defDataformate = defDataformate;
	}

	
}

