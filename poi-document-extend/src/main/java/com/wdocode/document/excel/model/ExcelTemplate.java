package com.wdocode.document.excel.model;

import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.dom4j.Element;

/**
 * EXCEL 模板文件
 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
 * @since 2016年4月5日 下午5:34:44
 *
 */
public final class ExcelTemplate{
	/**导出描述中模板文件配置 */
	public static final String template = "template";
	/**导出描述中模板文件配置-文件路径 */
	public static final String temp_titleIndex = "titleIndex";
	/**导出文件中-对应单个sheet表每行描述根节点-属性-导出后是否删除标题行*/
	public static final String removeTitle_key = "removeTitle";
	public static final String rootIndex_key = "rootIndex";
	
	/**模板中标题所在行号 */
	private int titleIndex;
	
	/**是否移除标题行 */
	private boolean removeTitle;
	
	/**表格落款行 */
	private int[] rootIndex;
	
	
	private List<ExcelTemplateReplaceRow> replacerows;
	
	public ExcelTemplate() {
	}
	
	
	public ExcelTemplate(int titleIndex,boolean removeTitle) {
		setTitleIndex(titleIndex);
		setRemoveTitle(removeTitle);
	}

	public List<ExcelTemplateReplaceRow> getReplacerows() {
		return replacerows;
	}

	public void setReplacerows(List<ExcelTemplateReplaceRow> replacerows) {
		this.replacerows = replacerows;
	}

	
	public int getTitleIndex() {
		return titleIndex;
	}

	public void setTitleIndex(int titleIndex) {
		this.titleIndex = titleIndex;
	}

	public boolean isRemoveTitle() {
		return removeTitle;
	}

	public void setRemoveTitle(boolean removeTitle) {
		this.removeTitle = removeTitle;
	}

	
	
	public int[] getRootIndex() {
		return rootIndex;
	}


	public void setRootIndex(int[] rootIndex) {
		this.rootIndex = rootIndex;
	}


	/**
	 * 
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param element
	 * @return
	 */
	public static ExcelTemplate readFromElement(Element element) {
		if(element == null)
			return null;
		ExcelTemplate tem = new ExcelTemplate();
		String title_index = element.attributeValue(temp_titleIndex,"0");
		if(StringUtils.isNumeric(title_index))
			tem.setTitleIndex(Integer.parseInt(title_index));
		String rootIndex =element.attributeValue(rootIndex_key);
		if(rootIndex != null){
			String [] roots = rootIndex.split(",");
			int rootarray[] = new int[roots.length];
			for(int i = 0,l = roots.length;i<l;i++){
				if(StringUtils.isNumeric(roots[i]))
				rootarray[i] = Integer.parseInt(roots[i]);
			}
			tem.setRootIndex(rootarray);
		}
		tem.setRemoveTitle("true".equalsIgnoreCase(element.attributeValue(removeTitle_key)));
		tem.setReplacerows(ExcelTemplateReplaceRow.readFromElement(element.element("replace_rows")));
		return tem;
	}


	
}