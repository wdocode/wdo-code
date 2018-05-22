package com.wdocode.document.excel.model;

import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 工作簿中sheet表中行号计数机
 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
 * @since 2016年4月5日 下午5:34:44
 *
 */
public class SheetRowDataCounter {
	/**标题顺序 */
	private String [] title;
	/**需要合并行的标题 */
	private Set<String> mergeTitles;
	
	/**合并时依赖左侧列的标题 */
	private Set<String> mergeRelyLeftTitle;
	/** 合并行数据*/
	private Map<String, Object> mergeTitleDate;
	/**行索引位置 */
	private volatile int index;
	/**数据所在的首行 */
	private int titleIndex;
	
	/**单元格样式 */
	private CellStyle[] cellstyles;
	/**最大行号 */
	private int maxRowSize = 65535;
	/**落款行号 */
	private int rootIndexs [];
	
	private ExcelSheetRoot excelSheetRoot;
	
	/**数据字典映射内容 */
	private Map<String, Map<String,String>> dataMapper;
	/**表格名称 */
	private String sheetName;
	
	/**
	 * 构造方法
	 * @param index
	 * @param titles
	 */
	public SheetRowDataCounter(int index, String [] titles){
		setIndex(index);
		setTitleIndex(index);
		setTitle(titles);
	}
	
	/**
	 * 构造方法
	 * @param index 首行行号
	 * @param maxRowSize 最大行数
	 * @param titles 标题数组
	 * @param cellstyles 单元格样式
	 */
	public SheetRowDataCounter(int index,int maxRowSize, 
			String [] titles,CellStyle[] cellstyles){
		setIndex(index);
		setMaxRowSize(maxRowSize);
		setTitleIndex(index);
		setTitle(titles);
		setCellstyles(cellstyles);
	}
	
	public SheetRowDataCounter(int index){
		setIndex(index);
	}
	public SheetRowDataCounter(){
		
	}
	
	public String[] getTitle() {
		return title;
	}
	public void setTitle(String[] title) {
		this.title = title;
	}
	public int getIndex() {
		return index;
	}
	public void setIndex(int index) {
		this.index = index;
	}
	
	/**
	 * 获取下一行行号
	 * 
	 * @author zhang zixiao
	 * @return
	 */
	public int getNextIndex() {
		int nex = index -getCurrentSheetNO()*maxRowSize;
//		if(nex <=0){
//			System.out.println(nex);
//		}
		return nex+1;
	}
	
	/**
	 * 获取上一行行号
	 * 
	 * @author zhang zixiao
	 * @return
	 */
	public int getLastIndex() {
		return index==0?0:index-1;
	}

	public CellStyle[] getCellstyles() {
		return cellstyles;
	}

	public void setCellstyles(CellStyle[] cellstyles) {
		this.cellstyles = cellstyles;
	}

	public int getMaxRowSize() {
		return maxRowSize;
	}

	public void setMaxRowSize(int maxRowSize) {
		this.maxRowSize = maxRowSize;
	}
	
	/**
	 * 标题行所在行号
	 * 
	 * @author zhang zixiao
	 * @return
	 */
	public int getTitleIndex() {
		return titleIndex;
	}

	public void setTitleIndex(int titleIndex) {
		this.titleIndex = titleIndex;
	}

	/**
	 * 获取当前sheet的编号
	 * @return
	 */
	public int getCurrentSheetNO(){
		if(maxRowSize<=0){
			return 0;
		}else{
			return index/maxRowSize;
		}
	}
	
	public void addRowIndex() {
		index++;
		
	}

	public Map<String, Map<String, String>> getDataMapper() {
		return dataMapper;
	}

	public void setDataMapper(Map<String, Map<String, String>> dataMapper) {
		this.dataMapper = dataMapper;
	}

	/**
	 * 获取需要合并行的标题
	 * 
	 * @author zhang zixiao
	 * @return
	 */
	public Set<String> getMergeTitles() {
		return mergeTitles;
	}

	public void setMergeTitles(Set<String> mergeTitles) {
		this.mergeTitles = mergeTitles;
	}

	public Map<String, Object> getMergeTitleDate() {
		return mergeTitleDate;
	}

	public void setMergeTitleDate(Map<String, Object> mergeTitleDate) {
		this.mergeTitleDate = mergeTitleDate;
	}

	public void setMergeRelyLeftTitle(Set<String> mergeRelyLeftTitle) {
		this.mergeRelyLeftTitle = mergeRelyLeftTitle;
		
	}

	public Set<String> getMergeRelyLeftTitle() {
		return mergeRelyLeftTitle;
	}

	/**
	 * 表格有落款的情况最大行号减去落款需要的行数
	 * @param length
	 */
	public void setRootRows(int length) {
		maxRowSize -= length; 
		
	}

	public void setRootIndexs(int[] rootIndex) {
		this.rootIndexs = rootIndex;
	}

	public int[] getRootIndexs() {
		return rootIndexs;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public ExcelSheetRoot getExcelSheetRoot() {
		return excelSheetRoot;
	}

	public void setExcelSheetRoot(ExcelSheetRoot excelSheetRoot) {
		this.excelSheetRoot = excelSheetRoot;
	}

	
}
