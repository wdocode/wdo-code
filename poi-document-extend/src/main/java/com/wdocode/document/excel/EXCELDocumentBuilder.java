package com.wdocode.document.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.wdocode.document.common.IOUtil;
import com.wdocode.document.excel.model.ExcelColumn;
import com.wdocode.document.excel.model.ExcelDataMapper;
import com.wdocode.document.excel.model.ExcelReplaceHolder;
import com.wdocode.document.excel.model.ExcelSheet;
import com.wdocode.document.excel.model.ExcelSheetRoot;
import com.wdocode.document.excel.model.ExcelTemplate;
import com.wdocode.document.excel.model.ExcelTemplateReplaceRow;
import com.wdocode.document.excel.model.SheetRowDataCounter;


/**
 * <pre>EXCEL 文件生成器
 * export data into excel file util <pre>
 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
 * @since POI3.15,JDK1.7+ <code>2016年4月5日 下午5:34:44</code>
 *
 */
public final class EXCELDocumentBuilder {
    
	private static  Logger logger = Logger.getLogger("EXCELDocumentBuilder");
	
	private static final short prePixelSize = 16;
	// 行计数器
    private Map<String, SheetRowDataCounter> rowDataCounter = 
    		new HashMap<String, SheetRowDataCounter>();
    // sheet 合并行坐标
    private Map<String, List<CellRangeAddress>> sheetRangeAddress = 
    		new HashMap<String, List<CellRangeAddress>>();
    // 数据映射
	private ExcelDataMapper datamapper;
	private OutputStream out;
	private Workbook workbook;
	// 模板
	private InputStream tempInputStream;
	// 表格最大行号
	private int maxRowSize;
	
	/**
	 * 创建EXCEL文档构造器
	 * <pre>获取excel文档生成器实例
	 * 	<p>示例</p>
	 * ExcelDataMapper datamapper = ExcelDataMapperReader.build()
	 * 			.getDataMapper("export_sample", "export_sample_mapper.xml");
	 * File file = new File(datamapper.getFileName());
	 * System.out.println(file.getAbsolutePath());
	 * FileOutputStream out = new FileOutputStream(file);
	 * EXCELDocumentBuilder builder = EXCELDocumentBuilder.build(datamapper, out);
	 * Map rowData = new HashMap();
	 * rowData.put("notice_list", list);
	 * Map replaceData = new HashMap();
	 * replaceData.put("占位符", "替换文本");
	 * builder.replaceTextHolder("sample_list", replaceData);
	 * boolean status = builder.pushData("sample_list", list);
	 * builder.pushData("sample_list", getDate("ZK"));
	 * builder.write();
	 * out.close();
	 * </pre>
	 * 
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param datamapper
	 * @param out
	 * @return
	 */
	public static EXCELDocumentBuilder build(ExcelDataMapper datamapper, OutputStream out) {
		EXCELDocumentBuilder builder = new EXCELDocumentBuilder(datamapper,out);
		return builder;
	}
	
	/**
	 * excel文档创建器构造方法
	 * @param datamapper
	 * @param out
	 */
	protected EXCELDocumentBuilder(ExcelDataMapper datamapper, OutputStream out) {
		this.datamapper = datamapper;
		this.out = out;
	}

	/**
	 * 创建excel文件
	 * 
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param list
	 * @param datamapper
	 * @param out
	 * @throws IOException 
	 */
	public Workbook createWorkbook() throws IOException {
		if(datamapper == null){
			logger.error("创建目标excel文件失败：excel描述对象为null");
			return null;
		}
		if(StringUtils.isBlank(datamapper.getFileName())){
			logger.error("创建目标excel文件失败：目标文件名不能为空");
			return null;
		}
		String file_suffix = datamapper.getFileSuffix();
		maxRowSize = 0; 
		if(EXCELDocumentCommon.EXCEL_SUFFIX_XLS.equalsIgnoreCase(file_suffix)){
			maxRowSize = EXCELDocumentCommon.SHEETMAXROWS_2003;
		}else if(EXCELDocumentCommon.EXCEL_SUFFIX_XLSX.equalsIgnoreCase(file_suffix)){
			maxRowSize =EXCELDocumentCommon.SHEETMAXROWS_2007;
		}
		
		String tempPath = datamapper.getTemplateFilePath();
		if(StringUtils.isNotBlank(tempPath))
			return createWorkbookFormTemplate(tempPath);
		
		if(EXCELDocumentCommon.EXCEL_SUFFIX_XLS.equalsIgnoreCase(file_suffix)){
			return new HSSFWorkbook();
		}else if(EXCELDocumentCommon.EXCEL_SUFFIX_XLSX.equalsIgnoreCase(file_suffix)){
			return new SXSSFWorkbook(1000);
		}else if(EXCELDocumentCommon.EXCEL_SUFFIX_CSV.equalsIgnoreCase(file_suffix)){
			return new SXSSFWorkbook(1000);
		}else{
			throw new IOException("当前不支持的excel工作簿文件格式："+file_suffix);
		}
	}

	/**
	 * 通过EXCEL文件模板创建 工作簿
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param templateFilePath
	 * @return
	 * @throws IOException 
	 */
	private Workbook createWorkbookFormTemplate(String templateFilePath) throws IOException {
		if(templateFilePath == null){
			throw new FileNotFoundException("模板路径不能为空");
		}
		tempInputStream = new FileInputStream(templateFilePath);
		String temp_suffix = templateFilePath.substring(templateFilePath.lastIndexOf("."));
		if(EXCELDocumentCommon.EXCEL_SUFFIX_XLS.equalsIgnoreCase(temp_suffix)){
			return new HSSFWorkbook(tempInputStream);
		}else if(EXCELDocumentCommon.EXCEL_SUFFIX_XLSX.equalsIgnoreCase(temp_suffix)){
			return new XSSFWorkbook(tempInputStream);
		}else if(EXCELDocumentCommon.EXCEL_SUFFIX_CSV.equalsIgnoreCase(temp_suffix)){
			return new HSSFWorkbook(tempInputStream);
		}else{
			throw new IOException("当前不支持的excel工作簿文件格式："+temp_suffix);
		}
	}
	
	
	/**
	 * 替换模板中占位符 
	 * 
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param sheetMapKey
	 * @param res
	 */
	public boolean replaceTextHolder(String sheetMapKey, 
			Map<String, String> replaceData) throws IOException {
		if(datamapper == null || 
				StringUtils.isBlank(datamapper.getTemplate_fileName()))
			return true;
		// 写数据
		Sheet sheet = getSheet(sheetMapKey);
		if(sheet == null){
			logger.warn("请检查excel描述文件中是否存在相同mapkey的sheet_mapper节点：map_key:"+sheetMapKey);
			return false;
		}
		ExcelTemplate temp = datamapper.getSheetByMapKey(sheetMapKey).getTemplate();
		if(temp == null || temp.getReplacerows() == null || replaceData == null)
			return true;
		List<ExcelTemplateReplaceRow> list = temp.getReplacerows();
		for(ExcelTemplateReplaceRow r : list){
			Row row = sheet.getRow(r.getRowIndex());
			if(row == null)
				continue;
			for(int s =row.getFirstCellNum(),l=row.getLastCellNum();s<l;s++){
				Cell cell = row.getCell(s);
				Object oldStr = EXCELDocumentCommon.getCellData(cell);
				if(oldStr instanceof String)
				 cell.setCellValue(
					replaceCellStr((String)oldStr,r.getReplaceHolder(),replaceData));
			}
			
		}
		return true;
		
	}

	/**
	 * 填充数据
	 * 
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param sheet
	 * @param list
	 * @return
	 * @throws IOException 
	 */
	public boolean pushData(String sheetMapKey, 
			List<Map<String, Object>> list) throws IOException {
		// 写数据
		Sheet sheet = getSheet(sheetMapKey);
		if( sheet == null){
			logger.warn("请检查excel描述文件中是否存在相同mapkey的sheet_mapper节点：map_key:"+sheetMapKey);
			return false;
		}
		SheetRowDataCounter ct = rowDataCounter.get(sheetMapKey);
		if(ct == null){
			logger.warn("创建计数器失败,map_key："+ sheetMapKey);
			return false;
		}
		if(list == null)
			return true;
		for(int i = 0,len = list.size();i<len;i++){
			logger.debug(String.format("mapkey=%s,Index=%d,nextindex=%d", 
					sheetMapKey,ct.getIndex(),ct.getNextIndex()));
			addRowDataToSheet(getSheet(sheetMapKey),list.get(i),ct);
		}
		rowDataCounter.put(sheetMapKey, ct);
		
		return true;
	}
	
	/**
	 * 获取新的sheet对象
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param workbook
	 * @param mapper
	 * @param sheetnum
	 * @return
	 * @throws IOException 
	 */
	private Sheet getSheet(String key) throws IOException {
		/*
		 *1 、判断来自模板的情况获取标题
		 *2、获取当前行号 
		 * */
		if(datamapper == null)
			return null;
		ExcelSheet sheetMapper = datamapper.getSheetByMapKey(key);
		if(sheetMapper == null){
			logger.warn("没找map_key到对应sheetmap描述,map_key:"+key);
			return null;
		}
		if(workbook == null)
			workbook = createWorkbook();
		if(workbook == null)
			throw new IOException("创建工作簿失败");
		
		SheetRowDataCounter ct = rowDataCounter.get(key);
		
		if(ct == null){
			Sheet sheet = workbook.getSheet(sheetMapper.getSheetName());
			sheet = sheet == null?workbook.createSheet(sheetMapper.getSheetName()):sheet;
			ct = initCounter(sheet,sheetMapper);
			ct.setSheetName(sheetMapper.getSheetName());
			rowDataCounter.put(key, ct);
			initSheet(sheet,sheetMapper);
			return sheet;
		}else{
			return getCurrentSheet(ct,sheetMapper);
		}
	}
	
	/**
	 * 获取当前sheet页
	 * @param ct
	 * @param sheetMapper
	 * @return
	 */
	private Sheet getCurrentSheet(SheetRowDataCounter ct,ExcelSheet sheetMapper) {
		Sheet sheet = workbook.getSheet(sheetMapper.getSheetName());
		int currentno = ct.getCurrentSheetNO();
		// 获取当前sheet页码为首页时返回当前sheet
		if(currentno<=0)
			return sheet;
		// 新建sheet页
		String sName = sheetMapper.getSheetName()+currentno;
		Sheet sheet2 = workbook.getSheet(sName);
		if(sheet2 != null)
			return sheet2;
		
		sheet2 = workbook.createSheet(sName);
		sheet2.setDefaultColumnWidth(sheet.getDefaultColumnWidth());
		sheet2.setDefaultRowHeight(sheet.getDefaultRowHeight());
		// TODO 复制过程中样式有丢失情况
		// 将上一页的标题行复制到当前页面
		
		for(int i=0,t=ct.getTitleIndex();i<=t;i++){
			
			if(i > 0)
				ct.addRowIndex();
			Row row = sheet.getRow(i);
			if(row == null){
				return sheet2;
			}
			Row row2 = sheet2.createRow(i);
			if(row.getRowStyle()!=null)
				row2.setRowStyle(row.getRowStyle());
			for(short j = 0;j<row.getLastCellNum();j++){
				Cell c1 = row.getCell(j);
				Cell c2 = row2.createCell(j);
				c2.setCellValue(c1.getRichStringCellValue());
				c2.setCellType(c1.getCellType());
				c2.setCellStyle(c1.getCellStyle());
			}
		}
		initSheet(sheet2,sheetMapper);
		return sheet2;
	}


	/**
	 * 新建sheet页时 初始化表格样式
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param sheet
	 * @param sheetMapper
	 */
	private void initSheet(Sheet sheet, ExcelSheet sheetMapper) {
		List<ExcelColumn> cls = sheetMapper.getColumns();
		for(int i=0,size = cls.size();i<size;i++){
			ExcelColumn col = cls.get(i);
			//列宽
			int width = col.getWidth()>0?col.getWidth():sheetMapper.getCellWidth();
			if(width>0)
				sheet.setColumnWidth(i, prePixelSize*width);
			Integer height = prePixelSize*sheetMapper.getRowHeight();
			if(height>0)
				sheet.setDefaultRowHeight(height.shortValue());
		}
	}

	/**
	 * 新建sheet表格时 初始化计数器
	 * 
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param sheet
	 * @param sheetMapper
	 * @return
	 */
	private SheetRowDataCounter initCounter(Sheet templateSheet, ExcelSheet sheetMapper) {
		// 在模板中初始化计数器
		SheetRowDataCounter counter = initFromTemplate(templateSheet,sheetMapper.getTemplate());
		counter = initCounterFromMapper(counter,templateSheet,sheetMapper);
		return counter;
	}

	/**
	 * 在datamapper中 初始化计数器需要的标题数据映射关系
	 * @param counter
	 * @param sheet
	 * @param sheetMapper
	 * @return
	 */
	private SheetRowDataCounter initCounterFromMapper(
			SheetRowDataCounter counter,Sheet sheet, ExcelSheet sheetMapper) {
		boolean reset =counter == null || counter.getTitle() == null;
		List<ExcelColumn> cls = sheetMapper.getColumns();
		int columnSize = cls.size();
		String[] titletext = null;
		if(reset){
			counter = new SheetRowDataCounter(counter == null?0:counter.getTitleIndex());
			// 没有在模板中获取到excel标题行时 根据xml中配置的映射关系初始化标题
			titletext = new String[columnSize];
			counter.setMaxRowSize(maxRowSize);
		}
		
		String column_keys[] = reset?new String[columnSize]:counter.getTitle();
		CellStyle [] cellstyles = reset?new CellStyle [columnSize]:counter.getCellstyles();
		
		Map<String, Map<String,String>> celldataMapper = new HashMap<String, Map<String,String>>();
		for(int i=0;i<columnSize;i++){
			ExcelColumn col = cls.get(i);
			initMergeTitleMapper(col,counter);
			
			Map<String, String> mapper = col.getDataMapper();
			if(mapper != null)
				celldataMapper.put(col.getColumn(), mapper);
			if(reset){
				// 根据xml配置项重新获取的 [数据-标题]映射关系时重新初始样式等问题
				column_keys[i] = col.getColumn();
				titletext[i]=StringUtils.isNotBlank(col.getTilte())?col.getTilte():col.getColumn();
				cellstyles[i] = getDefaultStyle(col.getDataformat(), col.getHorizontalAlign(),col.getVerticalAlign());
			}
		}
		counter.setDataMapper(celldataMapper);
		// 没有模板、或在模板中没有获取到标题行时 重新添加标题行到excel中
		if(reset){
			addRowDataToSheet(sheet,counter.getIndex(),titletext,cellstyles);
		}
		counter.setCellstyles(cellstyles);
		counter.setTitle(column_keys);
		return counter;
	}

	/**
	 * 初始化需要合并行的列配置 
	 * @param col
	 * @param counter
	 */
	private void initMergeTitleMapper(ExcelColumn col, SheetRowDataCounter counter) {
		if(!col.isMergeRow())
			return;
		Set<String> mergeTitles =counter.getMergeTitles();
		Set<String> mergeRelyLeftTitle= counter.getMergeRelyLeftTitle();
		// 判断是否有需要合并行的数据内容
		if(mergeTitles == null)
			mergeTitles = new HashSet<String>();
		mergeTitles.add(col.getColumn());
		if(col.isMergeRelyLeft()){
			if(mergeRelyLeftTitle == null)
				mergeRelyLeftTitle = new HashSet<String>();
			mergeRelyLeftTitle.add(col.getColumn());
		}
		counter.setMergeTitles(mergeTitles);
		counter.setMergeRelyLeftTitle(mergeRelyLeftTitle);
	}

	/**
	 * 根据模板配置初始化计数器
	 * @param sheet
	 * @param template
	 * @return
	 */
	private SheetRowDataCounter initFromTemplate(Sheet sheet,ExcelTemplate template) {
		// 有模板情况 获取标题所在行
		if(template == null)
			return null;
		int index = template.getTitleIndex();
		Row row = sheet.getRow(index);
		String column_keys[] = null;
		CellStyle [] cellstyles = null;
		if(row != null){
			int last = row.getLastCellNum();
			column_keys = new String[last];
			cellstyles = new CellStyle [last];
			for(short i=row.getFirstCellNum(); i<last; i++) {
				Cell cell = row.getCell(i);
				if(cell != null){
					cellstyles[i]=cell.getCellStyle();
					column_keys[i] = EXCELDocumentCommon.getCellDataToString(cell, null);
					if(template.isRemoveTitle())
						cell.setCellValue("");
				}
			}
			// 需要删除标题行 则将游标上移一行
			if(template.isRemoveTitle()){
				index -= 1;
			}
		}
		SheetRowDataCounter counter = new SheetRowDataCounter(index, maxRowSize, column_keys, cellstyles);
		// 获取落款行号
		readTemplateRootRows(counter,sheet,template.getRootIndex());
		
		return counter;
				
	}

	/**
	 * 
	 * 在模板中读取尾行数据
	 * @param counter
	 * @param sheet
	 * @param rootIndex
	 */
	private void readTemplateRootRows(SheetRowDataCounter counter,Sheet sheet, int[] rootIndex) {
		if(rootIndex == null || rootIndex.length == 0)
			return;
		
		 List<String []> cellList = new ArrayList<String []>();
		/**单元格样式 */
		 List<CellStyle []> cellStyles = new ArrayList<CellStyle []>();
		/**合并单元格坐标 */
		 List<CellRangeAddress> needMergeRow = new ArrayList<CellRangeAddress>();
//		 List<CellRangeAddress> regions =  sheet.getMergedRegions(); 
		counter.setRootRows(rootIndex.length);
		for(int i=0;i<rootIndex.length;i++){
			Row row = sheet.getRow(rootIndex[i]);
			if(row == null)
				break;
			// 先读取合并单元格坐标
//			if(regions != null){
//				for(CellRangeAddress r:regions){
//					if(r.getFirstRow() == rootIndex[i]){
//						needMergeRow.add(new CellRangeAddress(r.getFirstRow(),
//								r.getLastRow(), r.getFirstColumn(), r.getLastColumn()));
//					}
//				}
//			}
			// 获取数据 
			int length=row.getLastCellNum();
			String [] cellData = new String [length];
			CellStyle[] styles  = new CellStyle [length];
			for(short j=0;j<length;j++){
				Cell cell = row.getCell(j);
				styles[j] = cell.getCellStyle();
				cellData[j] = EXCELDocumentCommon.getCellDataToString(cell, null);
				removeMergedRegion(sheet, rootIndex[i], j);
			}
			cellList.add(cellData);
			cellStyles.add(styles);
			
			// 删除尾行
			sheet.removeRow(row);
		}
		if(!cellList.isEmpty()){
			ExcelSheetRoot roots = new ExcelSheetRoot(cellList,cellStyles,needMergeRow);
			counter.setExcelSheetRoot(roots);
		}
		counter.setRootIndexs(rootIndex);
		
	}

	/**
	 * SHEET添加行
	 * 
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param sheet
	 * @param index
	 * @param rowData
	 * @param cellstyles 
	 */
	private void addRowDataToSheet(Sheet sheet, int index, String[] rowData, CellStyle[] cellstyles) {
		if(rowData == null || sheet == null){
			logger.debug(String.format(
					"add RowData To Sheet params:rowData.size=%d,sheet is null=%s", 0,sheet == null));
			return;
		}
			
		Row row = sheet.getRow(index);
		if(row == null)
			row = sheet.createRow(index);
		for(int i=0;i<rowData.length;i++){
			Cell cell = row.getCell(i);
			if(cell == null)
				cell = row.createCell(i);
			if(cellstyles[i]!=null)
				cell.setCellStyle(cellstyles[i]);
			cell.setCellValue(rowData[i]);
		}
		
	}
	
	
	/**
	 * 添加数据到sheet表中
	 * 
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param sheet
	 * @param rowIndex
	 * @param rowData
	 * @param cellstyles 
	 * @param title 
	 */
	private void addRowDataToSheet(Sheet sheet,  
			Map<String, Object> rowData, SheetRowDataCounter ct) {
		int rowIndex = ct.getNextIndex();
//		version 1.0  每次添加都移动最后一行实现尾行标记
//		moveRootToNext(sheet,ct,rowIndex);
		String title[] = ct.getTitle();
//		logger.debug(String.format("add RowData To Sheet params:sheetname=%s,rowIndex=%d",
//				sheet.getSheetName(),rowIndex));
		Map<String, Map<String, String>> mapper = ct.getDataMapper();	
		Row row = sheet.getRow(rowIndex);
		if(row == null)
			row = sheet.createRow(rowIndex);
		String sheetName = sheet.getSheetName();
		CellStyle [] cellstyles = ct.getCellstyles();
		for(int i=0;i<title.length;i++){
			Map<String, String> cl_mapper = mapper.get(title[i]);
			Cell cell = row.getCell(i);
			if(cell == null)
				cell = row.createCell(i);
			if(cellstyles[i]!=null)
				cell.setCellStyle(cellstyles[i]);
			Object data = fillCellData(cl_mapper,cell,rowData==null?"":rowData.get(title[i]));
			// 计算是否有合并行
			addMergeRowIndex(sheetName,ct, rowIndex,cell,title,data,i);
		}
		ct.addRowIndex();
		ct.getCurrentSheetNO();
	}
	
	

	/**
	 * 填写单元格数据
	 * 
	 * @param cl_mapper
	 * @param cell
	 * @param data
	 * @return
	 */
	private Object fillCellData(Map<String, String> cl_mapper,
			 Cell cell,Object data) {
		Object result = null;
		if(data == null){
			result = getMapperVal(cl_mapper,data);
			cell.setCellValue((String)result);
		}else if(data instanceof Date){
			result = (Date)data;
			cell.setCellValue((Date)result);
			CellStyle st = cell.getCellStyle();
			int format = st.getDataFormat();
			if(format<=0 && StringUtils.isNotBlank(datamapper.getDefDataformate())) {
				cell.setCellStyle(getDefaultStyle(datamapper.getDefDataformate(), null, null));
			}
		}else if(data instanceof Number){
			String mappval = getMapperVal(cl_mapper, data);
			if(cl_mapper == null || StringUtils.isBlank(mappval)){
				Number d = (Number)data;
				result = d.doubleValue();
				cell.setCellValue((Double)result);
			}else{
				result = mappval;
				cell.setCellValue(mappval);
			}
			
		}else if(data instanceof Boolean){
			String mappval = getMapperVal(cl_mapper, data);
			result = mappval;
			if(StringUtils.isNotBlank(mappval)){
				cell.setCellValue(mappval);
			}else{
				result = "true".equalsIgnoreCase(String.valueOf(data));
				cell.setCellValue((Boolean)result);
			}
		}else{
			String mappval = getMapperVal(cl_mapper, data);
			result = mappval;
			if(StringUtils.isNotBlank(mappval)){
				cell.setCellValue(mappval);
			}else{
				result = data.toString();
				cell.setCellValue((String)result);
			}
		} 
		
		return result;
	}

	/**
	 * 添加合并行
	 * @param sheetName
	 * @param sheetRowDataCounter
	 * @param currentRow
	 * @param cell 
	 * @param title
	 * @param rowData 
	 * @param colIndex
	 */
	private void addMergeRowIndex(String sheetName, SheetRowDataCounter ct,  
			int currentRow,Cell cell, String[] title, Object rowData, int colIndex) {
		String index_key = title[colIndex]+"_index";
		boolean needMerge = checkNeedMergeRow(ct,title[colIndex],rowData,currentRow,index_key);
		// 不需要合并 退出
		
		if(!needMerge)
			return;
		Map<String, Object> mergedata = ct.getMergeTitleDate();
		int firstRow = (Integer) mergedata.get(index_key);
		
		List<CellRangeAddress> list = sheetRangeAddress.get(sheetName);
		if(list == null){
			
			list = new ArrayList<CellRangeAddress>();
			// 首次匹配到合并行 初始化合并坐标数据 并退出
			CellRangeAddress range = new CellRangeAddress(firstRow, 
					currentRow, colIndex, colIndex);
			list.add(range);
			sheetRangeAddress.put(sheetName, list);
			// 不是计数的第一行 将当前单元格数据清空
			//cell.setCellValue("");
			return ;
		}
		// 不是首次匹配到合并需求情况 计算是否需要依赖左侧合并（左侧分行时右侧也同时分行断开合并）
		boolean newMarage = mergeNextRow(list,ct,currentRow,title,
				colIndex,firstRow,index_key,rowData);
		if(!newMarage){
			// 不需要重新计数 将当前行数据设为空
			//cell.setCellValue("");
		}
		
	}

	/**
	 * 合并非首次匹配的场景
	 * 判断是否依赖左侧是否断开合并
	 * @param list
	 * @param ct
	 * @param currentRow
	 * @param title
	 * @param colIndex
	 * @param firstRow
	 * @param index_key
	 * @param rowData
	 * @return
	 */
	private boolean mergeNextRow(List<CellRangeAddress> list,SheetRowDataCounter ct,int currentRow,
			String [] title,int colIndex,int firstRow,String index_key, Object rowData) {
		Set<String> mergeRelyLeftTitle = ct.getMergeRelyLeftTitle();
		Map<String, Object> mergedata = ct.getMergeTitleDate();
		if(mergeRelyLeftTitle != null && mergeRelyLeftTitle.contains(title[colIndex]) &&colIndex>1){
			Integer leftIndex = (Integer) mergedata.get(title[colIndex-1]+"_index");
			leftIndex = leftIndex == null?0:leftIndex;
			firstRow = leftIndex>firstRow? leftIndex:firstRow;
			mergedata.put(index_key,firstRow);
			mergedata.put(title[colIndex],rowData);
		}
		for(CellRangeAddress r :list){
			// 合并行标记坐标中的首行 和标记的首行相同则将合并行标记坐标最后一行行号更新为当前行号
			if(r.getFirstRow() == firstRow && r.getFirstColumn()==colIndex){
				r.setLastRow(currentRow);
				return false;
			}
		}
		// 没有匹配到相同开始行的情况 重新添加合并行坐标
		list.add(new CellRangeAddress(firstRow, currentRow,
				colIndex, colIndex));
		return true;
		
	}

	/**
	 * 判断是否需要合并行
	 * @param ct
	 * @param title
	 * @param rowData
	 * @param currentRow
	 * @return
	 */
	private boolean checkNeedMergeRow(SheetRowDataCounter ct,String title,
			Object rowData,int currentRow,String index_key) {
		Set<String> marageTitle = null;
		if((marageTitle = ct.getMergeTitles()) == null 
				|| !marageTitle.contains(title)|| title == null 
				|| rowData == null || StringUtils.isBlank(rowData.toString())  ){
			// 没有要合并行的列标题 或者当前列不需要合并行 则跳过
			return false;	
		}
		Map<String, Object> maragedata = ct.getMergeTitleDate();
		if(maragedata == null)
			maragedata = new HashMap<String,Object>();
		ct.setMergeTitleDate(maragedata);
		// 要合并的列数据中上一行数据
		Object lastdata = maragedata.get(title);
		
		if(lastdata == null || !lastdata.equals(rowData)){
			// 上行数据为空 表明为首行数据 不做合并处理 或 当前行数据 与上行数据不相同 重新记录行号
			maragedata.put(title,rowData);
			maragedata.put(index_key,currentRow);
			return false;
		}
		// 取出上行的行号
		int firstRow = (Integer) maragedata.get(index_key);
		// 如果上行的行号大于等于当前行行号 表名已经分页 重新计数
		if(firstRow >= currentRow){
			maragedata.put(title,rowData);
			maragedata.put(index_key,currentRow);
			return false;
		}
		return true;
	}

	/**
	 * 获取字典数据
	 * @param cl_mapper
	 * @param data
	 * @return
	 */
	private String getMapperVal(Map<String, String> cl_mapper, Object data) {
		if(cl_mapper == null)
			return data == null?"":data.toString();
		String key = data == null?"NULL":data.toString();
		String val = cl_mapper.get(key);
		return val == null?"":val;
	}

	/**
	 * 替换单元格中占位符
	 * 
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param oldStr
	 * @param list
	 * @param replaceData 
	 * @return
	 */
	private String replaceCellStr(final String oldStr, 
			List<ExcelReplaceHolder> list, Map<String, String> replaceData) {
		if(oldStr == null || list == null || list.isEmpty() 
				|| replaceData == null ||replaceData.isEmpty())
			return oldStr;
		String newStr = oldStr;
		for(ExcelReplaceHolder h:list){
			if(h == null || StringUtils.isBlank(h.getPlaceholder()))
				continue;
			String vla = replaceData.get(h.getKey());
			newStr = newStr.replace(h.getPlaceholder(), vla == null?"":vla);
		}
		return newStr;
	}
	

	/**
	 * 将工作簿输出到文件
	 * 
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @throws IOException
	 */
	public void write() throws IOException{
		try {
			if(workbook == null){
				logger.error("创建目标excel文件失败：excel工作簿未创建");
				throw new IOException("创建目标excel文件失败：excel工作簿未创建");
			}
			
			if(out == null){
				logger.error("创建目标excel文件失败：输出流OutputStream对象为NULL");
				throw new IOException("创建目标excel文件失败：输出流OutputStream对象为NULL");
			}
//			addRootRows();
//			margeSheetRow();
			workbook.write(out);	
		} catch (IOException e) {
			throw e;
		}finally {
			IOUtil.close(tempInputStream);
		}
		
	}
	

	/**
	 * 获取默认的表头样式
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param dataformat
	 * @param horizontalAlign
	 * @param verticalAlign
	 * @return
	 */
	private CellStyle getDefaultStyle(String dataformat,String horizontalAlign,String verticalAlign) {
		
		CellStyle titleStyle = workbook.createCellStyle();  
		if(dataformat != null){
			DataFormat format = workbook.createDataFormat();
			titleStyle.setDataFormat(format.getFormat(dataformat));
//			HorizontalAlignment hali = HorizontalAlignment.
//					valueOf(horizontalAlign==null?"":horizontalAlign.toUpperCase());
//			if(hali!=null)
//				titleStyle.setAlignment(hali);
//			VerticalAlignment align = VerticalAlignment.valueOf(verticalAlign);
//			titleStyle.setVerticalAlignment(align);
		}
		titleStyle.setWrapText(true);
		return titleStyle;
	}
	
	/**
	 * 移除开始行号为row的合并的单元格
	 * @param sheet
	 * @param row
	 * @param column
	 */
	public static void removeMergedRegion(Sheet sheet,int row ,int column) {    
		int sheetMergeCount = sheet.getNumMergedRegions();//获取所有的单元格   
		for (int i = 0; i < sheetMergeCount; i++) {   
			CellRangeAddress ca = sheet.getMergedRegion(i); //获取第i个单元格  
			if(ca == null)
				continue;
			int firstColumn = ca.getFirstColumn();    
			int lastColumn = ca.getLastColumn();    
			int firstRow = ca.getFirstRow();    
			if(row == firstRow && column >= firstColumn && column <= lastColumn){    
				sheet.removeMergedRegion(i);//移除合并单元格  
			}    
		}  
	} 
}
