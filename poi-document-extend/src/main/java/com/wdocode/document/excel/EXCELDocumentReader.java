package com.wdocode.document.excel;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.wdocode.document.excel.handler.ExcelDataSaveHandler;
import com.wdocode.document.excel.model.ExcelColumn;
import com.wdocode.document.excel.model.ExcelDataMapper;
import com.wdocode.document.excel.model.ExcelSheet;



/**
 * 
 * excel 文档读取工具类
 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
 * @since 2016年4月5日 下午5:34:44
 *
 */
public class EXCELDocumentReader {
	private static Logger logger = LoggerFactory.getLogger(EXCELDocumentReader.class);
	
	private ExcelDataMapper datamapper;
	private ExcelDataSaveHandler insertDataHandler;
	private int countsize = 0;
	private Map<String, List<Map<String, Object>>> cacheData = 
			new HashMap<String, List<Map<String, Object>>>();
	
	
	
	public static EXCELDocumentReader build(ExcelDataMapper datamapper,
			ExcelDataSaveHandler insertDataHandler){
		return new EXCELDocumentReader(datamapper, insertDataHandler);
	}
	/**
	 * excel文档读取工具构造函数
	 * @param fileName
	 * @param datamapper
	 * @param insertDataHandler
	 */
	protected EXCELDocumentReader(ExcelDataMapper datamapper,
			ExcelDataSaveHandler insertDataHandler) {
		setDatamapper(datamapper);
		setInsertDataHandler(insertDataHandler);
	}
	
	/**
	 * 读取excel文件数据
	 * 
	 * @param inputStream
	 * @return
	 * @throws IOException 
	 * @throws ParseException 
	 */
	public boolean readExcelData(String fileName,InputStream inputStream) 
			throws IOException{
		try {
			Workbook wb = createWorkBook(fileName,inputStream);
			List<SheetTitle> sheets = getImportSheetTitle(wb);
			if(sheets == null || sheets.isEmpty()){
				throw new IOException(String.format(
						"读取excel文件失败,excel文件[%s]中未能匹配到相同的sheet_mapper描述",fileName));
			}
			for(SheetTitle s: sheets){
				Sheet sheet = wb.getSheet(s.sheetName);
				if(!readSheet(sheet,s)){
					logger.warn("读取excel数据失败");
					return false;
				}
			}
			return true;
			
		} catch (ParseException e) {
			throw new IOException("读取EXCEL文件失败",e);
		} catch (Exception e) {
			throw new IOException("读取EXCEL文件失败",e);
		}
		
	}
	
	
	/**
	 * 读取sheet表格数据
	 * @param sheet
	 * @param s
	 * @return
	 * @throws ParseException 
	 */
	private boolean readSheet(Sheet sheet, SheetTitle title) throws ParseException {
		if(sheet == null)
			return false;
		for(int i=1,s = sheet.getLastRowNum();i<=s;i++){
			Object[] rowdata = EXCELDocumentCommon.getRowDate(sheet.getRow(i));
			Map<String, Object> saveData = mapperData(rowdata,title);
			List<Map<String, Object>> list = cacheData.get(title.map_key);
			if(list == null)
				list = new ArrayList<Map<String, Object>>();
			list.add(saveData);
			cacheData.put(title.map_key, list);
			saveData(false);
		}
		saveData(true);
		return true;
		
	}


	/**
	 * 将要sheet中的行数据封装到map中
	 * @param rowdata
	 * @param title
	 * @return
	 * @throws ParseException 
	 */
	private Map<String, Object> mapperData(Object[] rowdata, SheetTitle sheetTitle) throws ParseException {
		if(sheetTitle == null || sheetTitle.title == null)
			return null;
		ExcelSheet mapper = datamapper.getSheetByMapKey(sheetTitle.map_key);
		if(mapper == null || mapper.getColumns() == null){
			return null;
		}
		List<ExcelColumn> cls = mapper.getColumns();
		Map<String, Object> result = new HashMap<String, Object>();
		int max = sheetTitle.title.length>rowdata.length?sheetTitle.title.length:rowdata.length;
		for(int i=0;i<max;i++){
			String t = sheetTitle.title[i];
			if(t == null)
				continue;
			for(ExcelColumn c:cls){
				if(c == null)
					continue;
				if(t.equalsIgnoreCase(c.getColumn()) || t.equalsIgnoreCase(c.getTilte())){
					result.put(c.getColumn(), mapperCellData(c,rowdata[i]));
				}
			}
		}
		return result;
	}

	/**
	 * 映射单元格数据
	 * @param excelColumn
	 * @param cellData
	 * @return
	 * @throws ParseException 
	 */
	private Object mapperCellData(ExcelColumn c, Object cellData) {
		
		try {
			Map<String, String> datam = c.getDataMapper();
			
			SimpleDateFormat sd = c.getFormat_date();
			DecimalFormat numformat = c.getFormat_number();
			if(cellData == null){
				return mapperCellData(datam,null);
				
			}else if(cellData instanceof String){
				if(datam != null)
					return mapperCellData(datam,(String)cellData);
				if(sd!=null){
					return sd.parse((String)cellData);
				}
				
			}else if(cellData instanceof Double){
				String temp = null;
				if(numformat != null){
					temp = numformat.format(cellData);
				}
				if(datam != null){
					temp =temp == null?EXCELDocumentCommon.formateData(cellData, null):temp; 
					return mapperCellData(datam,temp);
				}
				
			}else if(cellData instanceof Date){
				String temp = null;
				if(sd != null){
					temp = sd.format(cellData);
				}
				temp =temp == null?EXCELDocumentCommon.formateData(cellData, null):temp;
				return mapperCellData(datam,temp);
				
			}else if(cellData instanceof Boolean){
				if(datam != null)
					return mapperCellData(datam,String.valueOf(cellData));
				return cellData;
			}
			
			return mapperCellData(datam,String.valueOf(cellData));
		} catch (Exception e) {
			throw new RuntimeException("["+c.getTilte()+"="+String.valueOf(cellData)+"]数据转换失败",e);
		}
	}


	private Object mapperCellData(Map<String, String> datam, String cellData) {
		if(datam == null)
			return cellData;
		String val = null;
		if(cellData == null){
			val = datam.get("NULL");
		}else{
			val = datam.get(cellData.toUpperCase());
		}
		if(val == null)
			val = datam.get("DEFAULT");
		
		return val == null?cellData:val;
	}
	
	/**
	 * 保存数据方法
	 * @param last 是否最后一次
	 */
	private void saveData(boolean last) {
		if(last || ++countsize>=insertDataHandler.maxPreSaveSize()){
			int size = insertDataHandler.save(cacheData);
			logger.debug("写数据完成，本次共写入数据条数："+size);
			cacheData.clear();
			countsize =0;
		}
		
	}

	/**
	 * 获取需要导入的sheet表头
	 * 表格名（sheetName）必须与sheet_mapper映射中设置的sheet_name或map_key相同
	 * @param wb
	 * @return
	 * @throws IOException 
	 */
	private List<SheetTitle> getImportSheetTitle(Workbook wb) throws IOException {
		if(datamapper == null)
			throw new IOException("无效的excel映射说明：datamapper="+datamapper);
		
		List<ExcelSheet> sheets = datamapper.getSheets();
		if(sheets == null || sheets.size() == 0){
			throw new IOException("无效的excel映射说明：表格说明sheet_mapper="+sheets);
		}
		Map<String, String[]> title = getSheetTitle(wb);
		if(title == null || title.isEmpty()){
			throw new IOException("无效的excel文件：获取表格标题失败");
		}
		List<SheetTitle> result = new ArrayList<SheetTitle>();
		for(String k :title.keySet()){
			if(k == null)
				continue;
			for(ExcelSheet s:sheets){
				if(k.equals(s.getMapKey()) || k.equals(s.getSheetName())){
					result.add(new SheetTitle(k, s.getMapKey(), title.get(k)));
					break;
				}
			}
		}
		return result;
	}

	/**
	 * 取表格首行 作为标题
	 *
	 * @param wb
	 * @return
	 */
	private static Map<String, String[]> getSheetTitle(Workbook wb) {
		if (wb == null)
			return null;
		Map<String, String[]> titles = new HashMap<String, String[]>();
		for (int i = 0,l = wb.getNumberOfSheets(); i < l; i++) {
			Sheet st = wb.getSheetAt(i);
			String sheetName = st.getSheetName();
			Row row = st.getRow(0);
			if (row == null) {
				continue;
			}
			int s = row.getFirstCellNum(),m=row.getLastCellNum();
			String[] title = new String[m];
			Arrays.fill(title, "");
			for(;s<m;s++){
				title[s] =  EXCELDocumentCommon.getCellDataToString(row.getCell(s), null);
			}
			titles.put(sheetName, title);
		}
		return titles ;
	}
	
    
	/**
	 * 根据文件后缀创建excel文件工作簿实例
	 * @author <a href="mailto:zhangzixiao@foreveross.com">zhang zixiao</a>
	 * @param fileName2
	 * @param inputStream
	 * @return
	 * @throws IOException
	 */
	private Workbook createWorkBook(String fileName, InputStream inputStream) throws IOException {
		if(fileName == null || fileName.indexOf(".")<0){
			throw new IOException("无效的excel文件名："+fileName);
		}
		String file_suffix = fileName.substring(fileName.lastIndexOf("."));
		if (EXCELDocumentCommon.EXCEL_SUFFIX_XLS.equalsIgnoreCase(file_suffix)){
			return new HSSFWorkbook(inputStream); // EXCEL OFFICE2007 以前的版本
		} else if (EXCELDocumentCommon.EXCEL_SUFFIX_XLSX.equalsIgnoreCase(file_suffix)) {
			return new XSSFWorkbook(inputStream);// EXCEL OFFICE2007之后的版本
		} else if (EXCELDocumentCommon.EXCEL_SUFFIX_CSV.equalsIgnoreCase(file_suffix)) {
			return new HSSFWorkbook(inputStream);// CSV 表格文件读取
		} else
			throw new IOException("无法识别的文件：" + fileName);
	}


	public ExcelDataMapper getDatamapper() {
		return datamapper;
	}

	public void setDatamapper(ExcelDataMapper datamapper) {
		this.datamapper = datamapper;
	}

	public ExcelDataSaveHandler getInsertDataHandler() {
		return insertDataHandler;
	}

	public void setInsertDataHandler(ExcelDataSaveHandler insertDataHandler) {
		this.insertDataHandler = insertDataHandler;
	}
	
	
	private class SheetTitle{
		public String sheetName;
		public String map_key;
		public String title[];
		
		public SheetTitle(String sheetName,String map_key,String [] titles) {
			this.sheetName = sheetName;
			this.map_key = map_key;
			this.title = titles;
		}
	}
}
