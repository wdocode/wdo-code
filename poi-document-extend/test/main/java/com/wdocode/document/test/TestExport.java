package com.wdocode.document.test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.wdocode.document.excel.EXCELDocumentBuilder;
import com.wdocode.document.excel.model.ExcelColumn;
import com.wdocode.document.excel.model.ExcelDataMapper;
import com.wdocode.document.excel.model.ExcelSheet;

public class TestExport {
	
	public static void main(String[] args) throws IOException {
	
		ExcelDataMapper datamapper = new ExcelDataMapper();
		datamapper.setFileName("d:/hell.xlsx");
		datamapper.setFileSuffix(".xlsx");
		List<ExcelSheet> sheets = new ArrayList<ExcelSheet>();
		ExcelSheet e = new ExcelSheet("导出测试", "sample_list");
		List<ExcelColumn> columns = new ArrayList<ExcelColumn>();
		ExcelColumn cl = new ExcelColumn("标题", "title");
		columns.add(cl );
		e.setColumns(columns );
		sheets.add(e );
		datamapper.setSheets(sheets);
		
	 File file = new File(datamapper.getFileName());
	 System.out.println(file.getAbsolutePath());
	 if(!file.exists())
		 file.createNewFile();
	 FileOutputStream out = new FileOutputStream(file);
	 EXCELDocumentBuilder builder = EXCELDocumentBuilder.build(datamapper, out);
	 Map<String,String> rowData = new HashMap<String,String>();
	 List<Map<String,Object>> list = new ArrayList<Map<String,Object>>();
	
	 boolean status = builder.pushData("sample_list", list);
	 builder.write();
	 out.close();
	 
		
		
	}

}
