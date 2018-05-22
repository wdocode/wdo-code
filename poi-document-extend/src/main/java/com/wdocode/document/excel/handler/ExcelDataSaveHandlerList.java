package com.wdocode.document.excel.handler;

import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * 
 * excel 数据保存实现类[将解析出的数据保存至list集合中]
 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
 * @since 2016年4月5日 下午3:12:12
 *
 */
@Deprecated
public class ExcelDataSaveHandlerList extends ExcelDataSaveHandlerDefault {
	
	/**对应mybitis保存sql id */
	private String sqlStatementMappedID;
	/** 每次最大保存条数 该值必须大于0 
	 * 若sqlMappedStatementID对应的是单条记录保存sql 请将该值设置为1
	 * 若sqlMappedStatementID对应的sql支持批量插入请将该值设置为大于1的整数*/
	private int defMaxPreSaveSize = 1;
	private Map<String, List<Map<String, Object>>> list = new HashMap<String, List<Map<String, Object>>>();
	
	/**
	 * 若sqlMappedStatementID对应的是单条记录保存sql 请将maxPreSaveSize设置为1
	 * 若sqlMappedStatementID对应的sql支持批量插入请将maxPreSaveSize设置为大于1的整数
	 * @param sqlMappedStatementID 对应mybitis保存sql id
	 * @param maxPreSaveSize 每次保存最大条数 默认为1
	 */
	public ExcelDataSaveHandlerList(String sqlStatementMappedID,int maxPreSaveSize) {
		setSqlStatementMappedID(sqlStatementMappedID);
		setMaxPreSaveSize(maxPreSaveSize);
	}
	

	public void setSqlStatementMappedID(String sqlStatementMappedID) {
		this.sqlStatementMappedID = sqlStatementMappedID;
	}



	public void setMaxPreSaveSize(int maxPreSaveSize) {
		if(maxPreSaveSize>1)
			this.defMaxPreSaveSize = maxPreSaveSize;
	}

	public String getSQLStatementMappedID() {
		return sqlStatementMappedID;
	}

	@Override
	public int maxPreSaveSize() {
		return defMaxPreSaveSize;
	}

	@Override
	public int save(Map<String, List<Map<String, Object>>> dataModel){
		if(dataModel == null)
		System.out.println(dataModel.size());
//		System.out.println(JSONArray.toJSON(dataModel));
		list.putAll(dataModel);
		return dataModel.size();
	}

}
