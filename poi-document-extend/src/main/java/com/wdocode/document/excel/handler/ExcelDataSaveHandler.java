package com.wdocode.document.excel.handler;

import java.util.List;
import java.util.Map;

/**
 * 保存excel数据  回调接口
 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
 * @since 2016年4月1日 下午4:09:38
 *
 */
public interface ExcelDataSaveHandler {
	
	
	/**
	 * 获取每次保存最多条数
	 * @return
	 */
	public int maxPreSaveSize();
	
	/**
	 * 执行保存方法
	 * @param dataModel
	 * @return
	 */
	public int save(Map<String, List<Map<String, Object>>> cacheData);

}
