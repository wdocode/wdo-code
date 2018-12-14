package com.wdocode.document.excel.handler;

import java.util.List;
import java.util.Map;

/**
 * 保存excel数据  回调方法抽象类
 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
 *
 */
public abstract class ExcelDataSaveHandlerDefault implements ExcelDataSaveHandler {

	public int maxPreSaveSize() {
		return 20;
	}

	public abstract int save(Map<String, List<Map<String, Object>>> cacheData);

}
