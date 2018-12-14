package com.wdocode.mybatis.keygenerator.util;

import java.util.UUID;

/**
 * uuid辅助类
 * @author zhangzx
 *
 */
public class UUIDHelper {

	/**
	 * 获取32位uuid
	 * @return
	 */
	public static String uuid() {
		String uuid = UUID.randomUUID().toString();
		return uuid.replaceAll("-", "");
	}
	
	

}
