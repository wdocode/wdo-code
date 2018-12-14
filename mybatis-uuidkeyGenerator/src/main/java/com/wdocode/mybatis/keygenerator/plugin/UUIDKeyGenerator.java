package com.wdocode.mybatis.keygenerator.plugin;

import java.sql.Statement;

import org.apache.ibatis.executor.Executor;
import org.apache.ibatis.executor.ExecutorException;
import org.apache.ibatis.executor.keygen.KeyGenerator;
import org.apache.ibatis.mapping.MappedStatement;
import org.apache.ibatis.mapping.SqlCommandType;
import org.apache.ibatis.reflection.MetaObject;
import org.apache.ibatis.reflection.wrapper.MapWrapper;
import org.apache.ibatis.session.Configuration;

import com.wdocode.mybatis.keygenerator.util.UUIDHelper;

/**
 * uuid 主键生成器为要插入的对象主键赋值
 * 	前置赋值
 * @author zhangzx
 *
 */
public final class UUIDKeyGenerator implements KeyGenerator {
	
	public UUIDKeyGenerator() {
	}
	
	
	/**
	 *  前置 赋值
	 * @see org.apache.ibatis.executor.keygen.KeyGenerator#processBefore(org.apache.ibatis.executor.Executor, org.apache.ibatis.mapping.MappedStatement, java.sql.Statement, java.lang.Object)
	 */
	@Override
	public void processBefore(Executor executor, MappedStatement ms, Statement stmt, Object parameter) {
		processGeneratedKeys(executor, ms, parameter);
		
	}
	
	/**
	 *  后置 不做处理
	 */
	@Override
	public void processAfter(Executor executor, MappedStatement ms, Statement stmt, Object parameter) {
		// TODO Auto-generated method stub
		
	}
	
	/**
	 * 前置赋值
	 * @param executor
	 * @param ms
	 * @param parameter
	 */
	private void processGeneratedKeys(Executor executor, MappedStatement ms, Object parameter) {
		if (ms.getSqlCommandType() != SqlCommandType.INSERT) {
			return;
		}
		if (parameter != null && ms.getKeyProperties() != null ) {
			try {
				Configuration configuration = ms.getConfiguration();
				MetaObject metaParam = configuration.newMetaObject(parameter);
				String keyProperty = ms.getKeyProperties()[0];
				// 如果对象可写入主键 则为主键设值
				if(metaParam.hasSetter(keyProperty)) {
					// 如果参数为map直接赋值
					if(metaParam.getObjectWrapper() instanceof MapWrapper) {
						metaParam.setValue(keyProperty, UUIDHelper.uuid());
					}else {
						// 参数为其他类型 检查是否有写入方法 
						try {
							// 字符型主键 设置uuid
							if(String.class == metaParam.getSetterType(keyProperty)) {
								metaParam.setValue(keyProperty, UUIDHelper.uuid());
							}
						} catch (Exception e) {
							// 无法获取
							//e.printStackTrace();
						}
					}
				}
				
				
			} catch (Exception e) {
				throw new ExecutorException("Error selecting key or setting result to parameter object. Cause: " 
						+ e.getLocalizedMessage(), e);
			}
		}
	}
	
}
