package com.wdocode.mybatis.keygenerator.plugin;

import java.util.Properties;

import org.apache.ibatis.executor.Executor;
import org.apache.ibatis.mapping.MappedStatement;
import org.apache.ibatis.mapping.SqlCommandType;
import org.apache.ibatis.plugin.Interceptor;
import org.apache.ibatis.plugin.Intercepts;
import org.apache.ibatis.plugin.Invocation;
import org.apache.ibatis.plugin.Plugin;
import org.apache.ibatis.plugin.Signature;

import com.wdocode.mybatis.keygenerator.util.ReflectHelper;

@Intercepts( 
		{ @Signature(type = Executor.class, method = "update", args = { MappedStatement.class, Object.class })
		})
public class UUIDKeyGenInterceptor implements Interceptor {
	protected Properties properties;
	protected UUIDKeyGenerator uuidKeyGenerator;
	private String keyGenerator = "keyGenerator";
	
	public UUIDKeyGenInterceptor() {
		// UUIDKeyGenInterceptor0  初始化预处理事件";
		uuidKeyGenerator = new UUIDKeyGenerator();
	}

	@Override
	public Object intercept(Invocation invocation) throws Throwable {
		//1接获到 更新事件 时处理逻辑 
		//2 获取要插入的对象
		//3 判断Commandtype 是insert 则使用反射方式重置keyGenerator
		final MappedStatement mappedStatement = (MappedStatement) invocation.getArgs()[0];
		if (mappedStatement.getSqlCommandType() == SqlCommandType.INSERT) {
			ReflectHelper.setValueByFieldName(mappedStatement, keyGenerator, uuidKeyGenerator);
		}
		return invocation.proceed();
	}
	

	@Override
	public Object plugin(Object target) {
		//插件初始化时调用逻辑;
		return Plugin.wrap(target, this);
	}

	/**
	 *  TODO 添加 KeyGenInterceptor 配置 处理非字符型主键生成方式
	 * @see org.apache.ibatis.plugin.Interceptor#setProperties(java.util.Properties)
	 */
	@Override
	public void setProperties(Properties properties) {
		this.properties = properties;
//		uuidKeyGenerator = new UUIDKeyGenerator();

	}

}
