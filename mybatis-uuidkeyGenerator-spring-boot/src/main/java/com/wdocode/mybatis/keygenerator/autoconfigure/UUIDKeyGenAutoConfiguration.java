package com.wdocode.mybatis.keygenerator.autoconfigure;

import java.util.List;
import java.util.Properties;

import javax.annotation.PostConstruct;

import org.apache.ibatis.session.SqlSessionFactory;
import org.mybatis.spring.boot.autoconfigure.MybatisAutoConfiguration;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.autoconfigure.AutoConfigureAfter;
import org.springframework.boot.autoconfigure.condition.ConditionalOnProperty;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import com.wdocode.mybatis.keygenerator.plugin.UUIDKeyGenInterceptor;

@Configuration
@ConditionalOnProperty(prefix = "keygenerator",value = {"enable"},havingValue = "true")
//@EnableConfigurationProperties
@AutoConfigureAfter(MybatisAutoConfiguration.class)
public class UUIDKeyGenAutoConfiguration {
	
	@Autowired
	private List<SqlSessionFactory> sqlSessionFactoryList;
	
	 @Bean
	 @ConfigurationProperties(prefix = "keygenerator")
	 public Properties wdocodeUUIDPluginProperties() {
//		 keygenerator: 
//			  enable: true
//			  number: 
//			    generatorType: sequence
//			    generatorSql: select 
//			  string:
//			    generatorType: uuid
		 return new Properties();
	 }
	
	@PostConstruct
    public void addKeyGenInterceptor() {
		UUIDKeyGenInterceptor interceptor = new UUIDKeyGenInterceptor();
        interceptor.setProperties(wdocodeUUIDPluginProperties());
        for (SqlSessionFactory sqlSessionFactory : sqlSessionFactoryList) {
            sqlSessionFactory.getConfiguration().addInterceptor(interceptor);
        }
    }

}
