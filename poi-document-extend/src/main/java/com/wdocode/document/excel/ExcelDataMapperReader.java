package com.wdocode.document.excel;

import java.io.File;
import java.util.List;

import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import com.wdocode.document.excel.model.ExcelDataMapper;

/**
 * excel数据映射配置读取工具类
 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
 * @since 2016年4月5日 下午5:34:44
 *
 */
public class ExcelDataMapperReader {
	/**配置文件根节点 */
	public static final String root_name = "mapper";
	private Boolean realPath = false;
	
	private String xmlmapper_basefolder="excel_mapper/";
	
//	<export_data_mapper>
//    <!-- excel -mapper 样例 -->
//    <mapper id='exportload_sample' file_name="导出样例">
//    </mapper>
//</export_data_mapper>
	
	/**
	 * 创建mapper文件
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param basePath
	 * @return
	 */
	public static ExcelDataMapperReader build(String basePath){
		ExcelDataMapperReader b = new ExcelDataMapperReader(basePath);
		return b;
	}
	
	/**
	 * 创建mapper文件
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param basePath
	 * @return
	 */
	public static ExcelDataMapperReader build(){		
		ExcelDataMapperReader b = new ExcelDataMapperReader(null);
		return b;
	}
	
	protected ExcelDataMapperReader(String basePath){
		realPath = basePath!=null;
		this.xmlmapper_basefolder = basePath == null?xmlmapper_basefolder:basePath;
	}
	
	protected ExcelDataMapperReader(){
		
	}
	
	/**
	 * 获取导出excel数据映射对象
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年4月5日 下午5:34:44
	 * @param dataMapperId 数据映射文件中对应mapperid
	 * @param mapperFileName 数据映射说明文件
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public ExcelDataMapper getDataMapper(String dataMapperId,String mapperFileName) {
		Element root = getXmlRootElement(mapperFileName);
		if(root == null)
			return null;
		List<Element> list = root.elements(root_name);
		if(list == null)
			return null;
		ExcelDataMapper mapper = null;
		for(Element e:list){
			if(!dataMapperId.equals(e.attributeValue(ExcelDataMapper.attr_dataMapperId)))
				continue;
			mapper = ExcelDataMapper.readFromElement(getXMLBasePath(),e);
			break;
		}
		return mapper;
	}
	
	
	/**
	 * 获取xml 的根节点
	 * @param filePath
	 * @return
	 */
	private Element getXmlRootElement(String mapperName) {
		try {
			SAXReader r=new SAXReader();  
			File docFile = new File(getXMLBasePath()+mapperName);
			Document doc = r.read(docFile);
			return doc.getRootElement();
		} catch (Exception e) {
			throw new RuntimeException("read xml error ",e);
		}
	}
	
	/**
	 * 获取配置文件根目录
	 *  @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 *  @since 2016年4月5日 下午5:34:44
	 * @return
	 */
	public String getXMLBasePath(){
		if(realPath)
			return xmlmapper_basefolder;
		
		return ExcelDataMapperReader.class.getClassLoader()
				.getResource(xmlmapper_basefolder).getPath();
	}
	

}
