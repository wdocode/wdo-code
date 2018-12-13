package com.wdocode.document.common;

import java.io.Closeable;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.slf4j.Logger;

/**
 * IO相关工具类 如层级创建文件、关闭输入输出流等
 * 
 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
 * @since 2016年9月14日 下午2:12:23
 *
 */
public class IOUtil {
	private static Logger logger = org.slf4j.LoggerFactory.getLogger(IOUtil.class);


	/**
	 * 复制 IO流对象
	 * @param inputStream
	 * @param outputStream
	 * @return
	 * @throws IOException
	 */
	public static long copy(InputStream inputStream, OutputStream outputStream) throws IOException {
		long total = 0;
		byte[] buffer = new byte[1024*10];
		try {
			int len =0;
			while((len = inputStream.read(buffer)) !=-1) {
				total+=len;
				outputStream.write(buffer, 0, len);
			}
			outputStream.flush();
			return total;
		} catch (Exception e) {
			logger.error("",e);
			throw new IOException(e);
		}finally {
			close(inputStream);
		}
	}
	
	
	public static long copy(InputStream inputStream, String filePath) throws IOException {
		if(!createFile(filePath)) {
			throw new IOException("can not create file "+filePath) ;
		}
		OutputStream os = null;
		try {
			os = new FileOutputStream(filePath);
			return copy(inputStream, os);
		} catch (Exception e) {
			logger.error("",e);
			throw new IOException(e);
		} finally {
			close(os);
		}
	}
	
	/**
	 * 将流对象写入到文件 
	 * @param inputStream
	 * @param file
	 * @return
	 * @throws IOException
	 */
	public static long copy(InputStream inputStream, File file) throws IOException {
		if(!createFile(file)) {
			throw new IOException("can not create file "+file) ;
		}
		OutputStream os = null;
		try {
			os = new FileOutputStream(file);
			return copy(inputStream, os);
		} catch (Exception e) {
			logger.error("",e);
			throw new IOException(e);
		} finally {
			close(os);
		}
	}
	
	
	/**
	 * 创建空文件方法 包含创建父目录
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年9月14日 下午2:13:24
	 * @param target_file
	 * @return
	 * @throws IOException
	 */
	public static boolean  createFile(String target_file) throws IOException {
		logger.debug("into createFile(String target_file{})",target_file);
		File file = new File(target_file);
		return createFile(file);
	}
	
	/**
	 * 创建空文件方法 包含创建父目录
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2016年9月14日 下午2:13:24
	 * @param target_file
	 * @return
	 * @throws IOException
	 */
	public static boolean  createFile(File target_file) throws IOException {
		logger.debug("into createFile(String target_file{})",target_file);
		boolean status = true;
		if(!target_file.exists()){
			File parent = target_file.getParentFile();
			if(!parent.exists())
				status = parent.mkdirs();
		}
		if(!status)
			return status;
		logger.debug("end createFile(String target_file{})",target_file);
		return target_file.createNewFile();
		
	}
	
	
	/**
	 * 关闭IO流
	 * @author zhangzx
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since jdk 1.7
	 * @param streams
	 */
	public final static void close(Closeable... streams){  
		for (Closeable s : streams) {  
			try {  
				if (s != null)  
					s.close();  
			} catch (IOException ioe) {
				logger.debug("io object close error",ioe); 
			}  
		}  
	}

}
