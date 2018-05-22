package com.wdocode.document.common;

import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.List;

public class StringUtils {
	private static final String CHARSET = "UTF-8";
	
	/**
	 * 根据byte数组长度分割字符串至arrayList
	 * 采用默认的字符集编码 utf-8
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2017年5月19日 下午3:05:50
	 * @param originString
	 * @param byteLen
	 * @return
	 * @throws UnsupportedEncodingException
	 */
	public static List<String> splitStringWhitByte(String originString, int byteLen) 
			throws UnsupportedEncodingException {
		
		return splitStringWhitByte(originString,CHARSET, byteLen);
	}


	/**
	 * 根据byte数组长度分割字符串至arrayList
	 * 
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2017年5月19日 下午3:07:27
	 * @param originString 原字符串
	 * @param charsetName 字符集
	 * @param byteLen 数组长度
	 * @return
	 * @throws UnsupportedEncodingException
	 */
	public static List<String> splitStringWhitByte(String originString, String charsetName, int byteLen) 
			throws UnsupportedEncodingException {
		List<String> list = new ArrayList<String>();
		if (originString == null || originString.isEmpty() || byteLen <= 0) {
            return list;
        }
		int charAtStart =0;
		do{
			int charAt = caclNextCharAtByte(originString,charsetName, charAtStart, byteLen);
			list.add(originString.substring(charAtStart,charAt));
			charAtStart = charAt;
		}while(charAtStart<originString.length());
		
		return list;
	}
	
	/**
	 * 从start 位置开始计算下一个装满byte[byteLen]数组的字符串结束位置
	 * 采用默认字符集 UTF-8
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2017年5月19日 下午3:08:09
	 * @param originString
	 * @param start
	 * @param byteLen
	 * @return
	 * @throws UnsupportedEncodingException
	 */
	public static int caclNextCharAtByte(String originString,int start, int byteLen) 
			throws UnsupportedEncodingException{
		return caclNextCharAtByte(originString,CHARSET, start, byteLen);
	}
	
	/**
	 * 从start 位置开始计算字符串originString中下一个装满byte[byteLen]数组的字符串结束位置
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2017年5月19日 下午3:09:48
	 * @param originString
	 * @param charsetName
	 * @param start
	 * @param byteLen
	 * @return
	 * @throws UnsupportedEncodingException
	 */
	public static int caclNextCharAtByte(String originString,String charsetName,int start, int byteLen) 
			throws UnsupportedEncodingException{
		if (originString == null || originString.isEmpty() || byteLen <= 0) {
            return 0;
        }
		char[] chars = originString.toCharArray();
		
		return caclNextCharAtByte(chars, charsetName,start, byteLen);
	}

	/**
	 * 从start 位置开始计算下一个装满byte[byteLen]数组的字符串结束位置
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2017年5月19日 下午3:10:14
	 * @param chars
	 * @param charsetName
	 * @param start
	 * @param byteLen
	 * @return
	 * @throws UnsupportedEncodingException
	 */
	private static int caclNextCharAtByte(char[] chars,String charsetName, int start, int byteLen) 
			throws UnsupportedEncodingException {
		if(chars == null)
			return 0;
		int length = 0, index = chars.length;
		for (int i = start; i < chars.length; ++i) {
			length += String.valueOf(chars[i]).getBytes(charsetName).length;
			if(length>=byteLen)
				return i-1;
		}
		return index;
	}
	

	
	public static String upperWord(String word){
		char[] cs = word.toLowerCase().toCharArray();
		for(int i=0;i<cs.length-1;i++){
			if(cs[i]=='_'){
				if(i!=0){
					if(cs[i+1]>=48 && cs[i+1]<=57)
						continue;
					cs[i+1] -=32;
				}
			}
		}
		return String.valueOf(cs).replace("_", "");
	}
	
	/**
	 * 字符串首字母转大写
	 * <pre><b>example</b>
	 * upperWordFirst("the__word")=TheWord
	 * upperWordFirst("THE__WORD")=TheWord
	 * upperWordFirst("THE__WORD_2dd")=TheWord2dd
	 * upperWordFirst("theword_2dd")=Theword2dd
	 * </pre>
	 * @author <a href="mailto:zhangzixiao@189.cn">zhang zixiao</a>
	 * @since 2017年5月19日 下午3:04:38
	 * @param word
	 * @return
	 */
	public static String upperWordFirst(String word){
		char[] cs = word.toLowerCase().toCharArray();
		if(cs[0]>='a' && cs[0]<='z')
			cs[0] -= 32;
		for(int i=0;i<cs.length-1;i++){
			if(cs[i]=='_'){
				if(cs[i+1]>='a' && cs[i+1]<='z')
				cs[i+1] -=32;
			}
		}
		return String.valueOf(cs).replace("_", "");
	}
	
}
