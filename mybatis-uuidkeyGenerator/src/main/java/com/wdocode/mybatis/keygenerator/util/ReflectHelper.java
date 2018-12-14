package com.wdocode.mybatis.keygenerator.util;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;



public class ReflectHelper {
	
	private static boolean CACHE = false;
	
	private ReflectHelper() {
	}
	
	/**
	 * return if is use cache.
	 * @return
	 */
	public static boolean isCache() {
		return CACHE;
	}
	
	/**
	 * if true then cache then class fields and method.
	 * @param cache
	 */
	public static void setCache(boolean cache) {
		CACHE = cache;
	}
	
	
	/**
	 * get constructor by class and parameter type names.
	 * @param clazz
	 * @param parameterTypes
	 * @return
	 */
	public static Constructor<?> getConstructor(Class<?> clazz, String... parameterTypes) {
		try {
			if (parameterTypes == null || parameterTypes.length < 1) {
				try {
					return clazz.getConstructor();
				} catch (Exception e) {
					return null;
				}
			}
			Constructor<?>[] constructors = clazz.getConstructors();
			for (Constructor<?> c : constructors) {
				Class<?>[] types = c.getParameterTypes();
				if (types.length != parameterTypes.length) {
					continue;
				}
				int i = 0;
				for (; i < types.length; i++) {
					if (!types[i].getName().equals(parameterTypes[i])) {
						break;
					}
				}
				if (i == types.length) {
					return c;
				}
			}
		} catch (Exception e) {
			throw new RuntimeException("[NoConstructorFound] className:{"+clazz.getName()+"}, parameterTypes:"+parameterTypes, e);
		}
		return null;
	}
	
	/**
	 * get method by class name and method name and parameter type names.
	 * @param className
	 * @param method
	 * @param parameterTypes
	 * @return
	 */
	public static Method getMethod(String className, String method, String... parameterTypes) {
		
		try {
			Class<?> clazz = Class.forName(className);
			return getMethod(clazz, method, parameterTypes);
		} catch (Exception e) {
			throw new RuntimeException("[NoMethodFound] className:{"+className+
					"}, method:{"+method+"}, parameterTypes:{2}"+parameterTypes,e);
		}
	}
	
	/**
	 * get method by class and method name and parameter type names.
	 * @param clazz
	 * @param method
	 * @param parameterTypes
	 * @return
	 */
	public static Method getMethod(Class<?> clazz, String method, String... parameterTypes) {
		try {
			if (clazz == null) {
				return null;
			}
			parameterTypes = parameterTypes == null ? new String[0] : parameterTypes;
			for (Class<?> superClass = clazz; superClass != Object.class; superClass = superClass.getSuperclass()) {
				try {
					Method[] methods = superClass.getDeclaredMethods();
					for (Method m : methods) {
						if (!m.getName().equals(method)) {
							continue;
						}
						Class<?>[] types = m.getParameterTypes();
						if (types.length != parameterTypes.length) {
							continue;
						}
						int i = 0;
						for (; i < types.length; i++) {
							if (!types[i].getName().equals(parameterTypes[i])) {
								break;
							}
						}
						if (i == types.length) {
							m.setAccessible(true);
							return m;
						}
					}
				} catch (Exception e) {
				}
			}
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
		return null;
	}
	
	/**
	 * 获取obj对象fieldName的Field
	 * @param obj
	 * @param fieldName
	 * @return
	 */
	public static Field getFieldByFieldName(Object obj, String fieldName) {
		if (obj == null || fieldName == null) {
			return null;
		}
		for (Class<?> superClass = obj.getClass(); superClass != Object.class; superClass = superClass
				.getSuperclass()) {
			try {
				return superClass.getDeclaredField(fieldName);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return null;
	}
	
	/**
	 * 获取obj对象fieldName的属性值
	 * @param obj
	 * @param fieldName
	 * @return
	 */
	public static Object getValueByFieldName(Object obj, String fieldName) {
		Object value = null;
		try {
			Field field = getFieldByFieldName(obj, fieldName);
			if (field != null) {
				if (field.isAccessible()) {
					value = field.get(obj);
				} else {
					field.setAccessible(true);
					value = field.get(obj);
					field.setAccessible(false);
				}
			}
		} catch (Exception e) {
		}
		return value;
	}
	
	/**
	 * 获取obj对象fieldName的属性值
	 * @param obj
	 * @param fieldName
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public static <T> T getValueByFieldType(Object obj, Class<T> fieldType) {
		Object value = null;
		for (Class<?> superClass = obj.getClass(); superClass != Object.class; superClass = superClass
				.getSuperclass()) {
			try {
				Field[] fields = superClass.getDeclaredFields();
				for (Field f : fields) {
					if (f.getType() == fieldType) {
						if (f.isAccessible()) {
							value = f.get(obj);
							break;
						} else {
							f.setAccessible(true);
							value = f.get(obj);
							f.setAccessible(false);
							break;
						}
					}
				}
				if (value != null) {
					break;
				}
			} catch (Exception e) {
			}
		}
		return (T) value;
	}
	
	/**
	 * 设置obj对象fieldName的属性值
	 * @param obj
	 * @param fieldName
	 * @param value
	 * @throws SecurityException
	 * @throws NoSuchFieldException
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 */
	public static boolean setValueByFieldName(Object obj, String fieldName, Object value) {
		try {
			Field field = obj.getClass().getDeclaredField(fieldName);
			if (field.isAccessible()) {
				field.set(obj, value);
			} else {
				field.setAccessible(true);
				field.set(obj, value);
				field.setAccessible(false);
			}
			return true;
		} catch (Exception e) {
		}
		return false;
	}

	/**
	 * 对象中是否包含fieldName
	 * @param obj
	 * @param fieldName
	 * @return
	 */
	public static boolean hashFieldName(Object obj, String fieldName) {
		
		try {
			Field fields[] = obj.getClass().getDeclaredFields();
			String internedName = fieldName.intern();
			for (int i = 0; i < fields.length; i++) {
				if (fields[i].getName() == internedName) {
					return true;
				}
			}
		} catch (Exception e) {
		}
		return false;
		
	}
	
	
	/**
	 * obj对象fieldName是否可以访问
	 * @param obj
	 * @param fieldName
	 * @return
	 */
	public static boolean isAccessibleField(Object obj, String fieldName) {
		try {
			Field field = obj.getClass().getDeclaredField(fieldName);
			return field.isAccessible();
		} catch (Exception e) {
		}
		return false;
	}
	
}
