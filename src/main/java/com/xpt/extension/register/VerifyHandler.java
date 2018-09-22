package com.xpt.extension.register;

import java.io.Serializable;
import java.util.HashMap;
import java.util.Map;

import com.xpt.extension.exception.ExcelHandleException;

@SuppressWarnings("serial")
public class VerifyHandler implements Serializable {

	private VerifyHandler() {
	}

	private static Map<String, VerifyDataHandler> handlers = new HashMap<>();

	public static VerifyDataHandler getHandler(String key) {
		return handlers.get(key);
	}

	public static void registerHandler(String key, VerifyDataHandler handler) {
		if (handlers.containsKey(key)) {
			throw new ExcelHandleException("Make sure that duplicated key not exists!");
		}
		handlers.put(key, handler);
	}
}