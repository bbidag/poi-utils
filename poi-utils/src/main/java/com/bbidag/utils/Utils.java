package com.bbidag.utils;

public class Utils {

	public static boolean isBlank(String str) {
		if(str == null || str.trim().length() == 0) {
			return true;
		} else {
			return false;
		}
	}
	
}
