package com.bbidag.apache.poi.type;

import com.bbidag.utils.Utils;

public enum FileType {
	DOCX,
	DOC,
	PPTX,
	PPT,
	XLSX,
	XLS;
	
	public static FileType getFileType(String path) {
		FileType type = null;
		
		if(Utils.isBlank(path)) {
			return type;
		}
		
		//확장자
		path = path.split("\\.")[path.split("\\.").length-1];
		
		switch(FileType.valueOf(path.toUpperCase())) {
		case DOCX:
			type = DOCX;
			break;
		case DOC:
			type = DOC;
			break;
		case PPTX:
			type = PPTX;
			break;
		case PPT:
			type = PPT;
			break;
		case XLSX:
			type = XLSX;
			break;
		case XLS:
			type = XLS;
			break;
		default:
		}
		
		return type;
	}
}
