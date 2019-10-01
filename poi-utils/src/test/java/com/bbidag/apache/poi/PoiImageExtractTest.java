package com.bbidag.apache.poi;

import com.bbidag.apache.poi.PoiFileReader;

public class PoiImageExtractTest {

	public static void main(String[] args) {
		//전체적으로 그룹핑된 이미지에 대해서는 오피스 외부에서 가져온 이미지에 대해서만 저장이 되는듯함.

		//이미지 순서가 맞는 편임.
		String srcPath = "D:\\data\\1.docx";
//		String srcPath = "D:\\data\\1.doc";

		//이미지 순서가 슬라이드 단위로는 맞는 편임.
//		String srcPath = "D:\\data\\1.pptx";

		//시트별로는 이미지 순서가 맞음.
//		String srcPath = "D:\\data\\1.xlsx";
//		String srcPath = "D:\\data\\1.xls";

		//다른 파일 포멧과 비교했을 때 쓸데없는 텍스트 박스가 많이 추출된다.
//		String srcPath = "D:\\data\\1.ppt";

		String storePath = "D:\\data\\image";

		PoiFileReader reader = new PoiFileReader(srcPath, storePath);
		reader.getImages();
	}

}
