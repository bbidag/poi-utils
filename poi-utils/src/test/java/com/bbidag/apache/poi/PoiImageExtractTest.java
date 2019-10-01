package com.bbidag.apache.poi;

import com.bbidag.apache.poi.PoiFileReader;

public class PoiImageExtractTest {

	public static void main(String[] args) {
		//전체적으로 그룹핑된 이미지에 대해서는 오피스 외부에서 가져온 이미지에 대해서만 저장이 되는듯함.

		//이미지 순서가 맞는 편임.
		//				String srcPath = "D:\\data\\poi\\sample\\2019_0805-수집_프로젝트_개발자_인수인계서.docx";
		//				String srcPath = "D:\\data\\poi\\sample\\2019_0805-수집_프로젝트_개발자_인수인계서.doc";
		String srcPath = "D:\\data\\poi\\sample\\2019_0704-LG Platform Service 사용자 매뉴얼.docx";
		//				String srcPath = "D:\\data\\poi\\sample\\2019_0704-LG Platform Service 사용자 매뉴얼.doc";

		//이미지 순서가 슬라이드 단위로는 맞는 편임.
		//				String srcPath = "D:\\data\\poi\\sample\\2019_0628-중간산출물-솔트룩스.pptx";

		//시트별로는 이미지 순서가 맞음.
		//				String srcPath = "D:\\data\\poi\\sample\\레인보우 개편 기획서 ver 1.1.xlsx";
		//				String srcPath = "D:\\data\\poi\\sample\\레인보우 개편 기획서 ver 1.1.xls";

		//다른 파일 포멧과 비교했을 때 쓸데없는 텍스트 박스가 많이 추출된다.
		//				String srcPath = "D:\\data\\poi\\sample\\2019_0628-중간산출물-솔트룩스.ppt";

		String storePath = "D:\\data\\poi\\image";

		PoiFileReader reader = new PoiFileReader(srcPath, storePath);
		reader.getImages();
	}

}
