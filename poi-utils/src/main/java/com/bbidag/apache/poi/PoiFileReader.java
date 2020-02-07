package com.bbidag.apache.poi;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.bbidag.apache.poi.image.DocImageExtractor;
import com.bbidag.apache.poi.image.DocxImageExtractor;
import com.bbidag.apache.poi.image.ExtractImage;
import com.bbidag.apache.poi.image.PptImageExtractor;
import com.bbidag.apache.poi.image.PptxImageExtractor;
import com.bbidag.apache.poi.image.XlsImageExtractor;
import com.bbidag.apache.poi.image.XlsxImageExtractor;
import com.bbidag.apache.poi.type.FileType;

public class PoiFileReader {
	
	private static final Logger logger = LoggerFactory.getLogger(PoiFileReader.class);
	
	private String srcPath;
	private String storePath;
	private ExtractImage extractImage;
	
	public PoiFileReader(String srcPath, String storePath) {
		FileType type = FileType.getFileType(srcPath);
		
		if(type == FileType.DOCX) {
			this.extractImage = new DocxImageExtractor();
		} else if(type == FileType.PPTX) {
			this.extractImage = new PptxImageExtractor();
		} else if(type == FileType.XLSX){
			this.extractImage = new XlsxImageExtractor();
		} else if(type == FileType.DOC){
			this.extractImage = new DocImageExtractor();
		} else if(type == FileType.PPT){
			this.extractImage = new PptImageExtractor();
		} else if(type == FileType.XLS) {
			this.extractImage = new XlsImageExtractor();
		} else {
			logger.error("Not Found Extension.");
			throw new NullPointerException();
		}
		
		this.srcPath = srcPath;
		this.storePath = storePath;
	}
	
	public void getImages() {
		this.extractImage.extractImages(this.srcPath, this.storePath);
	}
	
	public static void main(String[] args) {
		//전체적으로 그룹핑된 이미지에 대해서는 오피스 외부에서 가져온 이미지에 대해서만 저장이 되는듯함.
		
		//이미지 순서가 맞는 편임.
		String srcPath = "D:\\data\\test.docx";

		//이미지 순서가 슬라이드 단위로는 맞는 편임.
//		String srcPath = "D:\\data\\test.pptx";

		//시트별로는 이미지 순서가 맞음.
//		String srcPath = "D:\\data\\test.xlsx";
//		String srcPath = "D:\\data\\test.xls";
		
		//다른 파일 포멧과 비교했을 때 쓸데없는 텍스트 박스가 많이 추출된다.
//		String srcPath = "D:\\data\\test.ppt";
		
		String storePath = "D:\\data\\poi\\image";
		
		PoiFileReader reader = new PoiFileReader(srcPath, storePath);
		reader.getImages();
	}
	
}
