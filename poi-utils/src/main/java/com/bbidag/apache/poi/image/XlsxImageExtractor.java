package com.bbidag.apache.poi.image;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class XlsxImageExtractor implements ExtractImage {
	
	private static Logger logger = LoggerFactory.getLogger(XlsxImageExtractor.class);

	@Override
	public void extractImages(String src, String storePath) {
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(src);

			List<XSSFPictureData> pictures = workbook.getAllPictures();
			
			int idx = 1;

			String fileName = "image";
			
			for(XSSFPictureData pic : pictures) {
				byte[] bytes = pic.getData();
				String extension = pic.suggestFileExtension();
				BufferedImage imag=ImageIO.read(new ByteArrayInputStream(bytes));
				
				//wdp, emf 파일은 BufferedImage로 이미지 추출 불가능
				if(imag == null) {
					continue;
				}
				
				StringBuilder sb = new StringBuilder();
				
				if(storePath.endsWith("/") || storePath.endsWith("\\")) {
					sb.append(storePath)
					  .append(fileName)
					  .append(idx)
					  .append(".")
					  .append(extension);
					
					ImageIO.write(imag, extension, new File(sb.toString()));
				} else {
					sb.append(storePath)
					  .append(File.separator)
					  .append(fileName)
					  .append(idx)
					  .append(".")
					  .append(extension);
					
					ImageIO.write(imag, extension, new File(storePath+File.separator+fileName+idx+"."+extension));
				}
				
				idx++;
			}
		} catch(Exception e) {
			logger.error(e.getMessage());
		}
	}

}