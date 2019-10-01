package com.bbidag.apache.poi.image;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;

import javax.imageio.ImageIO;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class PptxImageExtractor implements ExtractImage {

	private static Logger logger = LoggerFactory.getLogger(PptxImageExtractor.class);
	
	@Override
	public void extractImages(String src, String storePath) {
		// TODO Auto-generated method stub
		
		try {
			XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(src));
			for(XSLFPictureData data : ppt.getPictureData()){
			    byte[] bytes = data.getData();
			    String fileName = data.getFileName();
			    BufferedImage imag=ImageIO.read(new ByteArrayInputStream(bytes));
			    
			    //wdp, emf 파일은 BufferedImage로 이미지 추출 불가능
			    if(imag == null) {
			    	continue;
			    }
			    
			    String extension =fileName.split("\\.")[fileName.split("\\.").length-1];
			    StringBuilder sb = new StringBuilder();
			    
				if(storePath.endsWith("/") || storePath.endsWith("\\")) {
					sb.append(storePath)
					  .append(fileName);
					
					ImageIO.write(imag, extension, new File(storePath+fileName));
				} else {
					sb.append(storePath)
					  .append(File.separator)
					  .append(fileName);
					
					ImageIO.write(imag, extension, new File(storePath+File.separator+fileName));
				}
			}
		} catch(Exception e) {
			logger.error(e.getMessage());
		}
	}
	
}
