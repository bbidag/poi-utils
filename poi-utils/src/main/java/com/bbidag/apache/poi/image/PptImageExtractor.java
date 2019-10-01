package com.bbidag.apache.poi.image;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;

import javax.imageio.ImageIO;

import org.apache.poi.hslf.usermodel.HSLFPictureData;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.util.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class PptImageExtractor implements ExtractImage{
	
	private static Logger logger = LoggerFactory.getLogger(PptImageExtractor.class);
	
	@Override
	public void extractImages(String src, String storePath) {
		// TODO Auto-generated method stub
		
		try {
			//레코드 최대사이즈 
			IOUtils.setByteArrayMaxOverride(Integer.MAX_VALUE);
			
			HSLFSlideShow ppt = new HSLFSlideShow(new FileInputStream(src));
			
			String fileName = "image";
			int idx = 1;
			
			for(HSLFPictureData data : ppt.getPictureData()){
			    byte[] bytes = data.getData();
			    BufferedImage imag=ImageIO.read(new ByteArrayInputStream(bytes));
			    
			    //wdp, emf 파일은 BufferedImage로 이미지 추출 불가능
			    if(imag == null) {
			    	continue;
			    }
			    
			    String extension = data.getType().name().toLowerCase();
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
					
					ImageIO.write(imag, extension, new File(sb.toString()));
				}
				
				idx++;
			}
		} catch(Exception e) {
			logger.error(e.getMessage());
		}
	}
	
}
