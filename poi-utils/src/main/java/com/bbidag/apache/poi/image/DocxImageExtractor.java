package com.bbidag.apache.poi.image;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class DocxImageExtractor implements ExtractImage {

	private static Logger logger = LoggerFactory.getLogger(DocxImageExtractor.class);
	
	@Override
	public void extractImages(String src, String storePath) {
		try{
			//create file inputstream to read from a binary file
			FileInputStream fs=new FileInputStream(src);
			//create office word 2007+ document object to wrap the word file
			XWPFDocument docx=new XWPFDocument(fs);
			//get all images from the document and store them in the list piclist
			List<XWPFPictureData> piclist=docx.getAllPictures();
			//traverse through the list and write each image to a file
			Iterator<XWPFPictureData> iterator=piclist.iterator();
			while(iterator.hasNext()){
				XWPFPictureData pic=iterator.next();
				byte[] bytepic=pic.getData();
				BufferedImage imag=ImageIO.read(new ByteArrayInputStream(bytepic));
				
				//wdp, emf 파일은 BufferedImage로 이미지 추출 불가능
				if(imag == null) {
					continue;
				}
				
				String imageFileName = pic.getFileName();
				
				String extension =imageFileName.split("\\.")[imageFileName.split("\\.").length-1];
				
				StringBuilder sb = new StringBuilder();
				
				if(storePath.endsWith("/") || storePath.endsWith("\\")) {
					sb.append(storePath)
					  .append(imageFileName);
					
					ImageIO.write(imag, extension, new File(sb.toString()));
				} else {
					sb.append(storePath)
					  .append(File.separator)
					  .append(imageFileName);
					
					ImageIO.write(imag, extension, new File(sb.toString()));
				}
			}
		}catch(Exception e){
			logger.error(e.getMessage());
		}
	}
	
	
}
