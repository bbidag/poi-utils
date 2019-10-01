package com.bbidag.apache.poi.image;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.Picture;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class DocImageExtractor implements ExtractImage {
	
	private static Logger logger = LoggerFactory.getLogger(DocImageExtractor.class);
	
	@Override
	public void extractImages(String src, String storePath) {
		// TODO Auto-generated method stub
		try{
			FileInputStream fs=new FileInputStream(src);
			HWPFDocument doc=new HWPFDocument(fs);
			PicturesTable table=doc.getPicturesTable();
			List<Picture> pictures = table.getAllPictures();
			
			Iterator<Picture> iterator=pictures.iterator();
			
			String imageFileName = "image";
			int idx = 1;
			
			while(iterator.hasNext()){
				Picture pic=iterator.next();
				byte[] bytepic=pic.getRawContent();
				BufferedImage imag=ImageIO.read(new ByteArrayInputStream(bytepic));
				
				//wdp, emf 파일은 BufferedImage로 이미지 추출 불가능
				if(imag == null) {
					continue;
				}
				
				String extension = pic.suggestFileExtension();
				
				StringBuilder sb = new StringBuilder();
				
				if(storePath.endsWith("/") || storePath.endsWith("\\")) {
					sb.append(storePath)
					  .append(imageFileName)
					  .append(idx)
					  .append(".")
					  .append(extension);
					
					ImageIO.write(imag, extension, new File(sb.toString()));
				} else {
					sb.append(storePath)
					  .append(File.separator)
					  .append(imageFileName)
					  .append(idx)
					  .append(".")
					  .append(extension);
					
					ImageIO.write(imag, extension, new File(sb.toString()));
				}
				
				idx++;
			}
		}catch(Exception e){
			logger.error(e.getMessage());
		}
	}
	
}
