package com.example.demo.util;

import java.io.File;

import org.jodconverter.OfficeDocumentConverter;
import org.jodconverter.office.DefaultOfficeManagerBuilder;
import org.jodconverter.office.OfficeException;
import org.jodconverter.office.OfficeManager;

public class ConvertUtil {
	static OfficeManager officeManager;
	
	public static File openOfficeExperience(String sourceFilePath){
		String name="convertFile.pdf";
		File outputFile=new File(name);
		try {
			//开启OpenOffice服务
			OfficeManager manage=getOfficeManager();
			File sourceFile = new File(sourceFilePath);
			//设置转换后的文件存储路径，文件名
			
			//使用OfficeDocumentConverter类转换文件
			OfficeDocumentConverter converter=new OfficeDocumentConverter(manage);
			converter.convert(sourceFile,outputFile);
		}catch(Exception e) {
			e.printStackTrace();
		}finally {
			if(officeManager != null){
	               try {
	                   officeManager.stop();
	               } catch (OfficeException e) {
	                   e.printStackTrace();
	               }
	       }
		}
		return outputFile;
	}
	/**
	 * 
	 * @return 返回一个OfficeManager实例，用于处理转换业务
	 * @throws OfficeException 
	 */
	
	private static OfficeManager getOfficeManager() throws OfficeException {
		DefaultOfficeManagerBuilder builder=new DefaultOfficeManagerBuilder();
		//此处入参可以填写OpenOffice安装路径，本例子中，openOffice安装在E盘
		builder.setOfficeHome("E:/DevelopTools/openOffice/workspace");
		OfficeManager officeManager =builder.build();
		//officeManager提供了开启OpenOffice的API服务
		officeManager.start();
		return officeManager;
	}
	
	/**
	 * 简单测试一下，嗷
	 * @param args
	 */
	public static void main(String[] args) {
		//把我D盘下的一个图片转为pdf
		ConvertUtil.openOfficeExperience("D:\\image\\20190323\\1553312645151083380.png");
	}
}
