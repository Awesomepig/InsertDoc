package com.asomepig.Jacob;

import java.io.File;
import java.io.IOException;
import java.util.Map;
import java.util.Scanner;

import jxl.Sheet;
import jxl.Workbook;

import com.asomepig.image.ImageUtil;
import com.asomepig.jpedal.JpedalUtil;
import com.asomepig.jxl.JxlTools;
import com.asomepig.log.SimpleLogService;
import com.asomepig.log.SimpleLogServiceImpl;
import com.asomepig.util.FileUtil;
import com.asomepig.util.StringUtil;

/**
 * @author Eric Chen asomepig@gmail.com 
 *
 */
public class RunJcob {

	public static void main(String[] args) throws IOException {

		RunJcob r = new RunJcob();
		r.startIt();
		
	}
	public void startIt(){
		boolean ifVersion2 = false;
		ifVersion2 = this.getUserChoose();
		String sp = File.separator;
		String rootPath = System.getProperty("user.dir");
		String pics = rootPath+sp+"pic"+sp+"pics"+sp;
		String pics2 = rootPath+sp+"pic"+sp+"pics2"+sp;
		String pdfs = rootPath+sp+"pic"+sp+"pdf"+sp;
		String exl = rootPath+sp+"pic"+sp+"excel"+sp;
		// 获取图片和pdf文件夹,比较列表
		try {
			File pcps = new File(pics);
			File pcps2 = new File(pics2);
			File pfps = new File(pdfs);
			File exls = new File(exl);
			// 1.程序先决条件判断
			if(!pcps.isDirectory() || !exls.isDirectory())
			{
				System.out.println("图片文件夹   OR pdf文件夹   OR excel文件夹未准备好,程序终止!");
				RunJcob.sleep(3,"程序执行失败");
				return ;
			}
			File[] picArr = pcps.listFiles();
			File[] picArr2 = null;
			File[] pdfArr = null;
			File[] exlArr = exls.listFiles();
		
		/////////////////////////////////////=============两图版本的判断需求，单图版本不需要============////////////////////////////////////////
				if(!ifVersion2)
				{
					if(!pcps2.isDirectory())
						
					{
						System.out.println("pdf文件夹 未准备好,程序终止!");
						RunJcob.sleep(3,"程序执行失败");
						return ;
					}
					if(!pfps.isDirectory())
	
					{
						System.out.println("pdf文件夹 未准备好,程序终止!");
						RunJcob.sleep(3,"程序执行失败");
						return ;
					}
					picArr2 = pcps2.listFiles();
					pdfArr = pfps.listFiles();
					if(picArr.length!= pdfArr.length || pdfArr.length != picArr2.length)
					{
						System.out.println("图片数量与pdf数量不匹配,程序终止!");
						System.out.println("图片数pic目录:"+picArr.length+"图片数pics目录:"+picArr2.length+"  ; pdf数:"+pdfArr.length);
						RunJcob.sleep(5,"程序执行失败");
						return ;
					}else{
						for (int i = 0; i < picArr.length; i++) {
							String picIname = picArr[i].getName();
							picIname = picIname.substring(0, picIname.lastIndexOf("."));
							String pdfIname = pdfArr[i].getName();
							pdfIname = pdfIname.substring(0, pdfIname.lastIndexOf("."));
							if(!pdfIname.equalsIgnoreCase(picIname))
							{
								RunJcob.sleep(5,picIname+".gif <-----> "+pdfIname+".pdf :"+"DotNo不匹配 ");
								return ;
							}
						}
					}
	
				}
			/////////////////////////////////////=============两图版本的判断需求，单图版本不需要============////////////////////////////////////////
			
			if(exlArr.length<1 || (!exlArr[0].getName().endsWith(".xls") && !exlArr[0].getName().endsWith(".xlsx")))
			{
				RunJcob.sleep(5,"excel文件不存在\n程序执行失败");
				return ;
			}
			// 2.开始循环插入
			
				//2.1 获取excel对象
				File exlFile = exlArr[0];
				Workbook book = Workbook.getWorkbook(exlFile);
				Sheet st = book.getSheet(0);
				JxlTools jxl = new JxlTools();
			if(!ifVersion2)
			{
				// 版本1的执行方式，
				pdfCycle:for (int i = 0; i < pdfArr.length; i++)
				{
					String pdfName = pdfArr[i].getName();
					String dotno = pdfName.substring(0, pdfName.lastIndexOf("."));
					String picName = "";
					String picName2 = "";
					////////////////////////////////////////////////
					System.out.println(pdfArr[i].getName());
					////////////////////////////////////////////////
					
					picCycle:for (int j = 0; j < picArr.length; j++) {
						String picname = picArr[j].getName();
						if(StringUtil.startWithIgnoreCase(picname,dotno))
						{
							picName = picname;
							break picCycle;
						}
					}
					if(picName.equals("")){
						System.err.println(dotno+"的图片pics目录下不存在!");
						continue pdfCycle;
					}
					picCycle2:for (int j = 0; j < picArr2.length; j++) {
						String picname = picArr2[j].getName();
						if(StringUtil.startWithIgnoreCase(picname,dotno))
						{
							picName2 = picname;
							break picCycle2;
						}
					}
					if(picName2.equals("")){
						System.err.println(dotno+"的图片pics2目录下不存在!");
						continue pdfCycle;
					}
					Map<String,String> bookmarkResource = jxl.getBookMarkResource(st, dotno.toUpperCase(), pdfArr.length,false);
					this.convertIt(picName,picName2,pdfName,bookmarkResource );
				}
				
			}else//版本2的执行方式
			{
				gifCycle:for (int i = 0; i < picArr.length; i++)
				{
					String picName = picArr[i].getName();
					String dotno = picName.substring(0, picName.lastIndexOf("."));
					Map<String,String> bookmarkResource = jxl.getBookMarkResource(st, dotno.toUpperCase(), picArr.length,true);
					this.convertIt2(picName,bookmarkResource );
				}
			}
				
				RunJcob.sleep(5,"程序执行成功");
			
		} catch (Exception e) {
			System.err.println("文件出错!");
			e.printStackTrace();
		}
	}
	
	/**
	 * @return true版本2，false版本1
	 */
	private boolean getUserChoose() {
		boolean res = false;
		Scanner s = new Scanner(System.in); 
        System.out.print("请选择生成文档的版本（1.两图的版本，2.一张图版本.）\n（1或者2）默认2："); 
        while (true) { 
                String line = s.nextLine(); 
                if (line.equals("2") || line.equals(""))
                {
                	res = true;
                	System.out.println("您已选择版本>>>2"); 
                	break; 
                }else
                {
                	System.out.println("您已选择版本>>>1"); 
                	break; 
                }
        } 
		return res;
	}
	public void convertIt(String picName,String picName2,String pdfName,Map<String,String> bookMarkResource){
		String sp = File.separator;
		SimpleLogService log = new SimpleLogServiceImpl();
		String rootPath = System.getProperty("user.dir");
		String pic = rootPath+sp+"pic"+sp;
		String target = rootPath+sp+"target"+sp;
		String source = rootPath+sp+"source"+sp;
		
		String pdfNameWithoutSub = pdfName.substring(0, pdfName.lastIndexOf("."));
		
		String pdfPath = pic+"pdf"+sp+pdfName;
		String image1 = pic+"pics"+sp+picName;
		String image3 = pic+"pics2"+sp+picName2;
		String image2 = pic+"pdf2pics"+sp+pdfNameWithoutSub+".png";
		String imageTarget1 = target+picName;
		String imageTarget3 = target+picName2;
		String imageTarget2 = target+pdfNameWithoutSub+".png";
		String wordFile = source+"3.doc";
		String resFilePath = target+pdfNameWithoutSub+".doc";
		if(!FileUtil.ifFileExists(image1))return ;
		if(!FileUtil.ifFileExists(image3))return ;
		if(!FileUtil.ifFileExists(pdfPath))return ;
		if(!FileUtil.ifFileExists(wordFile))return ;
		
		try {
			//
			//------------------------------- CONVERT PDF TO PNG
			System.out.println("|-------开始转换pdf-----\n");
			log.append("|-------开始转换pdf-----\n");
			JpedalUtil ju = new JpedalUtil();
			ju.pdf2png(pdfPath, image2);
			System.out.println("|-------pdf转换完成-----\n");
			log.append("|-------pdf转换完成-----\n");
			
			//------------------------------- ZOOM PICS
			System.out.println("|-------开始缩放图片-----\n");
			log.append("|-------开始缩放图片-----\n");
			ImageUtil.resize(new File(image1), new File(imageTarget1), 310, 1);
			ImageUtil.compressImage(new File(image2), imageTarget2, 2343, 3113, false);
			System.out.println("|-------缩放图片完成-----\n");
			log.append("|-------缩放图片完成-----\n");
			// ------------------------------ COPY DOC MODEL 2 RES FOLDER
			System.out.println("|-------开始准备文档-----");
			log.append("|-------开始准备文档-----");
//			FileUtil.copyFile(image1, imageTarget1);
			FileUtil.copyFile(image3, imageTarget3);
			FileUtil.copyFile(wordFile, resFilePath);
			
			// ------------------------------ INSERT PICS 2  BOOKMARKS 
			System.out.println("|-------开始处理文档-----\n");
			log.append("|-------开始处理文档-----\n");
			JacobWordInsert poi = new JacobWordInsert(resFilePath);
			poi.addImageAtBookMark("tp1", imageTarget1);
			log.append("PIC 1 SUCCESS!");
			
			poi.addImageAtBookMark("tp3", imageTarget3);
			log.append("PIC 3 SUCCESS!");
			
			poi.addImageAtBookMark("tp2", imageTarget2);
			log.append("PIC 2 SUCCESS!");
			
			//poi.addTextAtBookMark("wz1", "insert by java!");
			// ------------------------------ INSERT EXCEL CONTENTS
			for(String bookmark:bookMarkResource.keySet())
			{
				poi.addTextAtBookMark(bookmark, bookMarkResource.get(bookmark));
			}
			log.append("insert excel SUCCESS!");
			
			
			
			// ------------------------------ delete files
//			FileUtil.deleteFile(pdfPath);
//			FileUtil.deleteFile(image1);
			FileUtil.deleteFile(image2);
			FileUtil.deleteFile(imageTarget1);
			FileUtil.deleteFile(imageTarget2);
			FileUtil.deleteFile(imageTarget3);
			poi.closeDocument();
			poi.closeWord();
			
		} catch (Exception e) {
			log.append(e.getMessage());
		}finally{
			log.close();
		}
	}
	
	
	public void convertIt2(String picName,Map<String,String> bookMarkResource){
		String sp = File.separator;
		SimpleLogService log = new SimpleLogServiceImpl();
		String rootPath = System.getProperty("user.dir");
		String pic = rootPath+sp+"pic"+sp;
		String target = rootPath+sp+"target"+sp;
		String source = rootPath+sp+"source"+sp;
		
		String pdfNameWithoutSub = bookMarkResource.get("A10");
		
		String image1 = pic+"pics"+sp+picName;
		String imageTarget1 = target+picName;
		String wordFile = source+"2.doc";
		String resFilePath = target+pdfNameWithoutSub+".doc";
		if(!FileUtil.ifFileExists(image1))return ;
		if(!FileUtil.ifFileExists(wordFile))return ;
		
		try {
			//
			// ------------------------------ COPY DOC MODEL 2 RES FOLDER
			System.out.println("|-------开始准备文档-----");
			System.out.println("|-------开始缩放图片-----");
//			ImageUtil.compressImage(new File(image1), imageTarget1, 540, 540, false);
			ImageUtil.resize(new File(image1), new File(imageTarget1), 540, 1);
			log.append("|-------开始准备文档-----");
//			FileUtil.copyFile(image1, imageTarget1);
			FileUtil.copyFile(wordFile, resFilePath);
			
			// ------------------------------ INSERT PICS 2  BOOKMARKS 
			System.out.println("|-------开始处理文档-----\n");
			log.append("|-------开始处理文档-----\n");
			JacobWordInsert poi = new JacobWordInsert(resFilePath);
			poi.addImageAtBookMark("tp1", imageTarget1);
			log.append("PIC 1 SUCCESS!");
			// ------------------------------ INSERT EXCEL CONTENTS
			for(String bookmark:bookMarkResource.keySet())
			{
				poi.addTextAtBookMark(bookmark, bookMarkResource.get(bookmark));
			}
			log.append("insert excel SUCCESS!");
			
			
			
			// ------------------------------ delete files
			FileUtil.deleteFile(imageTarget1);
			poi.closeDocument();
			poi.closeWord();
			
		} catch (Exception e) {
			log.append(e.getMessage());
		}finally{
			log.close();
		}
	}
	
	public static void sleep(int tim,String note){
		System.out.println("\n\n\n\n"+note+",程序将在 "+tim+"s后关闭!");
		try {
				for (int i = tim; i > 0; i--)
				{
					System.out.println(i);	
					Thread.sleep(1000);
				}
				System.out.println("程序结束！");
			} catch (InterruptedException e) {
				e.printStackTrace();
		}
	}
}
