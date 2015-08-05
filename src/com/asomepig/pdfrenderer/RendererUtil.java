package com.asomepig.pdfrenderer;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.Rectangle;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.RandomAccessFile;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;

import javax.imageio.ImageIO;

import com.sun.image.codec.jpeg.JPEGCodec;
import com.sun.image.codec.jpeg.JPEGImageEncoder;
import com.sun.pdfview.PDFFile;
import com.sun.pdfview.PDFPage;
import com.sun.pdfview.PDFRenderer;


public class RendererUtil {

	public RendererUtil() {
		// TODO Auto-generated constructor stub
	}

	public static void main(String[] args) {
		RendererUtil rd = new RendererUtil();
		String rootPath = System.getProperty("user.dir");
		String pdfPath = rootPath+"\\pics\\1.pdf";
		String pngPath = rootPath+"\\pics\\B.png";
//		rd.Pdf_Png(1, pdfPath, pngPath);
		try {
			rd.setup(pdfPath,pngPath);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public void Pdf_Png(int pageNumber,String pdfPath,String pngPath) {
		int pagen = pageNumber;
		File file = new File(pdfPath);

		PDFFile pdffile = null;
		// set up the PDF reading
		try {
			RandomAccessFile raf = new RandomAccessFile(file, "r");
			FileChannel channel = raf.getChannel();
			ByteBuffer buf = channel.map(FileChannel.MapMode.READ_ONLY, 0,
					channel.size());
			pdffile = new PDFFile(buf);
	
			if (pagen < pdffile.getNumPages())
			{
				raf.close();
				return;
			}
			// print出该pdf文档的页数
			System.out.println("文档页数为:"+pdffile.getNumPages());
	
			// 设置将第pagen也生成png图片
			PDFPage page = pdffile.getPage(pagen);
	
			// create and configure a graphics object
			int width = (int) page.getBBox().getWidth();
			int height = (int) page.getBBox().getHeight();
			BufferedImage img = new BufferedImage(width, height,BufferedImage.TYPE_INT_ARGB);
			Graphics2D g2 = img.createGraphics();
			g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING,RenderingHints.VALUE_ANTIALIAS_ON);
	
			// do the actual drawing
			PDFRenderer renderer = new PDFRenderer(
					page, 
					g2, 
					new Rectangle(0, 0,width, height),
					null,
					Color.WHITE);

			page.waitForFinish();
			renderer.run();
			g2.dispose();
			ImageIO.write(img, "png", new File(pngPath));
			raf.close();
			
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}
	public  void setup(String pdfPath,String pngPath) throws Exception {
        File file = new File(pdfPath);
        RandomAccessFile raf = new RandomAccessFile(file, "r");
        FileChannel channel = raf.getChannel();
        ByteBuffer buf = channel.map(FileChannel.MapMode.READ_ONLY, 0, channel.size());
        PDFFile pdffile = new PDFFile(buf);
        System.out.println("页数： " + pdffile.getNumPages());

        for (int i = 1; i <= pdffile.getNumPages(); i++) {
            // draw the first page to an image
            PDFPage page = pdffile.getPage(i);

            // get the width and height for the doc at the default zoom
            Rectangle rect = new Rectangle(0, 0, (int) page.getBBox()
                    .getWidth(), (int) page.getBBox().getHeight());

            // generate the image
            Image img = page.getImage(rect.width, rect.height, // width &
                                                                // height
                    rect, // clip rect
                    null, // null for the ImageObserver
                    true, // fill background with white
                    true // block until drawing is done
                    );

            BufferedImage tag = new BufferedImage(rect.width, rect.height,
                    BufferedImage.TYPE_INT_RGB);
            tag.getGraphics().drawImage(img, 0, 0, rect.width, rect.height,
                    null);
            FileOutputStream out = new FileOutputStream(pngPath); // 输出到文件流
            JPEGImageEncoder encoder = JPEGCodec.createJPEGEncoder(out);
            encoder.encode(tag); // JPEG编码

            out.close();
            
        }
        pdffile=null;
        buf.clear();   
        raf.close();
//        file.delete();     
            
   } 

}
