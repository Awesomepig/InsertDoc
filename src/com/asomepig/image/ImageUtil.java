package com.asomepig.image;

import java.awt.Color;
import java.awt.Graphics;
import java.awt.Image;
import java.awt.image.BufferedImage;
import java.awt.image.ConvolveOp;
import java.awt.image.Kernel;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;
import javax.swing.ImageIcon;

import com.sun.image.codec.jpeg.JPEGCodec;
import com.sun.image.codec.jpeg.JPEGEncodeParam;
import com.sun.image.codec.jpeg.JPEGImageEncoder;

public class ImageUtil {

	public static void resize(File originalFile, File resizedFile,
			int newWidth, float quality) throws IOException {

		if (quality > 1) {
			throw new IllegalArgumentException(
					"Quality has to be between 0 and 1");
		}

		ImageIcon ii = new ImageIcon(originalFile.getCanonicalPath());
		Image i = ii.getImage();
		Image resizedImage = null;

		int iWidth = i.getWidth(null);
		int iHeight = i.getHeight(null);

		if (iWidth > iHeight) {
			resizedImage = i.getScaledInstance(newWidth, (newWidth * iHeight)
					/ iWidth, Image.SCALE_SMOOTH);
		} else {
			resizedImage = i.getScaledInstance((newWidth * iWidth) / iHeight,
					newWidth, Image.SCALE_SMOOTH);
		}

		// This code ensures that all the pixels in the image are loaded.
		Image temp = new ImageIcon(resizedImage).getImage();

		// Create the buffered image.
		BufferedImage bufferedImage = new BufferedImage(temp.getWidth(null),
				temp.getHeight(null), BufferedImage.TYPE_INT_RGB);

		// Copy image to buffered image.
		Graphics g = bufferedImage.createGraphics();

		// Clear background and paint the image.
		g.setColor(Color.white);
		g.fillRect(0, 0, temp.getWidth(null), temp.getHeight(null));
		g.drawImage(temp, 0, 0, null);
		g.dispose();

		// Soften.
		float softenFactor = 0.05f;
		float[] softenArray = { 0, softenFactor, 0, softenFactor,
				1 - (softenFactor * 4), softenFactor, 0, softenFactor, 0 };
		Kernel kernel = new Kernel(3, 3, softenArray);
		ConvolveOp cOp = new ConvolveOp(kernel, ConvolveOp.EDGE_NO_OP, null);
		bufferedImage = cOp.filter(bufferedImage, null);

		// Write the jpeg to a file.
		FileOutputStream out = new FileOutputStream(resizedFile);

		// Encodes image as a JPEG data stream
		JPEGImageEncoder encoder = JPEGCodec.createJPEGEncoder(out);

		JPEGEncodeParam param = encoder
				.getDefaultJPEGEncodeParam(bufferedImage);

		param.setQuality(quality, true);

		encoder.setJPEGEncodeParam(param);
		encoder.encode(bufferedImage);
	} // Example usage

	public static boolean compressImage(File file, String directoryFileName,
			int width, int height, boolean proportion) {
		boolean ret = false;
		FileOutputStream fileOutputStream = null;
		try {
			if (file == null || directoryFileName == null) {
				return ret;
			}

			fileOutputStream = new FileOutputStream(new File(directoryFileName));
			Image image = ImageIO.read(file);
			if (image.getWidth(null) == -1) {
				return ret;
			}

			int newWidth = 0;
			int newHeight = 0;

			if (image.getWidth(null) > width || image.getHeight(null) > height) {
				if (proportion) {
					int rate1 = image.getWidth(null) / width;
					int rate2 = image.getHeight(null) / height;
					int rate = rate1 > rate2 ? rate1 : rate2;
					newWidth = image.getWidth(null) / rate;
					newHeight = image.getHeight(null) / rate;
				} else {
					newWidth = width;
					newHeight = height;
				}
			} else {
				newWidth = image.getWidth(null);
				newHeight = image.getHeight(null);
			}

			BufferedImage bufferedImage = new BufferedImage(newWidth,
					newHeight, BufferedImage.TYPE_INT_RGB);

			bufferedImage.getGraphics().drawImage(
					image.getScaledInstance(newWidth, newHeight,
							Image.SCALE_SMOOTH), 0, 0, null);

			JPEGImageEncoder encoder = JPEGCodec
					.createJPEGEncoder(fileOutputStream);
			encoder.encode(bufferedImage);
			fileOutputStream.close();

			ret = true;

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (fileOutputStream != null) {
				try {
					fileOutputStream.close();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		}
		return ret;
	}
	
	
}
