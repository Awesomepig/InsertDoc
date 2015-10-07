/**
 * 
 */
package com.asomepig.jxl;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import com.asomepig.util.FileUtil;
import com.asomepig.util.StringUtil;

/**
 * @author Sylvester
 *
 */
public class JxlTools {

	/**
	 * 根据sheet 和 唯一名字LotNo,获取书签对应的数据集
	 * @param st excel工作簿
	 * @param LotNo 唯一限定名
	 * @param pdfNumber 唯一限定名个数(按照pdf的数量确定)
	 * @return 数据集<key,value>=<书签名,数据值>
	 */
	public Map<String,String> getBookMarkResource(Sheet st,String LotNo,int pdfNumber,boolean ifVersion2){
		Map<String,String> map = new HashMap<String,String>();
		int LotNoLineNumber = 10;
		// 1.LotNo从A10开始,我们从(0,9)-->(0,10)开始,
		loop1:for(int k = 10;k<= 10+pdfNumber;k++)
		{
			Cell cell = st.getCell(0, k);
			String curLotNo = StringUtil.toStr(cell.getContents()).toUpperCase();
			if(curLotNo.equalsIgnoreCase(LotNo))
			{
				LotNoLineNumber = k;
				break loop1;
			}
		}
		
		//2.获取该行的各列的值,设置到map,这些值会根据LotNo改变
			// ////////////////////////------------ 区分版本1，2，设置不同的值----------------//////////////////////////
			if(!ifVersion2)
			{
				int[] colXNumber = {0,1,2,3,6,7,8,10,11,12,13,14,15,16,17};
				String[] bookmarkName = {"A10","B10","C10","D10","G10","H10","I10","K10","L10","M10","N10","O10","P10","Q10","R10"};
				for (int i = 0; i < bookmarkName.length; i++) {
					String curCellContent = StringUtil.toStr(st.getCell(colXNumber[i], LotNoLineNumber).getContents());
					map.put(bookmarkName[i], curCellContent);
				}
			}else//版本二的表格所需字段
			{
				int[] colXNumber = {0,1,2,3,6,7,8,10,11,13,14,15,16};
				String[] bookmarkName = {"A10","B10","C10","D10","G10","H10","I10","K10","L10","N10","O10","P10","Q10"};
				for (int i = 0; i < bookmarkName.length; i++) {
					String curCellContent = StringUtil.toStr(st.getCell(colXNumber[i], LotNoLineNumber).getContents());
					map.put(bookmarkName[i], curCellContent);
				}
			}
			// ////////////////////////------------ 区分版本1，2，设置不同的值----------------//////////////////////////
		//3.获取固定单元格的值
		map.put("C6", StringUtil.toStr(st.getCell(2,5).getContents()));
		
		return map;
	}


	/**
	 * 分割excel
	 * @param exlFile
	 * @param string
	 * @param modelFile 模板文件
	 * @throws WriteException 
	 * @throws RowsExceededException 
	 */
	public void devideExcel(File exlFile, String targetFolder,File modelFile) throws RowsExceededException, WriteException {
		//1、获取excel
		Workbook book;
		Map<String,String[][]> source = new HashMap<String,String[][]>();
		try {
			book = Workbook.getWorkbook(exlFile);
			Sheet st = book.getSheet(0);
			System.out.println("----获取excel数据源");
			// 从第二行开始循环 （1,1）(2,1)--------------------
			String ddh = "";
			List<String[]> list = new ArrayList<String[]>();
			int y = 0;
//			int[] xs = new int[]{2,3,4,5,6,7,8};
			cycle1:while(true)
			{
				y++;
				Cell cell1 = st.getCell(1, y);
				Cell cell2 = st.getCell(2, y);
				// line 1 judge
				if(!cell2.getContents().trim().equals(""))
				{
					if(!cell1.getContents().trim().equals(""))
					{
						if(!ddh.equals("")){
							String[][] a = new String[list.size()][];
							source.put(ddh, list.toArray(a));
						}
						ddh = cell1.getContents().trim();
						list.clear();
					}else if(ddh.equals(""))
					{
						continue cycle1;
					}
					String[] xxx = new String[8];
					// 获取2-8的参数放到list中
					for(int x = 2;x<10;x++)
					{
						xxx[x-2] = st.getCell(x, y).getContents();
					}
					for(String strd:xxx)
						System.out.print("\t "+strd+",");
					System.out.println();
					list.add(xxx);
				}else
					break cycle1;
			}
			book.close();
			System.out.println("----关闭excel数据源");
			// 开始根据source循环
			for (String key:source.keySet()) {
//				List<String[]> cols = new ArrayList<String[]>(); 
				String[][] carr = source.get(key);
				//1.copyFile
				String tarPath = targetFolder+key.trim()+".xls";
				System.out.println("----建立目标文件"+tarPath);
				Workbook sourceBook = Workbook.getWorkbook(modelFile);
				WritableWorkbook tarBook = Workbook.createWorkbook(new File(tarPath),sourceBook);
				WritableSheet ts = tarBook.getSheet(0);
				// 添加订单号
				setCellContent(ts, key, 0, 1);
				String scbh = "";
				// 循环行
				for (int i = 0; i < carr.length; i++) {
					String[] curLine = carr[i];
					// 循环列
					scbh += curLine[0]+",";
					for (int j = 1; j < curLine.length; j++) {
						setCellContent(ts, curLine[j], j,i+3);
					}
					System.out.println("添加了第"+(i+1)+"行");
				}
				// 添加生产编号
				setCellContent(ts, scbh.substring(0, scbh.length()-1), 0, 3);
				System.out.println("关闭了workbook："+tarPath);
				tarBook.write();
				tarBook.close();
				sourceBook.close();
			}
			
		} catch (BiffException | IOException e) {
			System.err.println("导入解析excel错误："+exlFile.getAbsolutePath());
		}
		
	}
	
	
	private void setCellContent(WritableSheet ts,String content,int x,int y){
		WritableCell wc = ts.getWritableCell(x,y);
		if(wc.getType() == CellType.LABEL)    
		{    
			Label label = (Label)wc;    
			label.setString(StringUtil.toStr(content));    
		}else{
			Label l=new Label(x,y,StringUtil.toStr(content)); 
			try {
				ts.addCell(l);
			} catch (RowsExceededException e) {
				e.printStackTrace();
			} catch (WriteException e) {
				e.printStackTrace();
			}
		}
	}
}
