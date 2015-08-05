/**
 * 
 */
package com.asomepig.jxl;

import java.util.HashMap;
import java.util.Map;

import jxl.Cell;
import jxl.Sheet;

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
	public Map<String,String> getBookMarkResource(Sheet st,String LotNo,int pdfNumber){
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
		int[] colXNumber = {0,1,2,3,6,7,8,10,11,12,13,14,15,16};
		String[] bookmarkName = {"A10","B10","C10","D10","G10","H10","I10","K10","L10","M10","N10","O10","P10","Q10"};
		for (int i = 0; i < bookmarkName.length; i++) {
			String curCellContent = StringUtil.toStr(st.getCell(colXNumber[i], LotNoLineNumber).getContents());
			map.put(bookmarkName[i], curCellContent);
		}
		//3.获取固定单元格的值
		map.put("C6", StringUtil.toStr(st.getCell(2,5).getContents()));
		
		return map;
	}
}
