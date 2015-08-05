package com.asomepig.Jacob;

import java.io.File;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class JacobWordInsert {
	public Dispatch documents ;
	public Dispatch doc ;
	public ActiveXComponent word = new ActiveXComponent("Word.Application");
	private Dispatch selection;
	private Object saveOnExit;
	
	/**
	 * creator
	 * @param docPath
	 */
	public JacobWordInsert(String docPath) {
		word.setProperty("Visible", false);  

		documents = word.getProperty("Documents").toDispatch();
//		doc = Dispatch.invoke(documents,"Open",	Dispatch.Method,
//				new Object[] { docPath, new Variant(false),
//				new Variant(false) }, 
//				new int[1]
//				).toDispatch();
		doc = Dispatch.invoke(documents, "Open", Dispatch.Method,  
				new Object[] { docPath,   
						new Variant(false),   
						new Variant(false),//是否只读  
						new Variant(false),   
						new Variant("pwd") },  
				new int[1]).toDispatch();  
		Dispatch.put(doc, "RemovePersonalInformation", false);  
//		Dispatch.call(doc, "ExportAsFixedFormat", pdfFilePath, wdFormatPDF); 

		selection = Dispatch.get(word, "Selection").toDispatch();
	}


	/**
	 * 打开一个已存在的文档
	 * 
	 * @param docPath
	 */
	public void openDocument(String docPath)
	{
//		doc = Dispatch.call(documents, "Open", docPath).toDispatch();
//		selection = Dispatch.get(word, "Selection").toDispatch();
		System.out.println("open a word file!!");
	}
	/**
	 * 在指定书签的添加图片
	 * 
	 * @param bookMarkKey
	 *            书签名
	 * @param imagePath
	 *            图片路径
	 */
	public void addImageAtBookMark(String bookMarkKey, String imagePath)
	{
		Dispatch activeDocument = word.getProperty("ActiveDocument").toDispatch();
		// 得到文档所有的书签
		Dispatch bookMarks = Dispatch.call(activeDocument, "Bookmarks").toDispatch();
		// 获得指定的书签
		boolean bookMarkExist1 = Dispatch
				.call(bookMarks, "Exists", bookMarkKey).getBoolean();
		if (bookMarkExist1 == true)
		{
			System.out.println("已找到书签 " + bookMarkKey);
			Dispatch rangeItem = Dispatch.call(bookMarks, "Item", bookMarkKey)
					.toDispatch();
			Dispatch range = Dispatch.call(rangeItem, "Range").toDispatch();
			// 在书签选中范围内插入图片
			Dispatch.call(Dispatch.get(range, "InLineShapes").toDispatch(),
					"AddPicture", imagePath);

			Dispatch.call(doc, "Save");// 保存
		}
		else
		{
			System.out.println("指定的书签不存在 " + bookMarkKey);
		}
	}
	/**
	 * 在指定书签的添加文本
	 * 
	 * @param bookMarkKey
	 *            书签名
	 * @param insertContent
	 *            插入内容
	 */
	public void addTextAtBookMark(String bookMarkKey, String insertContent)
	{
		Dispatch activeDocument = word.getProperty("ActiveDocument").toDispatch();
		// 得到文档所有的书签
		Dispatch bookMarks = Dispatch.call(activeDocument, "Bookmarks").toDispatch();
		// 获得指定的书签
		boolean bookMarkExist1 = Dispatch
				.call(bookMarks, "Exists", bookMarkKey).getBoolean();
		if (bookMarkExist1 == true)
		{
//			System.out.println("已找到书签 " + bookMarkKey);
			Dispatch rangeItem = Dispatch.call(bookMarks, "Item", bookMarkKey)
					.toDispatch();
			Dispatch range = Dispatch.call(rangeItem, "Range").toDispatch();
			// 在书签选中范围内插入内容
			Dispatch.put(range,"Text", new Variant(insertContent));
			Dispatch.call(doc, "Save");// 保存
		}
		else
		{
			System.out.println("指定的书签不存在 " + bookMarkKey);
		}
	}
	/**
	 * 在当前插入点插入图片
	 * 
	 * @param imagePath
	 *            图片路径
	 */
	public void insertImage(String imagePath)
	{
		Dispatch.call(Dispatch.get(selection, "InLineShapes").toDispatch(),
				"AddPicture", imagePath);
		Dispatch.call(doc, "Save");// 保存
	}
	
	/***************************************************************************
     * 根据书签插入数据
     * 
     * @param bookMarkKey 书签名
     * @param info  插入的数据
     * @return
     */
  
    public boolean intoValueBookMark(String bookMarkKey, String info) throws Exception{
               
        Dispatch activeDocument = word.getProperty("ActiveDocument")
                .toDispatch();
        Dispatch bookMarks = Dispatch.call(activeDocument, "Bookmarks")
                .toDispatch();
        @SuppressWarnings("deprecation")
		boolean bookMarkExist = Dispatch.call(bookMarks, "Exists", bookMarkKey)
                .toBoolean();
        if (bookMarkExist) {
            
            Dispatch rangeItem = Dispatch.call(bookMarks, "Item", bookMarkKey)
                    .toDispatch();
            Dispatch range = Dispatch.call(rangeItem, "Range").toDispatch();
            Dispatch.put(range, "Text", new Variant(info));
            return true;
        } 
        return false;
    }
	
	/**
	 * 把选定的内容或插入点向上移动
	 * 
	 * @param pos
	 *            移动的距离
	 */
	public void moveUp(int pos)
	{
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		for (int i = 0; i < pos; i++)
			Dispatch.call(selection, "MoveUp");
	}

	/**
	 * 把选定的内容或者插入点向下移动
	 * 
	 * @param pos
	 *            移动的距离
	 */
	public void moveDown(int pos)
	{
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		for (int i = 0; i < pos; i++)
			Dispatch.call(selection, "MoveDown");
	}

	/**
	 * 把选定的内容或者插入点向左移动
	 * 
	 * @param pos
	 *            移动的距离
	 */
	public void moveLeft(int pos)
	{
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		for (int i = 0; i < pos; i++)
		{
			Dispatch.call(selection, "MoveLeft");
		}
	}

	/**
	 * 把选定的内容或者插入点向右移动
	 * 
	 * @param pos
	 *            移动的距离
	 */
	public void moveRight(int pos)
	{
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		for (int i = 0; i < pos; i++)
			Dispatch.call(selection, "MoveRight");
	}

	/**
	 * 把插入点移动到文件首位置
	 * 
	 */
	public void moveStart()
	{
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		Dispatch.call(selection, "HomeKey", new Variant(6));
	}

	public void moveEnd()
	{
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		Dispatch.call(selection, "EndKey", new Variant(6));
	}
	/**
	 * 文件保存或另存为
	 * 
	 * @param savePath
	 *            保存或另存为路径
	 */
	public void save(String savePath)
	{
		Dispatch.call(
				(Dispatch) Dispatch.call(word, "WordBasic").getDispatch(),
				"FileSaveAs", savePath);
	}

	/**
	 * 关闭当前word文档
	 * 
	 */
	public void closeDocument()
	{
		if (doc != null)
		{
//			Dispatch.call(doc, "Save");
			Dispatch.call(doc, "Close", new Variant(saveOnExit));
//			Dispatch.call(doc, "Close", new Variant(false));
//			doc = null;
			System.out.println("close the word file!!");
		}
	}

	/**
	 * 关闭全部应用
	 * 
	 */
	public void closeWord()
	{

		if (word != null)
		{
			Dispatch.call(word, "Quit", new Variant(false));
//			word.invoke("Quit", new Variant[] {});
			word.safeRelease();
			word = null;
		}
		selection = null;
		documents = null;
	}

	public File readFile(String path){
		
		File res;
		res = new File(path);
		if(!res.exists())
			{
				System.err.println("文件不存在! in com.testRun.tika.TestPdfbo.java line 35.");
				return null;
			}
		return res;
	}
	
	
}
