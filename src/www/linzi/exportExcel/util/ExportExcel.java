package www.linzi.exportExcel.util;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class ExportExcel {

	public void setTitle(String... titleName) {

		List<String> list = new ArrayList<>();
		for (int i = 0; i < titleName.length; i++) {
			list.add(titleName[i]);
		}

		setTitleNameList(list);

	}

	private List<String> parameterNameList;
	private List<String> titleNameList;
	private List<Integer> colsizeList;

	/**
	 * @return the colsizeList
	 */
	public List<Integer> getColsizeList() {
		return colsizeList;
	}

	/**
	 * @param colsizeList the colsizeList to set
	 */
	public void setColsizeList(List<Integer> colsizeList) {
		this.colsizeList = colsizeList;
	}

	/**
	 * @return the parameterNameList
	 */
	public List<?> getParameterNameList() {
		return parameterNameList;
	}

	/**
	 * @return the titleNameList
	 */
	public List<?> getTitleNameList() {
		return titleNameList;
	}

	public void setParameterName(String... parameterName) {
		List<String> list = new ArrayList<>();
		for (int i = 0; i < parameterName.length; i++) {
			list.add(parameterName[i]);
		}
		setParameterNameList(list);
	}
	public void setColsize(Integer... colSize) {
		List<Integer> list = new ArrayList<>();
		for (int i = 0; i < colSize.length; i++) {
			list.add(colSize[i]);
		}
		setColsizeList(list);
	}
	

	/**
	 * @param parameterNameList
	 *            the parameterNameList to set
	 */
	public void setParameterNameList(List<String> parameterNameList) {
		this.parameterNameList = parameterNameList;
	}

	/**
	 * @param titleNameList
	 *            the titleNameList to set
	 */
	public void setTitleNameList(List<String> titleNameList) {
		this.titleNameList = titleNameList;
	}

	public void exportExcel(String title, List<?> queries, Class<?> clz) {
		export(title, titleNameList, queries, clz, parameterNameList);

	}

	/**
	 * title execel 大标题 titleList 各列的标题 queries 对象的结果集 className 实体类的名字
	 * parameterName 实体类中要导出的属性名
	 * 
	 * @param title
	 * @param titleList
	 * @param queries
	 * @param className
	 * @param parameterName
	 */
	public void export(String title, List<?> titleList, List<?> queries,
			Class<?> clz, List<String> parameterName) {
		try {

			System.out.println("导出EXCEL条" + queries.size() + "数据");
			String fileName = title + ".xls";
			OutputStream os = new FileOutputStream(fileName);
			WritableWorkbook book = Workbook.createWorkbook(os);
			WritableSheet sheet = book.createSheet("title", 0);

			// 字体18
			WritableFont bold = new WritableFont(WritableFont.ARIAL, 17,
					WritableFont.BOLD);
			WritableCellFormat titleFormat = new WritableCellFormat(bold);
			titleFormat.setAlignment(Alignment.CENTRE); // 单元格中的内容水平方向居中
			titleFormat.setVerticalAlignment(VerticalAlignment.CENTRE);// 单元格的内容垂直方向居中
			titleFormat.setBorder(jxl.format.Border.ALL,
					jxl.format.BorderLineStyle.THIN, jxl.format.Colour.BLACK); // BorderLineStyle边框
			titleFormat.setBackground(jxl.format.Colour.WHITE);// 象牙白

			// 字体14
			WritableFont boldtwo = new WritableFont(WritableFont.ARIAL, 12);
			WritableCellFormat titleFormattwo = new WritableCellFormat(boldtwo);
			titleFormattwo.setAlignment(Alignment.CENTRE);// 单元格中的内容水平方向居中
			titleFormattwo.setVerticalAlignment(VerticalAlignment.CENTRE);// 单元格的内容垂直方向居中
			titleFormattwo.setBorder(jxl.format.Border.ALL,
					jxl.format.BorderLineStyle.THIN, jxl.format.Colour.BLACK); // BorderLineStyle边框
			titleFormattwo.setBackground(jxl.format.Colour.WHITE); // 象牙白
			// 标题
			Label label = new Label(0, 0, title, titleFormat);
			sheet.addCell(label);
			sheet.mergeCells(0, 0, titleList.size() - 1, 0);

			// 标题

			for (int i = 0; i < titleList.size(); i++) {
				label = new Label(i, 1, titleList.get(i).toString(),
						titleFormattwo);
				sheet.addCell(label);
				if(colsizeList.get(i)!=null){
				sheet.setColumnView(i, Integer.parseInt(colsizeList.get(i).toString()));
				}else{
					sheet.setColumnView(i, 30);
				}
			}
			for (int i = 0; i < parameterName.size(); i++) {
				List<?> list = Proxy.getPoxy(queries,clz,
						parameterName.get(i));
				for (int j = 0; j < list.size(); j++) {
					label = new Label(i, j + 2, list.get(j).toString(),
							titleFormattwo);
					sheet.addCell(label);
				}
			}
			book.write();
			book.close();
		} catch (Exception ex) {
			String simplename = ex.getClass().getSimpleName();
			if (!"ClientAbortException".equals(simplename)) {
				ex.printStackTrace();
			}
		}
	}

}
