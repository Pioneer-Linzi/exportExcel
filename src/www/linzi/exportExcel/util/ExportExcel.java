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
	
	/*
	 *各类要设置的参数 
	 */
	//参数列表
	private List<String> parameterNameList;
	//标题列名列表
	private List<String> titleNameList;
	//各列宽度列表
	private List<Integer> colsizeList;
	//标题类，用于合并列时使用
	private List<Title> titlesList;

	/**
	 * 设置标题列名使用，列名用逗号隔开
	 * @param titleName
	 */
	public void setTitle(String... titleName) {

		List<String> list = new ArrayList<>();
		for (int i = 0; i < titleName.length; i++) {
			list.add(titleName[i]);
		}
		setTitleNameList(list);
	}

	/**
	 * 数据格式 titleName:colStart:rowStart:colEnd:rowEnd titleName 为这个格的内容
	 * colStart这个要合并的格开始列 rowStart要合并的格的开始的行 合并两列:0:1:0:2
	 * 
	 * @param titleList
	 */
	public void setTitleList(String... titleList) {

		List<Title> list = new ArrayList<>();
		for (int i = 0; i < titleList.length; i++) {
			String[] titlePar = titleList[i].split(":");
			Title title = new Title();
			title.setTitleName(titlePar[0]);
			title.setStartCol(Integer.parseInt(titlePar[1]));
			title.setStartRow(Integer.parseInt(titlePar[2]));
			title.setEndCol(Integer.parseInt(titlePar[3]));
			title.setEndRow(Integer.parseInt(titlePar[4]));
			list.add(title);
		}
		setTitlesList(list);
	}


	/**
	 * @return the titleList
	 */
	public List<Title> getTitlesList() {
		return titlesList;
	}

	/**
	 * @param titleList
	 *            the titleList to set
	 */
	public void setTitlesList(List<Title> titlesList) {
		this.titlesList = titlesList;
	}

	/**
	 * @return the colsizeList
	 */
	public List<Integer> getColsizeList() {
		return colsizeList;
	}

	/**
	 * @param colsizeList
	 *            the colsizeList to set
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
	//设置要输出列的属性名
	public void setParameterName(String... parameterName) {
		List<String> list = new ArrayList<>();
		for (int i = 0; i < parameterName.length; i++) {
			list.add(parameterName[i]);
		}
		setParameterNameList(list);
	}
	//要设置的列的宽度，如果不设置，每行30
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
			sheet.mergeCells(0, 0, parameterNameList.size() - 1, 0);

			// 标题

			if (titlesList == null) {

				for (int i = 0; i < titleList.size(); i++) {
					label = new Label(i, 1, titleList.get(i).toString(),
							titleFormattwo);
					sheet.addCell(label);
					if (colsizeList.get(i) != null) {
						sheet.setColumnView(i,
								Integer.parseInt(colsizeList.get(i).toString()));
					} else {
						sheet.setColumnView(i, 30);
					}
				}
			} else {
				for (int i = 0; i < titlesList.size(); i++) {
					Title megtitle = titlesList.get(i);
					label = new Label(megtitle.getStartCol(),
							megtitle.getStartRow(), megtitle.getTitleName(),
							titleFormattwo);
					sheet.addCell(label);
					System.out.println(megtitle.getTitleName());
					sheet.mergeCells(megtitle.getStartCol(),
							megtitle.getStartRow(), megtitle.getEndCol(),
							megtitle.getEndRow());

				}

				// 设置列的宽度
				for (int i = 0; i < parameterNameList.size(); i++) {
					if (colsizeList != null) {
						if (colsizeList.get(i) != null) {
							sheet.setColumnView(i, Integer.parseInt(colsizeList
									.get(i).toString()));
						} else {
							sheet.setColumnView(i, 30);
						}
					} else {
						sheet.setColumnView(i, 30);
					}
				}
			}

			for (int i = 0; i < parameterName.size(); i++) {
				List<?> list = Proxy
						.getPoxy(queries, clz, parameterName.get(i));
				int rowStrat;
				if (titlesList != null) {
					rowStrat = titlesList.get(0).getEndRow() + 1;
				} else {
					rowStrat = 2;
				}
				for (int j = 0; j < list.size(); j++) {
					label = new Label(i, rowStrat, list.get(j).toString(),
							titleFormattwo);
					rowStrat++;
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
