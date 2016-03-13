### 导出表格的一个小工具

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
	 	/**
	 * 数据格式 titleName:colStart:rowStart:colEnd:rowEnd titleName 为这个格的内容
	 * colStart这个要合并的格开始列 rowStart要合并的格的开始的行 合并两列:0:1:0:2
	 * 
	 * @param titleList
	 */
	public void setTitleList(String... titleList) 
	//要设置的列的宽度，如果不设置，每行30
	public void setColsize(Integer... colSize) 
	 
	 ### 使用方法
	 
	 1 先设置setParameterName(String... parameterName)，你的实体类要显示的属性名
	 2 setTitle(String... titleName) 表格的列名
	 3 exportExcel(String title, List<?> queries, String className) 
	 4 这里也可以设置合并单元格跟单无格的宽度，
