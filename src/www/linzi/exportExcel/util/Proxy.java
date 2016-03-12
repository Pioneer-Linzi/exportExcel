package www.linzi.exportExcel.util;


import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
public class Proxy {
	public static List<Object> getPoxy(List<?> list, Class<?> clz, String parameter) {
		ArrayList<Object> dataList=new ArrayList<>();
		try {
			String methodName = getMethodName(parameter);
			
			for(int i=0;i<list.size();i++){
			Object obj = list.get(i);
			Method m = null;
			m = clz.getMethod(methodName);
			Object  data = m.invoke(obj);
			System.out.println(data);
			dataList.add(data);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return dataList;
	}

	private static  String getMethodName(String parameter) {
		String c = parameter.substring(0, 1);
		c = c.toUpperCase();
		return "get" + c + parameter.substring(1);
	}

}
