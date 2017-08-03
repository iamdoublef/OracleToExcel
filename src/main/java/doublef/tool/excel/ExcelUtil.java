package doublef.tool.excel;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.Cell;
import jxl.CellType;
import jxl.DateCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableHyperlink;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import doublef.tool.excel.utils.DateSupportUtils;

public class ExcelUtil {



	public static void loadExcel(String driverClass,String url,String userName,String password, String outPath) {
		WritableWorkbook outwb = null;
		try {
			outwb = Workbook.createWorkbook(new File(outPath));
			exportTables(driverClass,url,userName,password,outwb);
			outwb.write();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				outwb.close();
			} catch (Exception e) {
			}
		}
	}

	private static void exportTables(String driverClass,String url,String userName,String password,WritableWorkbook outwb) throws Exception {
		Connection conn = getConnection(driverClass,url,userName,password);
		List<Map<String,String>> tables = listTables(conn);
		
		// 查询指定表下所有列
		String sql = "select cols.column_name as \"columnName\" " + 
				        ",cols.data_type as \"dataType\" " + 
				        ",cols.data_length as \"dataLength\" " + 
				        ",cols.data_precision as \"dataPrecision\" " +  
				        ",cols.data_scale as \"dataScale\" " +  
				        ",ucc.comments as \"comments\" " +  
				        ",case when cols.nullable='Y' then '否' else '是' end as \"notNull\" " + 
				        ",cols.table_name as \"tableName\" " +   
					"from user_tab_columns cols " +  
				        "left join user_col_comments ucc  " + 
				        "on ucc.column_name=cols.column_name and ucc.table_name=cols.table_name " +  
				     "where cols.table_name = ? order by cols.column_id asc ";
		
		PreparedStatement state = prepareStatement(conn, sql);
		
		WritableSheet outSheet = null;
		WritableSheet dirSheet = outwb.createSheet("目录", 0); // 目录
		dirSheet.addCell(new Label(0,0,"表名"));
		dirSheet.addCell(new Label(1,0,"链接"));
		dirSheet.addCell(new Label(2,0,"注释"));
		dirSheet.addCell(new Label(3,0,"备注"));
		dirSheet.setColumnView(0, 30);  
		dirSheet.setColumnView(1, 30);  
		dirSheet.setColumnView(2, 30); 
		dirSheet.setColumnView(3, 30); 
		
		
		int sheetIndex = 1;
		int ri = 0;
		for(Map<String,String> table:tables) {
			outSheet = outwb.createSheet(table.get("name"), sheetIndex++);
			outSheet.setColumnView(0, 30);  
			outSheet.setColumnView(1, 30);  
			outSheet.setColumnView(2, 30); 
			outSheet.setColumnView(3, 30);
			
			outSheet.addCell(new Label(0,1,"表名"));
			outSheet.addCell(new Label(1,1,table.get("name")));
			outSheet.addCell(new Label(2,1,table.get("comment")));
			
			outSheet.addHyperlink(new WritableHyperlink(0,0,"返回目录",dirSheet,1,sheetIndex-1));
			
			dirSheet.addCell(new Label(0,sheetIndex,table.get("name")));
			dirSheet.addHyperlink(new WritableHyperlink(1,sheetIndex,table.get("name"),outSheet,0,0));
			dirSheet.addCell(new Label(2,sheetIndex,table.get("comment")));
			
			outSheet.addCell(new Label(0,3,"字段名"));
			outSheet.addCell(new Label(1,3,"字段类型"));
			outSheet.addCell(new Label(2,3,"是否允许为空"));
			outSheet.addCell(new Label(3,3,"说明"));

			int outRow = 4;
			List<Map<String,String>> cols = listTableColumns(state, table.get("name"));
			for(Map<String,String> col:cols) {
				String columnName = col.get("columnName");
				String dataType = col.get("dataType");
				String dataLength = col.get("dataLength");
				String dataPrecision = col.get("dataPrecision");
				String dataScale = col.get("dataScale");
				String comments = col.get("comments");
				String notNull = col.get("notNull");
				
				outSheet.addCell(new Label(0,outRow,col.get("columnName")));
				outSheet.addCell(new Label(1,outRow,col.get("dataType")));
				outSheet.addCell(new Label(2,outRow,col.get("notNull")));
				outSheet.addCell(new Label(3,outRow++,col.get("comments")));
			}
		}
		
		/*
		
		for (int rowNum = 0; rowNum < rowCount; rowNum++) {
			String colA = getCellString(sheet, 0, rowNum);
			String colB = getCellString(sheet, 1, rowNum);
			String colC = getCellString(sheet, 2, rowNum);
			String colD = getCellString(sheet, 3, rowNum);
			String colE = getCellString(sheet, 4, rowNum);
			String colF = getCellString(sheet, 5, rowNum);
			
			if (colA.equals("实体名")) { // 表开始
				String commt = getTableCommt(state, colE);
				
				// 维护目录
				dirSheet.addCell(new Label(0,sheetIndex,colE));
				//dirSheet.addCell(new Label(1,sheetIndex,colE));
				dirSheet.addCell(new Label(2,sheetIndex,commt));
				
				ri = 1;
				outSheet = outwb.createSheet(colE, sheetIndex++);
				outSheet.setColumnView(0, 30);  
				outSheet.setColumnView(1, 30);  
				outSheet.setColumnView(2, 30); 
				
				// 返回目录
				outSheet.addHyperlink(new WritableHyperlink(0,0,"返回目录",dirSheet,1,sheetIndex-1));
				// 目录指向该sheet
				dirSheet.addHyperlink(new WritableHyperlink(1,sheetIndex-1,colE,outSheet,0,0));
			
				
				Label A = new Label(0, ri, "表名");
				Label B = new Label(1, ri, colE);
				Label C = new Label(2, ri++, commt);
				outSheet.addCell(A);
				outSheet.addCell(B);
				outSheet.addCell(C);
				ri++;
			} else if ("".equals(colA)) { // 表结束
				//System.out.println();
			} else {
				Label A = new Label(0, ri, colE);
				Label B = new Label(1, ri, colF);
				Label C = new Label(2, ri++, colB);
				outSheet.addCell(A);
				outSheet.addCell(B);
				outSheet.addCell(C);
			}

		}*/
		
		state.close();
		conn.close();
	}

	public static Connection getConnection(String driverClass,String url,String userName,String password) {
		Connection conn = null;
		try {
			Class.forName(driverClass);
			conn = DriverManager.getConnection(url, userName,password);
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
		} 
		return conn;
	}

	public static void colseConnection(Connection conn) {
		try {
			conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}
	
	public static PreparedStatement prepareStatement(Connection conn,String sql) {
		try {
			return conn.prepareStatement(sql);
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	
	public static List<Map<String,String>> listTables(Connection conn) throws Exception {
		String sql = "select table_name as \"name\",comments as \"comment\" from user_tab_comments where table_type='TABLE' order by table_name asc";
		Statement stat = conn.createStatement();
		ResultSet rs = stat.executeQuery(sql);
		List<Map<String,String>> list = new ArrayList<Map<String,String>>();
		while(rs.next()) {
			Map<String,String> map = new HashMap<String,String>();
			map.put("name", rs.getString("name"));
			map.put("comment", rs.getString("comment"));
			list.add(map);
		}
		
		rs.close();
		stat.close();
		return list;
	}
	
	/**
	 * "select cols.column_name as \"columnName\" " + 
				        ",cols.data_type as \"dataType\" " + 
				        ",cols.data_length as \"dataLength\" " + 
				        ",cols.data_precision as \"dataPrecision\" " +  
				        ",cols.data_scale as \"dataScale\" " +  
				        ",ucc.comments as \"comments\" " +  
				        ",case when cols.nullable='Y' then '否' else '是' end as \"notNull\" " + 
				        ",cols.table_name as \"tableName\" " +   
					"from user_tab_columns cols " +  
				        "left join user_col_comments ucc  " + 
				        "on ucc.column_name=cols.column_name and ucc.table_name=cols.table_name " +  
				     "where cols.table_name = ? order by cols.column_id asc ";
	 * @throws Exception 
	 */
	
	public static List<Map<String,String>> listTableColumns(PreparedStatement state,String tableName) throws Exception {
		List<Map<String,String>> list = new ArrayList<Map<String,String>>();
		state.setString(1, tableName);
		ResultSet rs = state.executeQuery();
		while(rs.next()) {
			Map<String,String> map = new HashMap<String,String>();
			String columnName = rs.getString("columnName");
			String dataType = rs.getString("dataType");
			String dataLength = rs.getString("dataLength");
			String dataPrecision = rs.getString("dataPrecision");
			String dataScale = rs.getString("dataScale");
			String comments = rs.getString("comments");
			String notNull = rs.getString("notNull");
			map.put("columnName", columnName);
			map.put("dataType", dataType);
			map.put("dataLength", dataLength);
			map.put("dataPrecision", dataPrecision);
			map.put("dataScale", dataScale);
			map.put("comments", comments);
			map.put("notNull", notNull);
			list.add(map);
		}
		rs.close();
		return list;
	}
	
	private static String getCellString(Sheet sheet, int i, int j) {
		String s = "";
		Cell cell = sheet.getCell(i, j);
		try {
			if (cell.getType() == CellType.DATE) {
				Date date = getCellDate(sheet, i, j);
				s = DateSupportUtils.date2str(date);
			} else {
				s = cell.getContents().trim();
			}

		} catch (Exception e) {
			s = "";
		}
		return s;
	}

	/**
	 * jxl默认使用格林时间处理日期，需把该数据转化为当地时间。
	 * 
	 * @param sheet
	 * @param i
	 * @param j
	 * @return
	 */
	private static Date getCellDate(Sheet sheet, int i, int j) {
		Date date = null;
		Cell cell = sheet.getCell(i, j);
		try {
			if (cell.getType() == CellType.DATE) {
				DateCell dc = (DateCell) cell;
				date = DateSupportUtils.GMT2Local(dc.getDate());
			} else {
				String sDate = cell.getContents().trim();
				DateSupportUtils.str2date(sDate);
			}

		} catch (Exception e) {
		}
		return date;
	}
	
	public static void main(String[] args) {
		
		/**
		 * Class.forName("oracle.jdbc.driver.OracleDriver");
			conn = DriverManager.getConnection(
					"jdbc:oracle:thin:@172.16.4.33:1521:ORCL", "CPMDBA",
					"CPMDBA");
		 */
		String driverClass = "oracle.jdbc.driver.OracleDriver";
		String url = "jdbc:oracle:thin:@172.16.4.37:18081:calm";
		String userName = "ERMS";
		String password = "SUNYARD";
		ExcelUtil.loadExcel(driverClass,url,userName,password, "d:/output.xls");
	}

}
