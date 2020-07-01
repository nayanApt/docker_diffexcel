import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.nocrala.tools.texttablefmt.BorderStyle;
import org.nocrala.tools.texttablefmt.ShownBorders;
import org.nocrala.tools.texttablefmt.Table;


public class DiffExcel	{

	static String sheetName = new String();
	static int sheetNum = 0;
	static String foodName = new String();
	static String cell = new String();
	static String cell1 = new String();
	static String cell2 = new String();
	static Table table = null;

	public static void main(String[] args)	{

		try	{
			FileInputStream db1 = new FileInputStream(new File(args[0]));
			FileInputStream db2 = new FileInputStream(new File(args[1]));

			XSSFWorkbook wb1 = new XSSFWorkbook(db1);
			XSSFWorkbook wb2 = new XSSFWorkbook(db2);

			for (int i = 0; i < wb1.getNumberOfSheets(); i++)	{

				System.out.println("Sheet: "+wb1.getSheetName(i)+"\n");

				sheetNum = i;

				table = new Table(5, BorderStyle.DESIGN_TUBES_WIDE, ShownBorders.HEADER_AND_COLUMNS);
				table.addCell("Row number");
				table.addCell("Item");
				table.addCell("Column name");
				table.addCell("File 1");
				table.addCell("File 2");

				if (compareSheets(wb1.getSheetAt(i), wb2.getSheetAt(i)))	{
					System.out.println("\nThe two sheets are equal.");
				}

				System.out.println("============================================================");
			}

			db1.close();
			db2.close();
		}
		catch(Exception e)	{
			e.printStackTrace();
		}
	}
	static boolean compareSheets(XSSFSheet s1, XSSFSheet s2)	{

		boolean equalSheets = true;

		XSSFRow r0 = s1.getRow(0);

		for (int i = s1.getFirstRowNum(); i <= s1.getLastRowNum(); i++)	{

			XSSFRow r1 = s1.getRow(i);
			XSSFRow r2 = s2.getRow(i);

			if (!(compareRows(r0, r1, r2)))	{
				equalSheets = false;
				table.addCell(""+(i+1));
				table.addCell(foodName);
				table.addCell(cell);
				table.addCell(cell1);
				table.addCell(cell2);
			}
		}
		System.out.println(table.render());
		return equalSheets;
	}
	static boolean compareRows(XSSFRow r0, XSSFRow r1, XSSFRow r2)	{
		if ((r1 == null) && (r2 == null))	{
			return true;
		}
		else if ((r1 == null) || (r2 == null))	{
			foodName = "***NEW ROW***";
			cell = "";
			cell1 = "";
			cell2 = "";
			return false;
		}

		boolean equalRows = true;

		for (int i = 0; i <= 31; i++)	{

			XSSFCell c1 = r1.getCell(i);
			XSSFCell c2 = r2.getCell(i);

			if (!(compareCells(c1, c2)))	{
				equalRows = false;
				XSSFCell cellName = r0.getCell(i);
				XSSFCell food = r1.getCell(2);

				if (sheetNum == 1)
					food = r1.getCell(0);

				cellName.setCellType(CellType.STRING);
				food.setCellType(CellType.STRING);

				foodName = food.getStringCellValue();
				cell = cellName.getStringCellValue();
			}
		}
		return equalRows;
	}
	static boolean compareCells(XSSFCell c1, XSSFCell c2)	{

		if ((c1 == null) && (c2 == null))	{
			return true;
		}

		else if ((c1 == null) && (c2 != null))	{

			if (c2.getCellTypeEnum() == CellType.BLANK)
				return true;

			c2.setCellType(CellType.STRING);
			cell1 = "";
			cell2 = c2.getStringCellValue();
			return false;
		}

		else if ((c1 != null) && (c2 == null))	{

			if (c1.getCellTypeEnum() == CellType.BLANK)
				return true;

			c1.setCellType(CellType.STRING);
			cell1 = c1.getStringCellValue();
			cell2 = "";
			return false;
		}

		else if ((c1 != null) && (c2 != null))	{

			c1.setCellType(CellType.STRING);
			c2.setCellType(CellType.STRING);

			if (c1.getStringCellValue().equals(c2.getStringCellValue()))
				return true;

			cell1 = c1.getStringCellValue();
			cell2 = c2.getStringCellValue();
			return false;
		}
		return true;
	}
}
