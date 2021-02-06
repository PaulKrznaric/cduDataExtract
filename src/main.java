
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public  class main
{
	static ArrayList<String> headers = new ArrayList<String>();
	static ArrayList<ArrayList<Object>> data = new ArrayList<ArrayList<Object>>();
	static boolean OS = isWindows();
	public static Boolean isWindows()
	{
		return System.getProperty("os.name").startsWith("Windows");
	}
	public static void main(String[] args) throws IOException
	{

		try
		{
			FileInputStream fis = null;
			if(OS)
			{
				fis = new FileInputStream(new File("C:\\Users\\Paul\\IdeaProjects\\Dec2020.xlsx"));
			}
			else
			{
				//do MacOS stuff here
			}
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheetAt(0);
			Iterator<Row> itr = sheet.iterator();
			while (itr.hasNext())
			{
				Row row = itr.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext())
				{
					Cell cell = cellIterator.next();
					switch (cell.getCellType())
					{
						case Cell.CELL_TYPE_STRING:
							System.out.print(cell.getStringCellValue() + "\t\t\t");
							break;
						case Cell.CELL_TYPE_NUMERIC:
							if(cell.getNumericCellValue() > 50)
							{
								System.out.print(cell.getDateCellValue() + "\t\t\t");
							}
							else
							{
								System.out.print(cell.getNumericCellValue() + "\t\t\t");
							}
							break;
						case Cell.CELL_TYPE_BLANK:
							System.out.print("\t\t\t");
							break;
						default:
							System.out.print("Bad value:" + cell.getCellType() + "\t\t\t");
					}
				}
				System.out.println("");

			}
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}

}
