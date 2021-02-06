
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.*;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

public  class main
{
	static ArrayList<ArrayList<Object>> data = new ArrayList<ArrayList<Object>>();
	static boolean OS = isWindows();
	public static Boolean isWindows()
	{
		return System.getProperty("os.name").startsWith("Windows");
	}

	public static int calculateTime(LocalDateTime timeIn, LocalDateTime timeOut)
	{
		return (int) Duration.between(timeIn, timeOut).toHours();
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
				fis = new FileInputStream((new File("/Users/paulkrznaric/IdeaProjects/Dec2020.xlsx")));
				//do MacOS stuff here
			}
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheetAt(0);
			Iterator<Row> itr = sheet.iterator();
			while (itr.hasNext())
			{
				Row row = itr.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				ArrayList<Object> currentRow = new ArrayList<Object>();
				while (cellIterator.hasNext())
				{
					Cell cell = cellIterator.next();
					switch (cell.getCellType())
					{
						case Cell.CELL_TYPE_STRING:
							currentRow.add(cell.getStringCellValue());
							break;
						case Cell.CELL_TYPE_NUMERIC:
							if(cell.getNumericCellValue() > 1000)
							{
								currentRow.add(cell.getDateCellValue().toInstant().atZone((ZoneId.systemDefault())).toLocalDateTime());
							}
							else
							{
								currentRow.add(cell.getNumericCellValue());
							}
							break;
						default:
							currentRow.add("");
							System.out.print("Bad value:" + cell.getCellType() + "\t\t\t");
					}
				}
				data.add(currentRow);
				System.out.println("");
			}
			ArrayList<Object> current = new ArrayList<Object>();
			int billingCount, duration;
			LocalDateTime timeIn, timeOut;
			for(int i = 1; i < data.size(); i++)
			{
				current = data.get(i);
				billingCount = 0;
				System.out.print(current.get(0) + "\t\t\t");
				timeIn = (LocalDateTime) current.get(4);
				if(((String) current.get(1)).equalsIgnoreCase("DIS IN"))
				{
					System.out.print("Admitted" + "\t\t\t");
					timeOut = (LocalDateTime) current.get(10);

				}
				else
				{
					System.out.print("Sent Home" + "\t\t\t");
					timeOut = (LocalDateTime) current.get(8);
				}
				duration =  calculateTime(timeIn, timeOut);
				if(duration == 2)
				{
					billingCount = 1;

				}
				else
				{
					billingCount = duration/2;
				}
			}
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}

}
