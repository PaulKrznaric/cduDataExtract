
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

	public static boolean isWindows()
	{
		return System.getProperty("os.name").startsWith("Windows");
	}

	private static boolean admitted(ArrayList<Object> value)
	{
		return ((String) value.get(1)).equalsIgnoreCase("DIS IN");
	}

	public static int calculateTime(LocalDateTime timeIn, LocalDateTime timeOut)
	{
		return (int) Duration.between(timeIn, timeOut).toHours();
	}

	private static boolean isWeekend(LocalDateTime timeIn)
	{
		int day = timeIn.getDayOfWeek().getValue();
		if(day == 6|| day == 7)
		{
			return true;
		}
		return false;

	}

	private static boolean isEarlyMorning(LocalDateTime timeIn)
	{
		int hour = timeIn.getHour();
		if(hour < 7)
		{
			return true;
		}
		return false;
	}

	//TODO: isNight Method
	private static boolean isNight(LocalDateTime timeIn)
	{
		return false;
	}

	private static String calculateBilling(int billings, LocalDateTime timeIn)
	{
		String billingInfo = "";
		billingInfo = "1x CDA";
		if(billings == 1)
		{
			if(isEarlyMorning(timeIn))
			{
				billingInfo += " 1x CD2R";
			}
			else if(isWeekend(timeIn))
			{
				billingInfo += " 1x CD5R";
			}
			else if(isNight(timeIn))
			{
				billingInfo += " 1x CD3R";
			}
			else
			{
				billingInfo += " 1x CD0R";
			}
		}
		else
		{
			billingInfo += " " + billings + "x CD0R";

		}

		return billingInfo;
	}

	private static void importSheet(String path) throws IOException
	{
		FileInputStream fis = new FileInputStream(new File(path));
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
				}
			}
			data.add(currentRow);
		}
	}

	public static void main(String[] args) throws IOException
	{
		String path;
		if(OS)
		{
			path = "C:\\Users\\Paul\\IdeaProjects\\Dec2020.xlsx";
		}
		else
		{
			path = "/Users/paulkrznaric/IdeaProjects/Dec2020.xlsx";
		}

		try
		{
			importSheet(path);
			ArrayList<Object> current = new ArrayList<Object>();
			int billingCount, duration;
			String jNumber;
			LocalDateTime timeIn, timeOut;

			for(int i = 1; i < data.size(); i++)
			{
				current = data.get(i);

				jNumber = (String) current.get(0);
				while(jNumber.startsWith("J") || jNumber.startsWith("0"))
				{
					jNumber = jNumber.substring(1);
				}
				System.out.print(jNumber + "\t\t\t");

				timeIn = (LocalDateTime) current.get(4);
				timeOut = (LocalDateTime) current.get(8);
				duration =  calculateTime(timeIn, timeOut);

				if(duration == 2)
				{
					billingCount = 1;

				}
				//TODO: check for multiple doctors
				else if(duration >= 7)
				{
					billingCount = 3;
				}
				else
				{
					billingCount = duration/2;
				}
				System.out.print(calculateBilling(billingCount, timeIn));
				if(admitted(current))
				{
					System.out.print(" 1x CDI");
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
