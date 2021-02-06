
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

@SuppressWarnings("ALL")
public  class main
{
	static ArrayList<ArrayList<Object>> data = new ArrayList<>();

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
		return day == 6 || day == 7;

	}

	private static boolean isEarlyMorning(LocalDateTime timeIn)
	{
		return timeIn.getHour() < 7;
	}

	private static boolean isNight(LocalDateTime timeIn)
	{
		return timeIn.getHour() > 5;
	}

	private static String calculateBilling(int billings, LocalDateTime timeIn)
	{
		String billingInfo = "1x CDA" + "\t\t\t";
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
		for (Row row : sheet)
		{
			Iterator<Cell> cellIterator = row.cellIterator();
			ArrayList<Object> currentRow = new ArrayList<>();
			while (cellIterator.hasNext())
			{
				Cell cell = cellIterator.next();
				switch (cell.getCellType())
				{
					case Cell.CELL_TYPE_STRING:
						currentRow.add(cell.getStringCellValue());
						break;
					case Cell.CELL_TYPE_NUMERIC:
						if (cell.getNumericCellValue() > 1000)
						{
							currentRow.add(cell.getDateCellValue().toInstant().atZone((ZoneId.systemDefault())).toLocalDateTime());
						} else
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
			ArrayList<Object> current = new ArrayList<>();
			int billingCount, duration;
			String jNumber;
			ArrayList<String> doctors;
			LocalDateTime timeIn, timeOut;

			for(int i = 1; i < data.size(); i++)
			{
				current = data.get(i);

				doctors = new ArrayList<>();
				doctors.add((String) current.get(5));
				String otherDoc = (String) current.get(7);
				if(!(otherDoc.equals("") || otherDoc.equals(".")))
				{
					doctors.add((String) current.get(7));
				}

				jNumber = (String) current.get(0);
				while(jNumber.startsWith("J") || jNumber.startsWith("0"))
				{
					jNumber = jNumber.substring(1);
				}
				System.out.print(jNumber + "\t\t\t");

				timeIn = (LocalDateTime) current.get(4);
				timeOut = (LocalDateTime) current.get(8);
				duration =  calculateTime(timeIn, timeOut);

				//TODO: Print Date
				if(duration == 2)
				{
					billingCount = 1;

				}
				else if(duration >= 7 && doctors.size() != 2)
				{
					billingCount = 3;
				}
				else if(duration > 14 && doctors.size() == 2)
				{
					billingCount = 6;
				}
				else
				{
					billingCount = duration/2;
				}
				System.out.print(calculateBilling(billingCount, timeIn) + "\t\t\t");
				if(admitted(current))
				{
					System.out.print(" 1x CDI" + "\t\t\t");
				}
				else
				{
					System.out.print(" 1x CDO" + "\t\t\t");
				}
				if(doctors.size() == 2)
				{
					System.out.print(" Admitting doctor: " + doctors.get(0) + "\t\t\t" + "  Billing Doctor: " + doctors.get(1));
				}
				else
				{
					System.out.print(" Only doctor: " + doctors.get(0));
				}

				System.out.println("");
			}
		}
		catch(Exception e)
		{
			e.printStackTrace();
			throw e;
		}
	}

}
