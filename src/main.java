import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;

@SuppressWarnings("ALL")
public class main
{

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
		String billingInfo = "";
		if (billings == 1)
		{
			if (isEarlyMorning(timeIn))
			{
				billingInfo += " 1x CD2R";
			} else if (isWeekend(timeIn))
			{
				billingInfo += " 1x CD5R";
			} else if (isNight(timeIn))
			{
				billingInfo += " 1x CD3R";
			} else
			{
				billingInfo += " 1x CD0R";
			}
		} else if(billings > 1)
		{
			billingInfo += " " + billings + "x CD0R";
		}
		return billingInfo;
	}

	private static HashSet<Integer> importMySheet(String path) throws IOException
	{
		FileInputStream fis = new FileInputStream(new File(path));
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);
		HashSet<Integer> values = new HashSet<>();
		for (Row row : sheet)
		{
			Iterator<Cell> cellIterator = row.cellIterator();
			Cell cell = cellIterator.next();
			if(cellIterator.hasNext())
			{
				cell = cellIterator.next();
			}
			else
			{
				continue;
			}
			if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
			{
				values.add((int) cell.getNumericCellValue());
			}
		}
		return values;
	}

	private static ArrayList<ArrayList<Object>> importSheet(String path) throws IOException
	{
		ArrayList<ArrayList<Object>> data = new ArrayList<ArrayList<Object>>();
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
		return data;
	}


	/**
	 * Build the billings based off the set rules
	 * Doesn't currently split the CD0R
	 * @param billingCount number of billings assigned to ID
	 * @param timeIn Time the patient checked in
	 * @param doctors The list of doctors for the patient
	 * @param admitted Whether the patient was admitted or not
	 * @return
	 */
	public static String buildBilling(int billingCount, LocalDateTime timeIn, ArrayList<String> doctors, boolean admitted)
	{
		String currentLine = doctors.get(0) + ": \t 1x CDA" + "\t";
		if(doctors.size() == 2)
		{
			if(admitted)
			{
				currentLine +=  " 1x CDI" + "\t";
			}
			else
			{
				currentLine += " 1x CDO" + "\t";
			}
			currentLine += doctors.get(1) + ": \t" + calculateBilling(billingCount, timeIn) + "\t";
		}
		else
		{
			currentLine += calculateBilling(billingCount, timeIn) + "\t";
			if (admitted)
			{
				currentLine += " 1x CDI" + "\t";
			} else
			{
				currentLine += " 1x CDO" + "\t";
			}
		}
		return currentLine;
	}

	public static void main(String[] args) throws IOException
	{
		//Set path here
		String path, myPath = "";
		if (OS)
		{
			path = "C:\\Users\\Paul\\IdeaProjects\\Dec2020.xlsx";
		} else
		{
			path = "/Users/paulkrznaric/IdeaProjects/Jan2021.xlsx";
			myPath = "/Users/paulkrznaric/Documents/Work/CDU/January CDU 2021 Final.xlsx";
		}

		try
		{
			ArrayList<ArrayList<Object>> data = importSheet(path);
			//filter based on Orange sheets
			HashSet<Integer> orangeIDs = importMySheet(myPath);

			ArrayList<Object> current;
			int billingCount, duration;
			String jNumber, currentLine;
			ArrayList<String> doctors;
			LocalDateTime timeIn, timeOut;

			for (int i = 1; i < data.size(); i++)
			{
				current = data.get(i);
				currentLine = "";

				doctors = new ArrayList<>();
				doctors.add((String) current.get(5));
				String otherDoc = (String) current.get(7);
				if (!(otherDoc.equals("") || otherDoc.equals(".")))
				{
					if(!otherDoc.equalsIgnoreCase(doctors.get(0)))
					{
						doctors.add((String) current.get(7));
					}
				}

				jNumber = (String) current.get(0);
				while (jNumber.startsWith("J") || jNumber.startsWith("0"))
				{
					jNumber = jNumber.substring(1);
				}
				if(orangeIDs.contains(Integer.parseInt(jNumber)))
				{
					continue;
				}
				currentLine += jNumber + "\t";

				timeIn = (LocalDateTime) current.get(4);
				timeOut = (LocalDateTime) current.get(8);
				duration = calculateTime(timeIn, timeOut);

				currentLine += timeIn.toString() + "\t";

				if (duration == 2)
				{
					billingCount = 1;

				} else if (duration >= 7 && doctors.size() != 2)
				{
					billingCount = 3;
				} else if (duration > 14 && doctors.size() == 2)
				{
					billingCount = 6;
				} else
				{
					billingCount = duration / 2;
				}

				currentLine += buildBilling(billingCount, timeIn, doctors, admitted(current));

				System.out.println(currentLine);
			}
		} catch (Exception e)
		{
			e.printStackTrace();
			throw e;
		}
	}

}
