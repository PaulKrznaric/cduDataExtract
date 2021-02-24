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
import java.util.*;

@SuppressWarnings("ALL")
public class main
{

	static boolean OS = isWindows();

	public static boolean isWindows()
	{
		return System.getProperty("os.name").startsWith("Windows");
	}

	/**
	 * Determine if the patient was admitted to the hospital or sent home
	 * @param value An arraylist looking for a string in position 1 of the Array
	 * @return True if the patient was admitted
	 */
	private static boolean wasAdmitted(ArrayList<Object> value)
	{
		return ((String) value.get(1)).equalsIgnoreCase("DIS IN");
	}

	/**
	 * Determine how long the patient was in the CDU clinic for
	 * @param timeIn The time the patient was chcked into the CDU
	 * @param timeOut The time teh patient left the ER
	 * @return an integer representing the number of hours the patient was present
	 */
	public static int lengthOfCDUStay(LocalDateTime timeIn, LocalDateTime timeOut)
	{
		return (int) Duration.between(timeIn, timeOut).toHours();
	}

	/**
	 * Determine if the patient was seen on a weekend (Saturday or Sunday)
	 * @param timeIn the time the patient was seen
	 * @return True if weekend (Saturday or Sunday)
	 */
	private static boolean isWeekend(LocalDateTime timeIn)
	{
		int day = timeIn.getDayOfWeek().getValue();
		return day == 6 || day == 7;

	}

	/**
	 * Check to see if the patietn was seen in the early morning (Before 7am)
	 * @param timeIn The time the patient was checked in
	 * @return whether or not the patient was seen in the morning
	 */

	private static boolean isEarlyMorning(LocalDateTime timeIn)
	{
		return timeIn.getHour() < 7;
	}

	/**
	 * Simple check to see if it was considered at night (after 5pm)
	 * @param timeIn The time the patient was checked in
	 * @return true if the hours fall within the range
	 */
	private static boolean isNight(LocalDateTime timeIn)
	{
		return timeIn.getHour() > 17;
	}

	/**
	 * Calculate what type of billing should be applied to the patient
	 * @param billings the number of assesments
	 * @param timeIn The time the patient checked in to CDU
	 * @return Gives a string that represents what to bill the patients
	 */
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

	/**
	 * Imports the CDU Billings generated by me from the orange sheets
	 * @param path The location of the orange sheets
	 * @return Returns a set that represents the billing numbers of the patients in the orange sheets
	 * @throws IOException
	 */
	private static HashMap<Integer,Boolean> orangeSheetImport(String path) throws IOException
	{
		FileInputStream fis = new FileInputStream(new File(path));
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);
		HashMap<Integer,Boolean> values = new HashMap<>();

		for (Row row : sheet)
		{

			Iterator<Cell> cellIterator = row.cellIterator();
			Cell cell = cellIterator.next();
			//my billing number is on the second row
			if(cellIterator.hasNext())
			{
				cell = cellIterator.next();
			}
			else
			{
				//safety
				continue;
			}
			//only care about hte numbers
			if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
			{
				values.put((int) cell.getNumericCellValue(), false);
			}
		}
		return values;
	}

	/**
	 * Imports the hospital generated sheet
	 * @param path The location of the file
	 * @return Creates an ArrayList where each row is represented by an ArrayList
	 * @throws IOException Need to throw an IOException if import goes wrong.
	 */
	private static ArrayList<ArrayList<Object>> automatedReportImport(String path) throws IOException
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
						//check if it's a date or not
						if (cell.getNumericCellValue() > 1000)
						{
							currentRow.add(cell.getDateCellValue().toInstant().atZone((ZoneId.systemDefault())).toLocalDateTime());
						} else
						{
							currentRow.add(cell.getNumericCellValue());
						}
						break;
					default:
						//add this for preservation of document format
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
	 * @return A string representing the billing to be assigned to the doctors for the patient
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
		String hospitalGeneratedPath, orangeSheetPath = "";
		if (OS)
		{
			hospitalGeneratedPath = "C:\\Users\\Paul\\IdeaProjects\\Dec2020.xlsx";
		} else
		{
			hospitalGeneratedPath = "/Users/paulkrznaric/IdeaProjects/Dec2020.xlsx";
			orangeSheetPath = "/Users/paulkrznaric/Documents/Work/CDU/December CDU 2020 With Automated Items.xlsx";
		}

		try
		{
			ArrayList<ArrayList<Object>> data = automatedReportImport(hospitalGeneratedPath);
			//filter based on Orange sheets
			HashMap<Integer,Boolean> orangeIDs = orangeSheetImport(orangeSheetPath);

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
				if(orangeIDs.containsKey(Integer.parseInt(jNumber)))
				{
					orangeIDs.put(Integer.parseInt(jNumber), true);
					continue;
				}
				currentLine += jNumber + "\t";

				timeIn = (LocalDateTime) current.get(4);
				timeOut = (LocalDateTime) current.get(8);
				duration = lengthOfCDUStay(timeIn, timeOut);

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

				currentLine += buildBilling(billingCount, timeIn, doctors, wasAdmitted(current));

				System.out.println(currentLine);
			}
			Set<Integer> orangeIDValues = orangeIDs.keySet();
			for(Integer i : orangeIDValues)
			{
				if(orangeIDs.get(i) == false)
				{
					System.out.println("Missing ID: " + i);
				}
			}
		} catch (Exception e)
		{
			e.printStackTrace();
			throw e;
		}
	}

}
