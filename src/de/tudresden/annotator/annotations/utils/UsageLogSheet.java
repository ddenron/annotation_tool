/**
 * 
 */
package de.tudresden.annotator.annotations.utils;


import org.eclipse.swt.ole.win32.OleAutomation;
import de.tudresden.annotator.oleutils.RangeUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;
import de.tudresden.annotator.oleutils.WorksheetUtils;

/**
 * @author Elvis Koci
 */
public class UsageLogSheet {
	
	protected static final String name = "Usage_Log";
	private static final int startColumnIndex = 1;
	private static final int startRow = 1; 
	
	
	/**
	 * Save the time for action
	 * @param workbookAutomation an OleAutomation to access the embedded workbook
	 */
	public static boolean saveTime(OleAutomation workbookAutomation, String actionName){
			
		OleAutomation usageLogSheet =  WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		
		if(usageLogSheet==null){
			return false;
		}
		
		
		WorksheetUtils.unprotectWorksheet(usageLogSheet);
		
		OleAutomation usedRange = WorksheetUtils.getUsedRange(usageLogSheet);		
		String usedAddress = RangeUtils.getRangeAddress(usedRange);
		usedRange.dispose();
				
		String[] cells = usedAddress.split(":");		
		int endRow = Integer.valueOf(cells[1].replaceAll("[^0-9]+",""));
		int row = endRow + 1;
		
		OleAutomation action = WorksheetUtils.getCell(usageLogSheet, row, startColumnIndex);
		RangeUtils.setValue(action, actionName);
		action.dispose();
		
		OleAutomation time = WorksheetUtils.getCell(usageLogSheet, row, startColumnIndex+1);
		RangeUtils.setValue(time, String.valueOf(System.currentTimeMillis()));
		time.dispose();
		
		WorksheetUtils.protectWorksheet(usageLogSheet);
		WorksheetUtils.setWorksheetVisibility(usageLogSheet, false);
		//WorksheetUtils.setWorksheetVeryHidden(usageLogSheet);
		usageLogSheet.dispose();
		return true;
	}
	
	
	
	
	/**
	 * Save the time files was openned
	 * @param workbookAutomation an OleAutomation to access the embedded workbook
	 */
	public static boolean saveOpenTime(OleAutomation workbookAutomation){
			
		OleAutomation usageLogSheet =  WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		
		if(usageLogSheet==null){
			try{
				usageLogSheet = createUsageLogSheet(workbookAutomation);
				if(usageLogSheet==null){
					return false;
				}
			}catch(Exception ex){
				return false;
			}
		}else{
			WorksheetUtils.setWorksheetVisibility(usageLogSheet, false);
		}
			
		WorksheetUtils.unprotectWorksheet(usageLogSheet);
		
		OleAutomation usedRange = WorksheetUtils.getUsedRange(usageLogSheet);		
		String usedAddress = RangeUtils.getRangeAddress(usedRange);
		usedRange.dispose();
				
		String[] cells = usedAddress.split(":");		
		int endRow = Integer.valueOf(cells[1].replaceAll("[^0-9]+",""));
		int row = endRow + 1;
		
		OleAutomation action = WorksheetUtils.getCell(usageLogSheet, row, startColumnIndex);
		RangeUtils.setValue(action, "Open");
		action.dispose();
				
		OleAutomation fileOpenTime = WorksheetUtils.getCell(usageLogSheet, row, startColumnIndex+1);
		RangeUtils.setValue(fileOpenTime, String.valueOf(System.currentTimeMillis()));
		fileOpenTime.dispose();
		
		WorksheetUtils.protectWorksheet(usageLogSheet);
		usageLogSheet.dispose();
		
		return true;
	}
	
	/**
	 * Create the sheet that will log the usage patterns of the user 
	 * @param workbookAutomation an OleAutomation to access the embedded workbook functionalities
	 * @return the OleAutomation of the created usage_log sheet
	 */
	private static OleAutomation createUsageLogSheet(OleAutomation workbookAutomation){
		
		WorkbookUtils.unprotectWorkbook(workbookAutomation);
		
		OleAutomation usageLogSheet = WorkbookUtils.addWorksheetAsLast(workbookAutomation);
		WorksheetUtils.setWorksheetName(usageLogSheet, name);
		
		createHeaderRow(usageLogSheet);
		
		WorksheetUtils.setWorksheetVisibility(usageLogSheet, false);
		WorkbookUtils.protectWorkbook(workbookAutomation, true, false);
		return usageLogSheet;
	}
	
	
	/**
	 * Create (write) the header row that contains the field names 
	 * @param usageLogSheet  an OleAutomation that provides access to the sheet that maintains the usage log
	 */
	private static void createHeaderRow(OleAutomation usageLogSheet){
		
		OleAutomation field1 = WorksheetUtils.getCell(usageLogSheet, startRow, startColumnIndex);
		RangeUtils.setValue(field1, "Action");
		field1.dispose();
		
		OleAutomation field2 = WorksheetUtils.getCell(usageLogSheet, startRow, startColumnIndex+1);
		RangeUtils.setValue(field2, "TimeInMillis");
		field2.dispose();
	}
	
	
	/**
	 * Hide/Show the sheet that stores the usage log status
	 * @param embeddedWorkbook an OleAutomation that is used to access the functionalities of the workbook that is currently embedded by the application
	 * @param visible true to show the sheet, false to hide it
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean setVisibility(OleAutomation embeddedWorkbook, boolean visible){
		
		OleAutomation usageLogSheet = WorkbookUtils.getWorksheetAutomationByName(embeddedWorkbook, name);
		
		if(usageLogSheet==null)
			return false; 
		
		boolean result = WorksheetUtils.setWorksheetVisibility(usageLogSheet, visible);
		usageLogSheet.dispose();
		return result;
	}
	
	
	/**
	 * Protect the sheet that stores the usage log
	 * @param embeddedWorkbook an OleAutomation that is used to access the functionalities of the workbook that is currently embedded by the application
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean protect(OleAutomation embeddedWorkbook){
		
		OleAutomation usageLogSheet = WorkbookUtils.getWorksheetAutomationByName(embeddedWorkbook, name);
		
		if(usageLogSheet==null)
			return false; 
		
		boolean result = WorksheetUtils.protectWorksheet(usageLogSheet);
		usageLogSheet.dispose();
		return result;
	}
	
	/**
	 * Unprotect the sheet that stores the usage log
	 * @param embeddedWorkbook an OleAutomation that is used to access the functionalities of the workbook that is currently embedded by the application
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean unprotect(OleAutomation embeddedWorkbook){
		
		OleAutomation usageLogSheet = WorkbookUtils.getWorksheetAutomationByName(embeddedWorkbook, name);
		
		if(usageLogSheet==null)
			return false; 
		
		boolean result = WorksheetUtils.unprotectWorksheet(usageLogSheet);
		usageLogSheet.dispose();
		return result;
	}
	
	
	/**
	 * Delete the sheet that stores the usage log
	 * @param embeddedWorkbook an OleAutomation that is used to access the functionalities of the workbook that is currently embedded by the application
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean delete(OleAutomation embeddedWorkbook){
		
		OleAutomation usageLogSheet = WorkbookUtils.getWorksheetAutomationByName(embeddedWorkbook, name);
		
		if(usageLogSheet==null)
			return false; 
		
		boolean result = WorksheetUtils.deleteWorksheet(usageLogSheet);
		usageLogSheet.dispose();
		return result;	
	}
	
	
	/**
	 * @return the name
	 */
	public static String getName() {
		return name;
	}
		
}
