/**
 * 
 */
package de.tudresden.annotator.oleutils;

import java.util.Arrays;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.eclipse.swt.SWT;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;
import org.eclipse.swt.widgets.MessageBox;

import de.tudresden.annotator.main.Launcher;

/**
 * @author Elvis Koci
 */
public class WorksheetUtils {
		
	private static final Logger logger = LogManager.getLogger(WorksheetUtils.class.getName());
	
	/**
	 * Get the name of the given worksheet
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return the name of the worksheet
	 */
	public static String getWorksheetName(OleAutomation worksheetAutomation){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		int[] namePropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Name"});	
		if (namePropertyIds == null){			
			logger.error("Could not get ids of property \"Name\" for \"Worksheet\" ole object!");
			return null;
		}		
		
		Variant nameVariant = worksheetAutomation.getProperty(namePropertyIds[0]);
		
		logger.debug("The result of invoking get \"Name\" property was "+nameVariant); 
		if (nameVariant == null) {
			logger.error("Null variant was returned for the \"Name\" of \"Worksheet\" ole object!");
			return null;
		}
		
		String worksheetName = nameVariant.getString();
		nameVariant.dispose();
		
		return worksheetName;
	}
	
	
	/**
	 * set the name of the given worksheet
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @param a string that represents the name to set for the worksheet
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean setWorksheetName(OleAutomation worksheetAutomation, String name){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		int[] namePropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Name"});	
		if (namePropertyIds == null){			
			logger.error("Could not get ids of property \"Name\" for \"Worksheet\" ole object!");
			return false;
		}		
		
		Variant nameVariant = new Variant(name); 
		boolean isSuccess = worksheetAutomation.setProperty(namePropertyIds[0], nameVariant);
		nameVariant.dispose();
		
		logger.debug("Was the action  of setting a name for the sheet successfull? "+isSuccess);
		
		return isSuccess;
	}
	
	
	/**
	 * Get the index of the given worksheet
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return the index number of the worksheet in the collection Workbook.Worksheets 
	 */
	public static int getWorksheetIndex(OleAutomation worksheetAutomation){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		int[] indexPropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Index"});	
		if (indexPropertyIds == null){		
			logger.error("Could not get ids of property \"Index\" for \"Worksheet\" ole object!");
			return 0;
		}		
		
		Variant indexVariant = worksheetAutomation.getProperty(indexPropertyIds[0]);
		logger.debug("The result of invoking get \"Index\" property was "+indexVariant);
		
		if (indexVariant == null) {
			logger.error("Invoking get \"Index\" property returned  null variant for \"Worksheet\" ole object!");
			return 0;
		}
		
		int worksheetIndex = indexVariant.getInt();
		indexVariant.dispose();
		
		return worksheetIndex;
	}
	
	
	/**
	 * set worksheet visibility
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @param true to set worksheet visible, false to hide it
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean setWorksheetVisibility(OleAutomation worksheetAutomation, boolean visible){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		if(worksheetAutomation==null)
			return false;
		
		int[] visiblePropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Visible"});			
		Variant visiblePropertyVariant = new Variant(visible); 
		boolean isSuccess = worksheetAutomation.setProperty(visiblePropertyIds[0], visiblePropertyVariant);
		visiblePropertyVariant.dispose();
		
		return isSuccess;
	}
	
	
	/**
	 * set worksheet very hidden
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean setWorksheetVeryHidden(OleAutomation worksheetAutomation){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		if(worksheetAutomation==null)
			return false;
		
		int[] visiblePropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Visible"});
		
		Variant visiblePropertyVariant = new Variant(2); 
		boolean isSuccess = worksheetAutomation.setProperty(visiblePropertyIds[0], visiblePropertyVariant);
		visiblePropertyVariant.dispose();
		
		return isSuccess;
	}
	
	
	/**
	 * Make the given worksheet the active sheet.
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean makeWorksheetActive(OleAutomation worksheetAutomation){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		int[] activateMethodsIds = worksheetAutomation.getIDsOfNames(new String[]{"Activate"});	
		if (activateMethodsIds == null){			
			logger.error("Could not get ids of method \"Activate\" for \"Worksheet\" ole object! "+
					"The worksheet name is "+WorksheetUtils.getWorksheetName(worksheetAutomation));
			return false;
		}		

		Variant result = worksheetAutomation.invoke(activateMethodsIds[0]);
		logger.debug("The result of invoking method \"Activate\" for \"Worksheet\" ole object was "+result);
		if(result==null){
			logger.error("Invoking method \"Activate\" returned  null variant for \"Worksheet\" ole object! "+ 
					"The worksheet name is "+WorksheetUtils.getWorksheetName(worksheetAutomation));
			return false;
		}
		
		result.dispose();
		return true;
	}
	
	
	/**
	 * Make all data in the worksheet visible. Also, invoking this method will remove filters on the sheet data.  
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean showAllWorksheetData(OleAutomation worksheetAutomation){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		int[] showAllDataMethodsIds = worksheetAutomation.getIDsOfNames(new String[]{"ShowAllData"});	
		if (showAllDataMethodsIds == null){			
			logger.error("Could not get ids of method \"ShowAllData\" for \"Worksheet\" ole object! "+
					"The worksheet name is "+WorksheetUtils.getWorksheetName(worksheetAutomation));
			return false;
		}		

		Variant result = worksheetAutomation.invoke(showAllDataMethodsIds[0]);
		logger.debug("The result of invoking method \"ShowAllData\" for \"Worksheet\" ole object was "+result);
		if(result==null){
			logger.error("Invoking method \"ShowAllData\" returned  null variant for \"Worksheet\" ole object! "+
					"The worksheet name is "+WorksheetUtils.getWorksheetName(worksheetAutomation));
			return false;
		}
		
		result.dispose();
		return true;
	}
	
	
	/**
	 * Get the Application automation from the given worksheet
	 * @param sheetAutomation an OleAutomation that provides access to the functionalities of the Worksheet OLE object 
	 * @return an OleAutomation to access the (Excel) application
	 */
	public static OleAutomation getApplicationAutomation(OleAutomation sheetAutomation){
		
		logger.debug("Is sheet automation null? "+sheetAutomation==null);
		
		int[] applicationPropertyIds = sheetAutomation.getIDsOfNames(new String[]{"Application"}); 
		Variant applicationVariant =  sheetAutomation.getProperty(applicationPropertyIds[0]);
		OleAutomation applicationAutomation = applicationVariant.getAutomation();
		applicationVariant.dispose();
		
		return applicationAutomation;
	}
	
	/**
	 * Get the specified range automation. The address of the top_left_cell and down_right_cell have to be provided.
	 * For single cell ranges the address of the down_right_cell is NULL or the same as the top_left_cell    
	 * @param worksheetAutomation an OleAutomation object for accessing the Worksheet OLE object
	 * @param topLeftCell address of top left cell (e.g., "A1" or "$A$1" )
	 * @param downRightCell address of down right cell (e.g., "C3" or "$C$3" )
	 * @return an OleAutomation that provides access the specified range 
	 */
	public static OleAutomation getRangeAutomation(OleAutomation worksheetAutomation, String topLeftCell, String downRightCell){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		// get the OleAutomation object for the selected range 
		int[] rangePropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Range"});
		
		Variant[] args;
		if(downRightCell!=null && downRightCell.length()>1){
			args = new Variant[2];
			args[0] = new Variant(topLeftCell);
			args[1] = new Variant(downRightCell);
		}else{
			args = new Variant[1];
			args[0] = new Variant(topLeftCell);
		}
		
		Variant rangeVariant = worksheetAutomation.getProperty(rangePropertyIds[0],args);
		
		if(rangeVariant==null || rangeVariant.getType()==0){
			return null;
		}
			
		OleAutomation rangeAutomation = rangeVariant.getAutomation();
		for (Variant arg : args) {
			arg.dispose();
		}
		rangeVariant.dispose();
		
		return rangeAutomation;
	}
	
	
	/**
	 * Get the specified range automation given its address. 
	 * This method supports simple ranges having only one area (i.e., not multi-selection)
	 * @param worksheetAutomation an OleAutomation object for accessing the Worksheet OLE object
	 * @param rangeAddress a string representing the address the range (e.g., "$A$1:$C$2", "$D$4", "B2:G6")
	 * @return an OleAutomation that provides access the specified range 
	 */
	public static OleAutomation getRangeAutomation(OleAutomation worksheetAutomation, String rangeAddress){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		String[] subStrings = rangeAddress.split(":");
		String topLeftCell = subStrings[0];
		String downRightCell = null;
		if (subStrings.length == 2)
			downRightCell = subStrings[1];
		
		return getRangeAutomation(worksheetAutomation, topLeftCell, downRightCell);
	}

	
	/**
	 * Get the OleAutomation for the specified multi-selection range automation 
	 * @param worksheetAutomation an OleAutomation object for accessing the Worksheet OLE object
	 * @param multiSelectionRange a string that represents the address of the multi-selection range (Ex. "$A$1:$C$2, $A$4:$C$4, $D$6" ) 
	 * @return an OleAutomation that provides access to the multi-selection range 
	 */
	public static OleAutomation getMultiSelectionRangeAutomation(OleAutomation worksheetAutomation, String multiSelectionRange){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		// get the OleAutomation object for the multi-selection (multi-area) range 
		int[] rangePropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Range"});
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		
		Variant[] args = new Variant[1];
		args[0] = new Variant(multiSelectionRange);

		Variant rangeVariant = worksheetAutomation.getProperty(rangePropertyIds[0], args);
		OleAutomation rangeAutomation = rangeVariant.getAutomation();
		
		for (Variant arg : args) {
			arg.dispose();
		}
		
		return rangeAutomation;
	}
	

	/**
	 * Get the used range for the given worksheet
	 * @param worksheetAutomation an OleAutomation object for accessing the Worksheet OLE object
	 * @return an OleAutomation for accessing the used range
	 */
	public static OleAutomation getUsedRange(OleAutomation worksheetAutomation){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		int[] usedRangePropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"UsedRange"});	
		Variant usedRangeVariant = worksheetAutomation.getProperty(usedRangePropertyIds[0]);
		OleAutomation usedRangeAutomation = usedRangeVariant.getAutomation();
		usedRangeVariant.dispose();
		
		return usedRangeAutomation;
	}
	
	
	/**
	 * Get a cell from the specified worksheet given the row and the column number. 
	 * @param worksheetAutomation an OleAutomation to access the worksheet that contains the cell
	 * @param row an integer that represents the row number (index)
	 * @param column an integer that represents the column number (index)
	 * @return an OleAutomation that provides access to the cell. 
	 */
	public static OleAutomation getCell(OleAutomation worksheetAutomation, int row, int column){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		int[] cellsPropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Cells"}); 
		
		Variant[] args = new Variant[2];
		args[0] = new Variant(row);
		args[1] = new Variant(column);
		
		Variant cellsVariant = worksheetAutomation.getProperty(cellsPropertyIds[0], args);	
		OleAutomation cellAutomation = cellsVariant.getAutomation();
		cellsVariant.dispose();
		
		args[0].dispose();
		args[1].dispose();
		
		return cellAutomation; 
	}
		
	/**
	 * Get a all cells from the specified sheet 
	 * @param worksheetAutomation an OleAutomation to access the sheet that contains the cells
	 * @return an OleAutomation that provides access to the cells of the sheet. 
	 */
	public static OleAutomation getCells(OleAutomation worksheetAutomation){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		int[] cellsPropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Cells"}); 
		
		Variant cellsVariant = worksheetAutomation.getProperty(cellsPropertyIds[0]);
		
		if(cellsVariant==null){
			return null;
		}
			
		OleAutomation cellsAutomation = cellsVariant.getAutomation();
		cellsVariant.dispose();
	
		return cellsAutomation; 
	}
	
	
	/**
	 * Get collection of columns in the worksheet 
	 * @param worksheetAutomation an OleAutomation to access the worksheet that contains the cell
	 * @return an OleAutomation that provides access to the collection of columns in the worksheet 
	 */
	public static OleAutomation getColumns(OleAutomation worksheetAutomation){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		int[] columnsPropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Columns"}); 
		Variant columnsPropertyVariant = worksheetAutomation.getProperty(columnsPropertyIds[0]);	
		
		OleAutomation columnsAutomation =  columnsPropertyVariant.getAutomation();
		columnsPropertyVariant.dispose();
		
		return columnsAutomation;
	}
	

	/**
	 * Get a specific column in the worksheet 
	 * @param worksheetAutomation an OleAutomation to access the worksheet that contains the cell
	 * @return an OleAutomation that provides access to the range that represents the entire specified column in the worksheet
	 */
	public static OleAutomation getColumn(OleAutomation worksheetAutomation, String column){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		int[] columnsPropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Columns"}); 
		
		Variant[] args= new Variant[1];
		args[0] = new Variant(column);
		
		Variant columnPropertyVariant = worksheetAutomation.getProperty(columnsPropertyIds[0], args);	
		args[0].dispose();
		
		OleAutomation columnAutomation =  columnPropertyVariant.getAutomation();
		columnPropertyVariant.dispose();
		
		return columnAutomation;
	}
	
	
	/**
	 * Get collection of rows in the worksheet 
	 * @param worksheetAutomation an OleAutomation to access the worksheet that contains the cell
	 * @return an OleAutomation that provides access to the collection of rows in the worksheet 
	 */
	public static OleAutomation getRows(OleAutomation worksheetAutomation){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		int[] rowsPropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Rows"}); 
		Variant rowsPropertyVariant = worksheetAutomation.getProperty(rowsPropertyIds[0]);	
		OleAutomation rowsAutomation = rowsPropertyVariant.getAutomation();
		rowsPropertyVariant.dispose();
		
		return rowsAutomation;
	}
	
	
	/**
	 * Get a specific row from the worksheet
	 * @param worksheetAutomation an OleAutomation to access the worksheet that contains the cell
	 * @return an OleAutomation that provides access to the range that represents the entire specified row in the worksheet
	 */
	public static OleAutomation getRow(OleAutomation worksheetAutomation, int row){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		int[] rowsPropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Rows"}); 

		Variant[] args= new Variant[1];
		args[0] = new Variant(row);
		
		Variant rowPropertyVariant = worksheetAutomation.getProperty(rowsPropertyIds[0], args);
		args[0].dispose();
		
		OleAutomation rowAutomation = rowPropertyVariant.getAutomation();
		rowPropertyVariant.dispose();
		
		return rowAutomation;
	}
	
	
	/**
	 * Get the OleAutomation object for the "Shapes" property of the given worksheet  
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return an OleAutomation to access the Shapes of the worksheet 
	 */
	public static OleAutomation getWorksheetShapes(OleAutomation worksheetAutomation){
		
		logger.debug("Is sheet automation null? ".concat(String.valueOf(worksheetAutomation==null)));
		
		int[] shapesPropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Shapes"});	
		if (shapesPropertyIds == null){		
			logger.error("Could not get the id of the \"Shapes\" property for \"Worksheet\" ole object. "+
					"The worksheet name is "+WorksheetUtils.getWorksheetName(worksheetAutomation));			
			return null;
		}		
		
		Variant shapesVariant = worksheetAutomation.getProperty(shapesPropertyIds[0]);
		
		logger.debug("Invoking get \"Shapes\" property for \"Worksheet\" object returned: "+shapesVariant);
		if (shapesVariant == null) {
			logger.error("Invoking get \"Shapes\" property for \"Worksheet\" ole object returned null variant. "+
					"The worksheet name is "+WorksheetUtils.getWorksheetName(worksheetAutomation));
			return null;
		}
		
		OleAutomation worksheetShapes = shapesVariant.getAutomation();
		shapesVariant.dispose();
		
		return worksheetShapes;		
	}
	
	
	/**
	 * Protect the data, formating, and structure of the specified worksheet
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean protectWorksheet(OleAutomation worksheetAutomation){
		
		// get the id of the "Protect" method and the considered parameters
		// you can find the documentation of this OLE method at 
		// https://msdn.microsoft.com/EN-US/library/ff840611.aspx
		int[] protectMethodIds = worksheetAutomation.getIDsOfNames(new String[]{"Protect", 
				"AllowFormattingColumns", "AllowFormattingRows"});
		
		if (protectMethodIds == null) {
			logger.fatal("Could not get the ids of the \"Protect\" method for \"Worksheet\" ole object. "+
					"The worksheet name is "+WorksheetUtils.getWorksheetName(worksheetAutomation));
		}else{
			Variant[] args = new Variant[2];
			args[0] = new Variant(false); // allow user to resize columns
			args[1] = new Variant(false); // allow user to resize rows
			
			int argsIds[] = Arrays.copyOfRange(protectMethodIds, 1, protectMethodIds.length);
			
			Variant result = worksheetAutomation.invoke(protectMethodIds[0], args, argsIds);	
			logger.debug("Invoking \"Protect\" method for \"Worksheet\" object returned: "+result);
			
			if(result==null){	
				logger.fatal("Invoking \"Protect\" method for \"Worksheet\" ole object returned null variant. "+
						"The worksheet name is "+WorksheetUtils.getWorksheetName(worksheetAutomation));
				
				MessageBox messageBox = Launcher.getInstance().createMessageBox(SWT.ICON_ERROR);
				messageBox.setText("ERROR");
	            messageBox.setMessage("ERROR: Could not protect the sheet "
	            		+ "\""+WorksheetUtils.getWorksheetName(worksheetAutomation)+"\"!");
	            messageBox.open();
	            
	            Launcher.getInstance().quitApplication();			
				// return false;
			}
			
			result.dispose();
			for (Variant arg: args) {
				arg.dispose();
			}
		}				

		return true;
	}
	
	
	/**
	 * Unprotect the specified worksheet 
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean unprotectWorksheet(OleAutomation worksheetAutomation){
		
		// get the id of the "Unprotect" method for worksheet OLE object 
		int[] unprotectMethodIds = worksheetAutomation.getIDsOfNames(new String[]{"Unprotect"});
		if(unprotectMethodIds==null){
			logger.error("Could not get the ids of the \"Unprotect\" method for \"Worksheet\" ole object "+
					"The worksheet name is "+WorksheetUtils.getWorksheetName(worksheetAutomation));
		}
		
		// invoke the unprotect method  
		Variant result = worksheetAutomation.invoke(unprotectMethodIds[0]);
		logger.debug("Invoking \"Unprotect\" method for \"Worksheet\" object returned: "+result);
		if(result==null){
			logger.fatal("Invoking \"Unprotect\" method for \"Worksheet\" ole object returned null variant. "+
					"The worksheet name is "+WorksheetUtils.getWorksheetName(worksheetAutomation));
			
			MessageBox messageBox = Launcher.getInstance().createMessageBox(SWT.ICON_ERROR);
			messageBox.setText("ERROR");
            messageBox.setMessage("ERROR: Could not unprotect the sheet "
            		+ "\""+WorksheetUtils.getWorksheetName(worksheetAutomation)+"\"!");
            messageBox.open();
            
            Launcher.getInstance().quitApplication();
         	// return false;
		}	
		
		result.dispose();
		return true;
	}
	
	/**
	 * Export the data in the given worksheet as a CSV file 
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @param filePath the path of the file to save the data 
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean saveAsCSV(OleAutomation worksheetAutomation, String filePath){
		
		int[] saveAsMethodIds = worksheetAutomation.getIDsOfNames(new String[]{"SaveAs", "FileName", "FileFormat"});	
		
		Variant[] args = new Variant[2];
		args[0] = new Variant(filePath);
		args[1] = new Variant(6); // xlCSV = 6 
		
		int argsIds[] = Arrays.copyOfRange(saveAsMethodIds, 1, saveAsMethodIds.length); 
		Variant result = worksheetAutomation.invoke(saveAsMethodIds[0], args, argsIds);
		
		if(result==null)
			return false;
		
		for (Variant arg : args) {
			arg.dispose();
		}
		result.dispose();
		
		return true;
	}
	
	
	/**
	 * Delete the specified worksheet 
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean deleteWorksheet(OleAutomation worksheetAutomation){
		
		// get the id of the "Unprotect" method for worksheet OLE object 
		int[] deleteMethodIds = worksheetAutomation.getIDsOfNames(new String[]{"Delete"});
		if(deleteMethodIds==null){
			logger.error("Could not get the ids of the \"Delete\" method for \"Worksheet\" ole object "+
					"The worksheet name is "+WorksheetUtils.getWorksheetName(worksheetAutomation));
		}
		
		// invoke the unprotect method  
		Variant result = worksheetAutomation.invoke(deleteMethodIds[0]);
		logger.debug("Invoking \"Delete\" method for \"Worksheet\" object returned: "+result);
		if(result==null){
			logger.error("Invoking \"Delete\" method for \"Worksheet\" ole object returned null variant. "+
					"The worksheet name is "+WorksheetUtils.getWorksheetName(worksheetAutomation));
		}	
		
		boolean isSuccess = result.getBoolean();
		result.dispose();
		
		return isSuccess;
	}
}
