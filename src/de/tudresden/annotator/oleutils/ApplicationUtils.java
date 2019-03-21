/**
 * 
 */
package de.tudresden.annotator.oleutils;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleControlSite;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class ApplicationUtils {
	
	private static final Logger logger = LogManager.getLogger(ApplicationUtils.class.getName());
	
	/**
	 * Get Excel application as an OleAutomation object
	 * @param controlSite the OleControlSide for the embedded spreadsheet file 
	 * @return an OleAutomation that provides access to the functionalities of the (Excel) Application OLE object 
	 */
	public static OleAutomation getApplicationAutomation(OleControlSite controlSite){
		
		OleAutomation excelClient = null;
		try {
			excelClient = new OleAutomation(controlSite);
		} catch (IllegalArgumentException iaEx) {
			logger.fatal("Illegal argument exception on creation of excel client OleAutomation", iaEx);
		} catch (Exception e) {
			logger.error("Genereric exception on creation of excel client OleAutomationn", e);
		} 
		
		OleAutomation application = null;
		if(excelClient!=null){
			
			int[] dispIDs = excelClient.getIDsOfNames(new String[] {"Application"});
			
			if(dispIDs==null){	
				logger.error("Could not get \"Application\" property ids for \"Excel Client\"!");
				return null;
			}
			
			Variant pVarResult = excelClient.getProperty(dispIDs[0]);
			if(pVarResult==null){	
				logger.error("Get \"Application\" property for \"Excel Client\" returned null variant!");
				return null;
			}
			
			logger.debug("Get \"Application\" property for \"Excel Client\" returned variant: "+pVarResult);
			application = pVarResult.getAutomation();
			
			pVarResult.dispose();
			excelClient.dispose();
		}
		
		return application;
	}
	

	/**
	 * Get OleAutomation for the active workbook using the "ActiveWorkbook" property. 
	 * Excel application considers the workbook which has the focus to be the "active" one.
	 *  
	 * @param application an OleAutomation that provides access to the functionalities of the (Excel) Application OLE object
	 * @return an OleAutomation that provides access to the functionalities of the Active Workbook OLE object
	 */
	public static OleAutomation getActiveWorkbookAutomation(OleAutomation application){
		
		if(application==null){
			logger.debug("Method getActiveWorkbookAutomation received null application OleAutomation object");
		}
		
		int[] workbookIds = application.getIDsOfNames(new String[]{"ActiveWorkbook"});	
		if (workbookIds == null){			
			logger.error("Could not get \"ActiveWorkbook\" property ids for \"Application\" object!");
			return null;
		}		
		Variant workbookVariant = application.getProperty(workbookIds[0]);
		if (workbookVariant == null) {
			logger.error("Get \"ActiveWorkbook\" property for \"Application\" returned null variant!");
			return null;
		}		
		
		logger.debug("Get \"ActiveWorkbook\" property for \"Application\" returned variant: "+workbookVariant);
		
		OleAutomation workbookAutomation =  workbookVariant.getAutomation();
		workbookVariant.dispose();
		
		return workbookAutomation;
	}
	
	/**
	 * Get OleAutomation for the active window using the "ActiveWindow" property. 
	 * Excel application considers the window which is on top to be the "active" one.
	 *  
	 * @param application an OleAutomation that provides access to the functionalities of the (Excel) Application OLE object
	 * @return an OleAutomation that provides access to the functionalities of the Active Window OLE object
	 */
	public static OleAutomation getActiveWindowAutomation(OleAutomation application){
		
		if(application==null){
			logger.debug("Method getActiveWindowAutomation received a null Application OleAutomation object");
		}
		
		int[] workbookIds = application.getIDsOfNames(new String[]{"ActiveWindow"});	
		if (workbookIds == null){			
			logger.error("Could not get \"ActiveWindow\" property ids for \"Application\" object! \n"+application.getLastError());
			return null;
		}		
		Variant workbookVariant = application.getProperty(workbookIds[0]);
		if (workbookVariant == null) {
			logger.error("Get \"ActiveWindow\" property for \"Application\" returned null variant! \n"+application.getLastError());
			return null;
		}		
		
		logger.debug("Get \"ActiveWindow\" property for \"Application\" returned variant: "+workbookVariant);
		
		OleAutomation workbookAutomation =  workbookVariant.getAutomation();
		workbookVariant.dispose();
		
		return workbookAutomation;
	}
	
	/**
	 * Get the OleAutomation object for the embedded workbook using (given) its name   
	 * @param application an OleAutomation that provides access to the functionalities of the (Excel) Application OLE object
	 * @param workbookName the name of the embedded workbook
	 * @return an OleAutomation that provides access to the functionalities of the Embedded Workbook OLE object
	 */
	public static OleAutomation getEmbeddedWorkbookAutomation(OleAutomation application, String workbookName){
		
		int[] workbooksIds = application.getIDsOfNames(new String[]{"Workbooks"});	
		if (workbooksIds == null){			
			System.out.println("\"Workbooks\" property not found for \"Application\" object!");
			return null;
		}		
		
		Variant workbooksVariant = application.getProperty(workbooksIds[0]);
		if (workbooksVariant == null) {
			System.out.println("Workbooks variant is null!");
			return null;
		}
		
		OleAutomation workbooksAutomation = workbooksVariant.getAutomation();
		workbooksVariant.dispose();
			
		OleAutomation embeddedWorkbook = CollectionsUtils.getItemByName(workbooksAutomation, workbookName, false);
		workbooksAutomation.dispose();
		
		return embeddedWorkbook;	
	}
	
	
	/**
	 * Get the workbook OleAutomation using the "ThisWorkbook" property  
	 * @param application an OleAutomation that provides access to the functionalities of the (Excel) Application OLE object
	 * @return an OleAutomation that provides access to the functionalities of the This Workbook OLE object
	 */
	public static OleAutomation getThisWorkbookAutomation(OleAutomation application){
		
		int[] thisWorkbookIds = application.getIDsOfNames(new String[]{"ThisWorkbook"});	
		if (thisWorkbookIds == null){			
			System.out.println("\"ThisWorkbook\" property not found for \"Application\" object!");
			return null;
		}		
		
		Variant thisWorkbookVariant = application.getProperty(thisWorkbookIds[0]);
		if (thisWorkbookVariant == null) {
			System.out.println("ThisWorkbook variant is null!");
			return null;
		}
		
		OleAutomation workbookAutomation = thisWorkbookVariant.getAutomation();
		thisWorkbookVariant.dispose();
		
		return workbookAutomation;
	}
		
	/**
	 * Get the Worksheets automation from the "active" Workbook
	 * @param application an OleAutomation that provides access to the functionalities of the (Excel) Application OLE object 
	 * @return an OleAutomation that provides access to the functionalities of the Worksheets OLE object
	 */
	public static OleAutomation getWorksheetsAutomation(OleAutomation application){
		
		// get ID of Worksheets property
		int[] worksheetsObjectIds = application.getIDsOfNames(new String[]{"Worksheets"});
		if (worksheetsObjectIds == null) {
			System.out.println("Property \"Worksheets\" was not found for the given Application OLE object!");
			return null;
		}
		
		// get property using the ID 
		Variant worksheetsVariant =  application.getProperty(worksheetsObjectIds[0]);	
		if(worksheetsVariant == null){
			System.out.println("\"Worksheets\" variant is null!");
			return null;		
		}
		
		// get worksheets automation from the variant
		OleAutomation worksheetsAutomation = worksheetsVariant.getAutomation();
		worksheetsVariant.dispose();
		
		return worksheetsAutomation;
	}
	
	/**
	 * Get the active worksheet automation using the "ActiveSheet" property. 
	 * @param applicationAutomation an OleAutomation that provides access to the functionalities of the Excel (Application) OLE object
	 * @return an OleAutomation for the ActiveWorksheet
	 */
	public static OleAutomation getActiveWorksheetAutomation(OleAutomation applicationAutomation){
		
		int[] worksheetIds = applicationAutomation.getIDsOfNames(new String[]{"ActiveSheet"});	
		if (worksheetIds == null){			
			System.out.println("\"ActiveSheet\" property not found for the given OleAutomation object!");
			return null;
		}		
		Variant worksheetVariant = applicationAutomation.getProperty(worksheetIds[0]);
		if (worksheetVariant == null) {
			System.out.println("Worksheet variant is null!");
			return null;
		}		
		OleAutomation worksheetAutomation = worksheetVariant.getAutomation();
		worksheetVariant.dispose();
		
		return worksheetAutomation;
	}
	
	/**
	 * Get the OleAutomation for the WorksheetFunction
	 * @param applicationAutomation an OleAutomation that provides access to the functionalities of the Excel (Application) OLE object
	 * @return an OleAutomation that provides access to the functionalities of WorksheetFunction
	 */
	public static  OleAutomation getWorksheetFunctionAutomation(OleAutomation applicationAutomation){
		
		int[] wfIds = applicationAutomation.getIDsOfNames(new String[]{"WorksheetFunction"});	
		if (wfIds == null){			
			System.out.println("\"WorksheetFunction\" property not found for the given Application OleAutomation object!");
			return null;
		}		
		Variant worksheetFunctionVariant = applicationAutomation.getProperty(wfIds[0]);
		if (worksheetFunctionVariant == null) {
			System.out.println("WorksheetFunction variant is null!");
			return null;
		}		
		OleAutomation worksheetFunctionAutomation = worksheetFunctionVariant.getAutomation();
		worksheetFunctionVariant.dispose();
		
		return worksheetFunctionAutomation;
	}	
	
	
	/**
	 * Get the intersection between two ranges
	 * 
	 * @param applicationAutomation an OleAutomation object for accessing the (Excel) Application OLE object
	 * @param range1 an OleAutomation to access the functionalities of the first range 
	 * @param range2 an OleAutomation to access the functionalities of the second range
	 * @return an OleAutomation of a range object that represents the intersection between two ranges
	 */
	public static OleAutomation getIntersection(OleAutomation applicationAutomation, OleAutomation range1, OleAutomation range2){
		
		int[] intersectMethodIds = applicationAutomation.getIDsOfNames(new String[]{"Intersect"});
		if (intersectMethodIds == null){			
			System.out.println("Could not get the ids of the \"Intersect\" method using the given "
					+ "Application OleAutomation object! \n"+applicationAutomation.getLastError());
			return null;
		}
		
		Variant[] args = new Variant[2];
		args[0] = new Variant(range1);
		args[1] = new Variant(range2);
		
		Variant result =applicationAutomation.invoke(intersectMethodIds[0], args);
		
		if(result==null ){

			logger.error("Invoking \"Intersect\" method for \"Application\" returned null variant! "
					+ "\n"+applicationAutomation.getLastError());
			return null;
		}
		
		if(result.getType()==0){
			return null;
		}
		
		logger.debug("Invoking \"Intersect\" method for \"Application\" returned variant: "+result);
		OleAutomation intersection = result.getAutomation();
		result.dispose();
		
		return intersection;
	}
	
	
	/**
	 * Get the union of the two ranges
	 * 
	 * @param applicationAutomation an OleAutomation object for accessing the (Excel) Application OLE object
	 * @param range1 an OleAutomation to access the functionalities of the first range 
	 * @param range2 an OleAutomation to access the functionalities of the second range
	 * @return an OleAutomation of a range object that represents the union between two ranges
	 */
	public static OleAutomation getUnion(OleAutomation applicationAutomation, OleAutomation range1, OleAutomation range2){
		
		int[] unionMethodIds = applicationAutomation.getIDsOfNames(new String[]{"Union"});
		if (unionMethodIds == null){			
			System.out.println("Could not get the ids of the \"Union\" method using the given "
					+ "Application OleAutomation object! \n"+applicationAutomation.getLastError());
			return null;
		}
		
		Variant[] args = new Variant[2];
		args[0] = new Variant(range1);
		args[1] = new Variant(range2);
		
		Variant result =applicationAutomation.invoke(unionMethodIds[0], args);
		
		if(result==null ){

			logger.error("Invoking \"Union\" method for \"Application\" returned null variant! "
					+ "\n"+applicationAutomation.getLastError());
			return null;
		}
		
		if(result.getType()==0){
			return null;
		}
		
		logger.debug("Invoking \"Union\" method for \"Application\" returned variant: "+result);
		OleAutomation union = result.getAutomation();
		result.dispose();
		
		return union;
	}
	
	/**
	 * Set application alerts on or off 
	 * @param applicationAutomation an OleAutomation object for accessing the (Excel) Application OLE object
	 * @param display true to display alerts, false to suppress them 
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean setDisplayAlerts(OleAutomation applicationAutomation, boolean display){
		
		int[] displayAlertsPropertyIds = applicationAutomation.getIDsOfNames(new String[]{"DisplayAlerts"});
		Variant valueVariant = new Variant(display);
		boolean isSuccess = applicationAutomation.setProperty(displayAlertsPropertyIds[0], valueVariant);
		valueVariant.dispose();
		
		return isSuccess;
	}

		
	/**
	 * Set the value for the "ScreenUpdating" property 
	 * @param applicationAutomation an OleAutomation object for accessing the (Excel) Application OLE object
	 * @param update true to turn on updating for the screen, false to turn off 
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean setScreenUpdating(OleAutomation applicationAutomation, boolean update){
		
		int[] screenUpdatingPropertyIds = applicationAutomation.getIDsOfNames(new String[]{"ScreenUpdating"});
		Variant valueVariant = new Variant(update);
		boolean isSuccess = applicationAutomation.setProperty(screenUpdatingPropertyIds[0], valueVariant);
		valueVariant.dispose();
		
		return isSuccess;
	}
	
	
	/**
	 * Hide Ribbon from Excel GUI
	 * @param application an OleAutomation that provides access to the functionalities of the (Excel) Application OLE object
	 * @return
	 */
	public static boolean hideRibbon(OleAutomation application){
		
		int[] ee4mIds = application.getIDsOfNames(new String[]{"ExecuteExcel4Macro"});
		
		Variant[] parameters = new Variant[1];
	    parameters[0] = new Variant("SHOW.TOOLBAR(\"Ribbon\",False)");
	    
	    Variant result = application.invoke(ee4mIds[0], parameters);
	   
	    boolean isSuccess = false;
	    if(result!=null){
	    	isSuccess = true;
	    	result.dispose();
	    }
	    parameters[0].dispose();
	    
	    return isSuccess;
	}
	
	
	public static boolean setVisibilityStatusBar(OleAutomation application, boolean visible){
		int[] displayStatusBarMethodIds = application.getIDsOfNames(new String[]{"DisplayStatusBar"});
		return  application.setProperty(displayStatusBarMethodIds[0], new Variant(visible));
	}
	
	/**
	 * Quit Excel application. Use the given Application OleAutomation to invoke the "Quit" method. 
	 * @param application an OleAutomation that provides access to the functionalities of the (Excel) Application OLE object
	 */
	public static boolean quitExcelApplication(OleAutomation application){
		
		if(application==null){
			System.err.println("ERROR: Application is null!!!");
			return false;
		}
			
		int[] quitMethodIds = application.getIDsOfNames(new String[]{"Quit"});
		if (quitMethodIds == null){			
			System.err.println("\"Quit\" method not found for \"Application\" object!");
			return false;
		}	
		
		Variant result = application.invoke(quitMethodIds[0]);
		if(result==null){ 
			return false;
		}
		
		result.dispose();
		return true;
	}
}
