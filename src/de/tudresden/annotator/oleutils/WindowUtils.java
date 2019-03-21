/**
 * 
 */
package de.tudresden.annotator.oleutils;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class WindowUtils {
	
	private static final Logger logger = LogManager.getLogger(WindowUtils.class.getName());
			
	/**
	 * Sets the number of the row that appears at the top of the pane or window 
	 * @param windowAutomation an OleAutomation that provides access to the functionalities of the excel window
	 * @return true if the property was successfully updated, false otherwise 
	 */
	public static boolean setScrollRow(OleAutomation windowAutomation, int row){
		
		if(windowAutomation==null){
			logger.error("Method setScrollRow received null windowAutomation object");
		}
		
		int[] scrollRowPropertyIds = windowAutomation.getIDsOfNames(new String[]{"ScrollRow"});	
		
		if(scrollRowPropertyIds==null){
			logger.error("Could not get \"ScrollRow\" property ids for \"Window\" ole object!");
		}
		
		Variant rowNum = new Variant(row);
		logger.debug("The value to set for \"ScrollRow\" property is "+row);
		
		boolean isSuccess = windowAutomation.setProperty(scrollRowPropertyIds[0], rowNum);
		logger.debug("Invoking set value for \"ScrollRow\" property returned "+isSuccess);
		
		rowNum.dispose();
		return isSuccess;
	}
	
	
	/**
	 * Sets the number of the column that appears on the leftmost of the pane or window 
	 * @param windowAutomation an OleAutomation that provides access to the functionalities of the excel window
	 * @return true if the property was successfully updated, false otherwise 
	 */
	public static boolean setScrollColumn(OleAutomation windowAutomation, int column){
		
		if(windowAutomation==null){
			logger.error("Method setScrollColumn received null windowAutomation object");
		}
		
		int[] scrollColumnPropertyIds = windowAutomation.getIDsOfNames(new String[]{"ScrollColumn"});	
		
		if(scrollColumnPropertyIds==null){
			logger.error("Could not get \"ScrollColumn\" property ids for \"Window\" ole object!");
		}
		
		Variant colNum = new Variant(column);
		logger.debug("The value to set for \"ScrollColumn\" property is "+column);
		
		boolean isSuccess = windowAutomation.setProperty(scrollColumnPropertyIds[0], colNum);
		logger.debug("Invoking set value for \"ScrollColumn\" property returned "+isSuccess);
		
		colNum.dispose();
		return isSuccess;
	}

	
	/**
	 * Display or hide formulas in the active sheet
	 * @param windowAutomation an OleAutomation that provides access to the functionalities of the excel window
	 * @return true if the property was successfully updated, false otherwise 
	 */
	public static boolean setDisplayFormulas(OleAutomation windowAutomation, boolean display){
		
		if(windowAutomation==null){
			logger.error("Method setDisplayFormulas received null windowAutomation object");
		}
		
		int[] displayFormulasPropertyIds = windowAutomation.getIDsOfNames(new String[]{"DisplayFormulas"});	
		
		if(displayFormulasPropertyIds==null){
			logger.error("Could not get \"DisplayFormulas\" property ids for \"Window\" ole object!");
		}
		
		Variant valueVariant = new Variant(display);
		logger.debug("The value to set for \"DisplayFormulas\" property is "+display);
		
		boolean isSuccess = windowAutomation.setProperty(displayFormulasPropertyIds[0], valueVariant);
		logger.debug("Invoking set value for \"DisplayFormulas\" property returned "+isSuccess);
		
		valueVariant.dispose();
		return isSuccess;
	}
		
	
	/**
	 * Get the value of the "DisplayFormulas" for the Window ole object
	 * @param windowAutomation an OleAutomation that provides access to the functionalities of the excel window
	 * @return true if the property was successfully updated, false otherwise 
	 */
	public static boolean getDisplayFormulas(OleAutomation windowAutomation){
		
		if(windowAutomation==null){
			logger.error("Method getDisplayFormulas received null windowAutomation object");
		}
		
		int[] displayFormulasPropertyIds = windowAutomation.getIDsOfNames(new String[]{"DisplayFormulas"});	
		
		if(displayFormulasPropertyIds==null){
			logger.error("Could not get \"DisplayFormulas\" property ids for \"Window\" ole object!");
		}
			
		Variant result= windowAutomation.getProperty(displayFormulasPropertyIds[0]);	
		if(result==null){
			logger.error("Invoking get value for \"DisplayFormulas\" property returned null variant ");
		}
		
		logger.debug("Invoking get value for \"DisplayFormulas\" property returned variant: "+result);
		
		boolean areFormulasVisible = result.getBoolean();
		result.dispose();
		
		return areFormulasVisible;
	}
	
	/**
	 * Get the zoom level for the window
	 * @param windowAutomation an OleAutomation that provides access to the functionalities of the excel window
	 * @return an integer that represents the current zoom level for the window
	 */
	public static int getZoom(OleAutomation windowAutomation){
		
		if(windowAutomation==null){
			logger.error("Method getDisplayFormulas received null windowAutomation object");
		}
		
		int[] zooomPropertyIds = windowAutomation.getIDsOfNames(new String[]{"Zoom"});	
		
		if(zooomPropertyIds==null){
			logger.error("Could not get \"Zoom\" property ids for \"Window\" ole object!");
		}
			
		Variant result= windowAutomation.getProperty(zooomPropertyIds[0]);	
		if(result==null){
			logger.error("Invoking get value for \"Zoom\" property returned null variant ");
		}
		
		logger.debug("Invoking get value for \"Zoom\" property returned variant: "+result);
		
		int zoomLevel = result.getInt();
		result.dispose();
		
		return zoomLevel;
	}
		
	/**
	 * Set the zoom level for the window
	 * @param windowAutomation an OleAutomation that provides access to the functionalities of the excel window
	 * @return true if the property was successfully updated, false otherwise 
	 */
	public static boolean setZoom(OleAutomation windowAutomation, int zoom){
		
		if(windowAutomation==null){
			logger.error("Method getDisplayFormulas received null windowAutomation object");
		}
		
		int[] zooomPropertyIds = windowAutomation.getIDsOfNames(new String[]{"Zoom"});	
		
		if(zooomPropertyIds==null){
			logger.error("Could not get \"Zoom\" property ids for \"Window\" ole object!");
		}
		
		Variant arg = new Variant(zoom);
		
		Boolean result= windowAutomation.setProperty(zooomPropertyIds[0], arg);	
		
		logger.debug("Invoking set value for \"Zoom\" property returned: "+result);
		
		arg.dispose();
		return result;
	}
}
