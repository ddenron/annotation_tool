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
public class FillFormatUtils {
	
	private static final Logger logger = LogManager.getLogger(FillFormatUtils.class.getName());
	
	/**
	 * Show or hide the fill of the shape.
	 * @param fillFormatAutomation an OleAutomation that provides access to the FillFormat OLE Object
	 * @param visible true to show the fill, false to hide it
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setFillVisibility(OleAutomation fillFormatAutomation, boolean visible){
		
		logger.debug("Is FillFormat oleautomation null? ".concat(String.valueOf(fillFormatAutomation==null)));
		
		int[] visiblePropertyIds = fillFormatAutomation.getIDsOfNames(new String[]{"Visible"});
		if(visiblePropertyIds==null)
			logger.error("Could not get id of property \"Visible\" for \"FillFormat\" ole object");
		
		logger.debug("The value to set for property \"Visible\" of \"FillFormat\" ole object is: "+String.valueOf(visible));
		Variant visiblePropertyVariant = new Variant(visible); 
		
		boolean isSuccess = fillFormatAutomation.setProperty(visiblePropertyIds[0], visiblePropertyVariant);
		logger.debug("Invoking set property \"Visible\" for \"FillFormat\" ole object returned : "+String.valueOf(isSuccess));
		visiblePropertyVariant.dispose();
	
		return isSuccess;
	}
	
	
	/**
	 * Set the transparency of the shape fill
	 * @param fillFormatAutomation an OleAutomation that provides access to the FillFormat OLE Object
	 * @param transparency as a double. The expected values are between 0 and 1.
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setFillTransparency(OleAutomation fillFormatAutomation, double transparency){
		
		logger.debug("Is FillFormat oleautomation null? ".concat(String.valueOf(fillFormatAutomation==null)));
		
		int[] transparencyPropertyIds = fillFormatAutomation.getIDsOfNames(new String[]{"Transparency"}); 
		if(transparencyPropertyIds==null)
			logger.error("Could not get id of property \"Transparency\" for \"FillFormat\" ole object");
		
		logger.debug("The value to set for property \"Transparency\" of \"FillFormat\" ole object is: "+String.valueOf(transparency));
		return fillFormatAutomation.setProperty(transparencyPropertyIds[0], new Variant(transparency));	
	}
	
	/**
	 * Set the pattern of the shape fill
	 * @param fillFormatAutomation an OleAutomation that provides access to the FillFormat OLE Object
	 * @param patternID as a integer. Check the MsoPatternType Enumeration
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setFillPattern(OleAutomation fillFormatAutomation, int patternID){
		
		logger.debug("Is FillFormat oleautomation null? ".concat(String.valueOf(fillFormatAutomation==null)));
		
		int[] patternPropertyIds = fillFormatAutomation.getIDsOfNames(new String[]{"Pattern"}); 
		if(patternPropertyIds==null)
			logger.error("Could not get id of property \"Pattern\" for \"FillFormat\" ole object");
		
		logger.debug("The value to set for property \"Pattern\" of \"FillFormat\" ole object is: "+String.valueOf(patternID));
		return fillFormatAutomation.setProperty(patternPropertyIds[0], new Variant(patternID));	
	}
}
