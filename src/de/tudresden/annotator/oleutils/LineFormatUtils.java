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
public class LineFormatUtils {
	
	private static final Logger logger = LogManager.getLogger(LineFormatUtils.class.getName());
			
	/**
	 * Set the line style
	 * @param lineFormatAuto an OleAutomation object that provides access to the LineFormat object functionalities.
	 * @param style one of the constant values from MsoLineStyle enumeration 
	 * @return true if the operation was successful, false otherwise
	 */	
	public static boolean setLineStyle(OleAutomation lineFormatAuto, int style ){
		
		logger.debug("Is LineFormat oleautomation null? ".concat(String.valueOf(lineFormatAuto==null)));
		
		int stylePropertyIds[] = lineFormatAuto.getIDsOfNames(new String[] {"Style"});
		if(stylePropertyIds==null)
			logger.error("Could not get id of property \"Style\" for \"LineFormat\" ole object");
		
		Variant styleVariant = new Variant(style); 	
		logger.debug("The value to set for property \"Style\" of \"LineFormat\" ole object is: "+style);
		
		boolean isSuccess = lineFormatAuto.setProperty(stylePropertyIds[0], styleVariant);
		logger.debug("Invoking set property \"Style\" of \"LineFormat\" ole object returned: "+String.valueOf(isSuccess));
		
		styleVariant.dispose();		
		return isSuccess;	
	}
	
	/**
	 * Set the weight of the line 
	 * @param lineFormatAuto an OleAutomation object that provides access to the LineFormat object functionalities.
	 * @param weight a double value that represents the weight of the line 
	 * @return true if the operation was successful, false otherwise
	 */	
	public static boolean setLineWeight(OleAutomation lineFormatAuto, double weight ){
		
		logger.debug("Is LineFormat oleautomation null? ".concat(String.valueOf(lineFormatAuto==null)));
		
		int weightPropertyIds[] = lineFormatAuto.getIDsOfNames(new String[] {"Weight"});
		if(weightPropertyIds==null)
			logger.error("Could not get id of property \"Weight\" for \"LineFormat\" ole object");
		
		Variant weightVariant = new Variant(weight); 
		logger.debug("The value to set for property \"Weight\" of \"LineFormat\" ole object is: "+weight);
		
		boolean isSuccess = lineFormatAuto.setProperty(weightPropertyIds[0], weightVariant);
		logger.debug("Invoking set property \"Weight\" of \"LineFormat\" ole object returned: "+String.valueOf(isSuccess));
		
		weightVariant.dispose();
		return isSuccess;	
	}
	
	/**
	 * Set if line is visible or not
	 * @param lineFormatAuto an OleAutomation object that provides access to the LineFormat object functionalities.
	 * @param visible true if the line should be visible, false if the line should be hidden 
	 * @return true if the operation was successful, false otherwise
	 */	
	public static boolean setLineVisibility(OleAutomation lineFormatAuto, boolean visible ){
		
		logger.debug("Is LineFormat oleautomation null? ".concat(String.valueOf(lineFormatAuto==null)));
		
		int visiblePropertyIds[] = lineFormatAuto.getIDsOfNames(new String[] {"Visible"});
		if(visiblePropertyIds==null)
			logger.error("Could not get id of property \"Visible\" for \"LineFormat\" ole object");
		
		Variant visibleVariant = new Variant(visible); 
		logger.debug("The value to set for property \"Visible\" of \"LineFormat\" ole object is: "+String.valueOf(visible));
		
		boolean isSuccess = lineFormatAuto.setProperty(visiblePropertyIds[0], visibleVariant);
		logger.debug("Invoking set property \"Visible\" of \"LineFormat\" ole object returned: "+String.valueOf(isSuccess));
		
		visibleVariant.dispose();
		return isSuccess;	
	}
	
	
	/**
	 * Set the transparency of the line 
	 * @param lineFormatAuto an OleAutomation object that provides access to the LineFormat object functionalities.
	 * @param transparency a double value that represents the line transparency
	 * @return true if the operation was successful, false otherwise
	 */	
	public static boolean setLineTransparency(OleAutomation lineFormatAuto, double transparency ){
		
		logger.debug("Is LineFormat oleautomation null? ".concat(String.valueOf(lineFormatAuto==null)));
		
		int transparencyPropertyIds[] = lineFormatAuto.getIDsOfNames(new String[] {"Transparency"});
		if(transparencyPropertyIds==null)
			logger.error("Could not get id of property \"Transparency\" for \"LineFormat\" ole object");
		
		Variant transparencyVariant = new Variant(transparency); 
		logger.debug("The value to set for property \"Transparency\" of \"LineFormat\" ole object is: "+transparency);
		
		boolean isSuccess = lineFormatAuto.setProperty(transparencyPropertyIds[0], transparencyVariant);
		logger.debug("Invoking set property \"Transparency\" of \"LineFormat\" ole object returned: "+String.valueOf(isSuccess));
		
		transparencyVariant.dispose();
		return isSuccess;	
	}
}
