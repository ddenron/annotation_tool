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
public class ColorFormatUtils {
	
	private static final Logger logger = LogManager.getLogger(ColorFormatUtils.class.getName());
			
	/**
	 * Set the fore color of the shape fill 
	 * @param automation an OleAutomation that has the "ForeColor" property
	 * @param color a long that represents a RGB color. Is calculated as B * 65536 + G * 256 + R
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setForeColor(OleAutomation automation, long color){
		
		logger.debug("Is oleautomation null? ".concat(String.valueOf(automation==null)));
		
		int[] foreColorPropertyIds = automation.getIDsOfNames(new String[]{"ForeColor"}); 
		
		if(foreColorPropertyIds==null)	{		
			logger.error("Could not retrieve id of property \"ForeColor\"");
			return false;
		}
		
		Variant foreColorVariant = automation.getProperty(foreColorPropertyIds[0]);
		logger.debug("Invoking get property \"ForeColor\" returned variant: "+foreColorVariant);
		OleAutomation foreColorAutomation = foreColorVariant.getAutomation();

		int[] rgbPropertyIds = foreColorAutomation.getIDsOfNames(new String[]{"RGB"}); //alternatively use "SchemeColor" 
		if(rgbPropertyIds==null)	{		
			logger.error("Could not retrieve id of property \"RGB\" for \"ForeColor\" ole object");
			return false;
		}
		
		logger.debug("The BackColor to set is: "+color);
		
		boolean wasColorUpdated = foreColorAutomation.setProperty(rgbPropertyIds[0], new Variant(color)); 
		logger.debug("Invoking set property \"RGB\" for \"ForeColor\" ole object returned: "+wasColorUpdated);
		
		boolean isSuccess = automation.setProperty(foreColorPropertyIds[0], foreColorVariant);			
		logger.debug("Invoking set property \"ForeColor\" returned: "+isSuccess);
		
		foreColorVariant.dispose();
		foreColorAutomation.dispose();
		
		return isSuccess;
	}
	
	
	/**
	 * Set the back color of the shape fill 
	 * @param automation an OleAutomation that has the "BackColor" property
	 * @param colorIndex a long that represents a RGB color. Is calculated as B * 65536 + G * 256 + R 
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setBackColor(OleAutomation automation,  long color){
	
		logger.debug("Is oleautomation null? ".concat(String.valueOf(automation==null)));
		
		int[] backColorPropertyIds = automation.getIDsOfNames(new String[]{"BackColor"}); 
		
		if(backColorPropertyIds==null)	{		
			logger.error("Could not retrieve id of property \"BackColor\"");
			return false;
		}
		
		Variant backColorVariant = automation.getProperty(backColorPropertyIds[0]);
		logger.debug("Invoking get property \"BackColor\" returned variant: "+backColorVariant);
		OleAutomation backColorAutomation = backColorVariant.getAutomation();
	
		int[] rgbPropertyIds = backColorAutomation.getIDsOfNames(new String[]{"RGB"}); //alternatively use "SchemeColor" 
		if(rgbPropertyIds==null)	{		
			logger.error("Could not retrieve id of property \"RGB\" for \"BackColor\" ole object");
			return false;
		}
		
		logger.debug("The BackColor to set is: "+color);
		
		boolean wasColorUpdated = backColorAutomation.setProperty(rgbPropertyIds[0], new Variant(color)); 
		logger.debug("Invoking set property \"RGB\" for \"BackColor\" ole object returned: "+wasColorUpdated);
		
		boolean isSuccess = automation.setProperty(backColorPropertyIds[0], backColorVariant);
		logger.debug("Invoking set property \"BackColor\" returned: "+isSuccess);
		
		backColorVariant.dispose();
		backColorAutomation.dispose();
		
		return isSuccess;
	}
	
	/**
	 * Get RGB color as long value
	 * @param red value of red color
	 * @param green value of green color
	 * @param blue value of blue color
	 * @return a long value that represents an RGB color 
	 */
	public static long getRGBColorAsLong(int red, int green, int blue){	
		return blue * 65536 + green * 256 + red;
	}
}
