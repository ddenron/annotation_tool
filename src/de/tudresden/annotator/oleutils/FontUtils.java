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
public class FontUtils {
	
	private static final Logger logger = LogManager.getLogger(FontUtils.class.getName());
	
	/**
	 * Set text font size
	 * @param fontAutomation an OleAutomation that provides access to a Font OLE Object
	 * @param size the font size
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean setFontSize(OleAutomation fontAutomation, int size){	
		
		logger.debug("Is Font oleautomation null? ".concat(String.valueOf(fontAutomation==null)));
		
		int[] sizePropertyIds = fontAutomation.getIDsOfNames(new String[]{"Size"}); 
		if(sizePropertyIds==null)
			logger.error("Could not get id of property \"Size\" for \"Font\" ole object");
		
		logger.debug("The value to set for property \"Size\" of \"FillFormat\" ole object is: "+size);
		return fontAutomation.setProperty(sizePropertyIds[0], new Variant(size));
	}
	
	/**
	 * Make text bold 
	 * @param fontAutomation an OleAutomation that provides access to a Font OLE Object
	 * @param bold true to make text bold, false otherwise 
	 * @return true if operation was successful, false if an error was encountered 
	 */
	public static boolean setBoldFont(OleAutomation fontAutomation, boolean bold){		
		logger.debug("Is Font oleautomation null? ".concat(String.valueOf(fontAutomation==null)));
		
		int[] boldPropertyIds = fontAutomation.getIDsOfNames(new String[]{"Bold"});
		if(boldPropertyIds==null)
			logger.error("Could not get id of property \"Bold\" for \"Font\" ole object");
		
		logger.debug("The value to set for property \"Bold\" of \"FillFormat\" ole object is: "+String.valueOf(bold));
		return fontAutomation.setProperty(boldPropertyIds[0], new Variant(bold));
	}
	
	/**
	 * Set font color for the text
	 * @param fontAutomation an OleAutomation that provides access to a Font OLE Object
	 * @param color a long that represents a RGB color. Is calculated as B * 65536 + G * 256 + R
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean setFontColor(OleAutomation fontAutomation, long color){	
		logger.debug("Is Font oleautomation null? ".concat(String.valueOf(fontAutomation==null)));
		
		int[] longPropertyIds = fontAutomation.getIDsOfNames(new String[]{"Color"}); 
		if(longPropertyIds==null)
			logger.error("Could not get id of property \"Color\" for \"Font\" ole object");
		
		logger.debug("The value to set for property \"Color\" of \"FillFormat\" ole object is: "+color);
		return fontAutomation.setProperty(longPropertyIds[0], new Variant(color));
	}

}
