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
public class TextFrameUtils {
	
	private static final Logger logger = LogManager.getLogger(TextFrameUtils.class.getName());
			
	/**
	 * Set text horizontal alignment in the shape TextFrame  
	 * @param textFrameAutomation an OleAutomation that provides access to a TextFrame Ole object. It represents a textframe in shape object.
	 * @param alignment one of the XlVAlign constants
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setVerticalAlignment(OleAutomation textFrameAutomation, int alignment){
		
		logger.debug("Is TextFrame oleautomation null? ".concat(String.valueOf(textFrameAutomation==null)));
		
		int[] verticalAlignmentPropertyIds = textFrameAutomation.getIDsOfNames(new String[]{"VerticalAlignment"});
		if(verticalAlignmentPropertyIds==null)
			logger.error("Could not get id of property \"VerticalAlignment\" for \"TextFrame\" ole object");
		
		logger.debug("The value to set for property \"VerticalAlignment\" of \"TextFrame\" ole object is: "+alignment);
		return textFrameAutomation.setProperty(verticalAlignmentPropertyIds[0], new Variant(alignment));
	}
	
	/**
	 * Set text vertical alignment in the shape TextFrame
	 * @param textFrameAutomation an OleAutomation that provides access to a TextFrame Ole object. It represents a textframe in shape object.
	 * @param alignment one of the XlHAlign constants
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setHorizontalAlignment(OleAutomation textFrameAutomation, int alignment){
		
		logger.debug("Is TextFrame oleautomation null? ".concat(String.valueOf(textFrameAutomation==null)));
		
		int[] horizontalAlignmentPropertyIds = textFrameAutomation.getIDsOfNames(new String[]{"HorizontalAlignment"});
		if(horizontalAlignmentPropertyIds==null)
			logger.error("Could not get id of property \"HorizontalAlignment\" for \"TextFrame\" ole object");
		
		logger.debug("The value to set for property \"HorizontalAlignment\" of \"TextFrame\" ole object is: "+alignment);
		return textFrameAutomation.setProperty(horizontalAlignmentPropertyIds[0], new Variant(alignment));
	}

	/**
	 * Get Characters OleAutomation. Use this object to set and format text in a TextFrame 
	 * @param textFrameAutomation an OleAutomation that provides access to a TextFrame Ole object. It represents a textframe in shape object.
	 * @return an OleAutomation that provides access to a Characters object.
	 */
	public static OleAutomation getCharactersAutomation(OleAutomation textFrameAutomation){
		
		logger.debug("Is TextFrame oleautomation null? ".concat(String.valueOf(textFrameAutomation==null)));
		
		int[] charactersMethodIds = textFrameAutomation.getIDsOfNames(new String[]{"Characters"});
		if(charactersMethodIds==null)
			logger.error("Could not get ids of method \"Characters\" for \"TextFrame\" ole object");
		
		Variant charactersVariant = textFrameAutomation.invoke(charactersMethodIds[0]);
		logger.debug("Invoking method \"Characters\" of \"TextFrame\" ole object returned variant: "+charactersVariant);
		
		if(charactersVariant==null){
			logger.error("Invoking method \"Characters\" of \"TextFrame\" ole object returned null variant!");
		}
		
		OleAutomation charactersAutomation = charactersVariant.getAutomation();
		charactersVariant.dispose();
		
		return charactersAutomation;
	}

}
