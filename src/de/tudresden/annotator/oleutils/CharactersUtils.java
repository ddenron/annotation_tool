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
public class CharactersUtils {
	
	private static final Logger logger = LogManager.getLogger(CharactersUtils.class.getName());
			
	/**
	 * Set Text using a Characters OleAutomation
	 * @param charactersAutomation an OleAutomation of a Characters object, which represents a range of characters within a shape’s text frame
	 * @param text the string to set as text
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean setText(OleAutomation charactersAutomation, String text){	
		
		logger.debug("Is character oleautomation null? ".concat(String.valueOf(charactersAutomation==null)));
		
		int[] textPropertyIds = charactersAutomation.getIDsOfNames(new String[]{"Text"}); 
		if(textPropertyIds==null)
			logger.error("Could not get id of property \"Text\" for \"Character\" ole object");
		
		logger.debug("The text to set is: "+text);
		return charactersAutomation.setProperty(textPropertyIds[0], new Variant(text));
	}
	
	/**
	 * Get Font OleObject. This object can be used to change the font attributes of the text
	 * @param charactersAutomation an OleAutomation of a Characters object, which represents a range of characters within a shape’s text frame
	 * @return an OleAutomation that provides access to the Font Ole object. 
	 */
	public static OleAutomation getFontAutomation(OleAutomation charactersAutomation){
		
		logger.debug("Is character oleautomation null? ".concat(String.valueOf(charactersAutomation==null)));
		
		int[] fontPropertyIds = charactersAutomation.getIDsOfNames(new String[]{"Font"});
		if(fontPropertyIds==null)
			logger.error("Could not get id of property \"Font\" for \"Character\" ole object");
		
		Variant fontVariant = charactersAutomation.getProperty(fontPropertyIds[0]);
		logger.debug("Invoking get property \"Font\" for \"Character\" ole object returned variant: "+fontVariant);
		
		if(fontVariant==null)
			logger.error("Invoking get property \"Font\" for \"Character\" ole object returned null variant ");

		OleAutomation fontAutomation = fontVariant.getAutomation();
		fontVariant.dispose();
		
		return fontAutomation;
	}
}
