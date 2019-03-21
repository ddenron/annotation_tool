/**
 * 
 */
package de.tudresden.annotator.oleutils;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class ShadowFormatUtils {
	
	/**
	 * Set the shadow size
	 * @param shadowAutomation an OleAutomation object that provides access to the ShadowFormat object functionalities.
	 * @param size an integer value that represents the size of the shadow
	 * @return true if the operation was successful, false otherwise
	 */	
	public static boolean setShadowSize(OleAutomation shadowAutomation, int size ){
		
		int sizePropertyIds[] = shadowAutomation.getIDsOfNames(new String[] {"Size"});
		Variant sizeVariant = new Variant(size); 
		boolean isSuccess = shadowAutomation.setProperty(sizePropertyIds[0], sizeVariant);
		sizeVariant.dispose();
		
		return isSuccess;	
	}
	
	/**
	 * Set the shadow OffsetX
	 * @param shadowAutomation an OleAutomation object that provides access to the ShadowFormat object functionalities.
	 * @param offsetX an integer value that represents the X offset of the shadow from the center 
	 * @return true if the operation was successful, false otherwise
	 */	
	public static boolean setShadowOffsetX(OleAutomation shadowAutomation, int offsetX ){
		
		int offsetXPropertyIds[] = shadowAutomation.getIDsOfNames(new String[] {"OffsetX"});
		Variant offsetXVariant = new Variant(offsetX); 
		boolean isSuccess = shadowAutomation.setProperty(offsetXPropertyIds[0], offsetXVariant);
		offsetXVariant.dispose();
		
		return isSuccess;	
	}
	
	/**
	 * Set the shadow OffsetY
	 * @param shadowAutomation an OleAutomation object that provides access to the ShadowFormat object functionalities.
	 * @param offsetY an integer value that represents the Y offset of the shadow from the center 
	 * @return true if the operation was successful, false otherwise
	 */	
	public static boolean setShadowOffsetY(OleAutomation shadowAutomation, int offsetY ){
		
		int offsetYPropertyIds[] = shadowAutomation.getIDsOfNames(new String[] {"OffsetY"});
		Variant offsetYVariant = new Variant(offsetY); 
		boolean isSuccess = shadowAutomation.setProperty(offsetYPropertyIds[0], offsetYVariant);
		offsetYVariant.dispose();
		
		return isSuccess;	
	}
	
	
	/**
	 * Set the shadow style
	 * @param shadowAutomation an OleAutomation object that provides access to the ShadowFormat object functionalities.
	 * @param style one of the MsoShadowStyle enumeration values
	 * @return true if the operation was successful, false otherwise
	 */	
	public static boolean setShadowStyle(OleAutomation shadowAutomation, int style ){
		
		int stylePropertyIds[] = shadowAutomation.getIDsOfNames(new String[] {"Style"});
		Variant styleVariant = new Variant(style); 
		boolean isSuccess = shadowAutomation.setProperty(stylePropertyIds[0], styleVariant);
		styleVariant.dispose();
		
		return isSuccess;	
	}
	
	
	/**
	 * Set the shadow type
	 * @param shadowAutomation an OleAutomation object that provides access to the ShadowFormat object functionalities.
	 * @param type one of the MsoShadowType enumeration values
	 * @return true if the operation was successful, false otherwise
	 */	
	public static boolean setShadowType(OleAutomation shadowAutomation, int type ){
		
		int typePropertyIds[] = shadowAutomation.getIDsOfNames(new String[] {"Type"});
		Variant typeVariant = new Variant(type); 
		boolean isSuccess = shadowAutomation.setProperty(typePropertyIds[0], typeVariant);
		typeVariant.dispose();
		
		return isSuccess;	
	}
	
	
	/**
	 * Set shadow blur
	 * @param shadowAutomation an OleAutomation object that provides access to the ShadowFormat object functionalities.
	 * @param blur an integer value that represents the line blur in points
	 * @return true if the operation was successful, false otherwise
	 */	
	public static boolean setShadowBlur(OleAutomation shadowAutomation, int blur ){
		
		int blurPropertyIds[] = shadowAutomation.getIDsOfNames(new String[] {"Blur"});
		Variant blurVariant = new Variant(blur); 
		boolean isSuccess = shadowAutomation.setProperty(blurPropertyIds[0], blurVariant);
		blurVariant.dispose();
		
		return isSuccess;	
	}
	
	
	/**
	 * Set the transparency of the shadow 
	 * @param shadowAutomation an OleAutomation object that provides access to the ShadowFormat object functionalities.
	 * @param transparency a double value that represents the line transparency
	 * @return true if the operation was successful, false otherwise
	 */	
	public static boolean setShadowTransparency(OleAutomation shadowAutomation, double transparency ){
		
		int transparencyPropertyIds[] = shadowAutomation.getIDsOfNames(new String[] {"Transparency"});
		Variant transparencyVariant = new Variant(transparency); 
		boolean isSuccess = shadowAutomation.setProperty(transparencyPropertyIds[0], transparencyVariant);
		transparencyVariant.dispose();
		
		return isSuccess;	
	}
	
	
	/**
	 * Set if shadow is visible or not
	 * @param shadowAutomation an OleAutomation object that provides access to the ShadowFormat object functionalities.
	 * @param visible true if the shadow should be visible, false if the shadow should be hidden 
	 * @return true if the operation was successful, false otherwise
	 */	
	public static boolean setShadowVisibility(OleAutomation shadowAutomation, boolean visible ){
		
		int visiblePropertyIds[] = shadowAutomation.getIDsOfNames(new String[] {"Visible"});
		Variant visibleVariant = new Variant(visible); 
		boolean isSuccess = shadowAutomation.setProperty(visiblePropertyIds[0], visibleVariant);
		visibleVariant.dispose();
		
		return isSuccess;	
	}
}
