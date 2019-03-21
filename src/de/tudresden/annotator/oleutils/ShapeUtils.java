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
public class ShapeUtils {
	
	private static final Logger logger = LogManager.getLogger(ShapeUtils.class.getName());
	
	/**
	 * Draw a textbox at the specified location 
	 * @param shapesAutomation an OleAutomation that provides access to the "Shapes" Ole object. It represents a collection of shapes.
	 * @param left the distance, in points, from the left edge of column A to the left edge of the shape.
	 * @param top the distance, in points, from the top edge of row 1 to the top edge of the shape.
	 * @param width the width, in units, of the shape.
	 * @param height the height, in units, of the shape.
	 * @return an OleAutomation that provides access to the functionalities of the textbox that was just created
	 */
	public static OleAutomation drawTextBox(OleAutomation shapesAutomation, double left, double top, double width, double height){
		
		logger.debug("Is Shapes OleAutomation null? ".concat(String.valueOf(shapesAutomation==null)));
		
		int[] addTextboxMethodIds = shapesAutomation.getIDsOfNames(new String[]{"AddTextbox", "Orientation", "Left", "Top", "Width", "Height"}); 
		if(addTextboxMethodIds==null)
			logger.error("Could not get ids of the method \"AddTextbox\" for the \"Shapes\" ole object");
				
		Variant methodParams[] = new Variant[5];
		methodParams[0] = new Variant(1);
		methodParams[1] = new Variant(left); 
		methodParams[2] = new Variant(top); 
		methodParams[3] = new Variant(width); 
		methodParams[4] = new Variant(height);	
		
		Variant textboxVariant = shapesAutomation.invoke(addTextboxMethodIds[0], methodParams);
		logger.debug("Invoking the method \"AddTextbox\" for \"Shapes\" ole object returned variant: "+textboxVariant);
		
		for (Variant v : methodParams) {
			v.dispose();
		}
		
		if(textboxVariant==null){
			logger.error("Invoking the method \"AddTextbox\" for \"Shapes\" ole object returned a null variant");
			return null;
		}
		
		OleAutomation textboxAutomation = null;
		textboxAutomation = textboxVariant.getAutomation();
		textboxVariant.dispose();

		return textboxAutomation;
	}
	
	/**
	 * Create a shape at the specified location
	 * @param shapesAutomation an OleAutomation that provides access to the "Shapes" Ole object. It represents a collection of shapes.
	 * @param msoAutoShapeType the type of AutoShape to create
	 * @param left the distance, in points, from the left edge of column A to the left edge of the shape.
	 * @param top the distance, in points, from the top edge of row 1 to the top edge of the shape.
	 * @param width the width, in units, of the shape.
	 * @param height the height, in units, of the shape.
	 * @return an OleAutomation that provides access to the functionalities of the shape that was just created
	 */
	public static OleAutomation drawShape(OleAutomation shapesAutomation, int msoAutoShapeType, double left, double top, double width, double height){
		
		logger.debug("Is Shapes OleAutomation null? ".concat(String.valueOf(shapesAutomation==null)));
		
		int[] addShapeMethodIds = shapesAutomation.getIDsOfNames(new String[]{"AddShape", "Type", "Left", "Top", "Width", "Height"}); 
		if(addShapeMethodIds==null)
			logger.error("Could not get ids of the method \"AddShape\" for the \"Shapes\" ole object");
		
		Variant methodParams[] = new Variant[5];
		methodParams[0] = new Variant(msoAutoShapeType);
		methodParams[1] = new Variant(left); 
		methodParams[2] = new Variant(top); 
		methodParams[3] = new Variant(width); 
		methodParams[4] = new Variant(height);	
		
		Variant shapeVariant = shapesAutomation.invoke(addShapeMethodIds[0], methodParams);
		logger.debug("Invoking the method \"AddShape\" for \"Shapes\" ole object returned variant: "+shapeVariant);
		
		for (Variant v : methodParams) {
			v.dispose();
		}
		
		OleAutomation shapeAutomation = null;
		if(shapeVariant!=null){
			shapeAutomation = shapeVariant.getAutomation();
			shapeVariant.dispose();
		}else{
			logger.error("Invoking the method \"AddShape\" for \"Shapes\" ole object returned null variant");
		}
		
		return shapeAutomation;	
	}

	
	/**
	 * Get the title of the given shape		
	 * @param shapeAutomation an OleAutomation that provides access to the "Shape" Ole object. It represents a single shape.
	 * @return a string the represents the title of the shape
	 */
	public static String getShapeTitle(OleAutomation shapeAutomation){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] titlePropertyIds = shapeAutomation.getIDsOfNames(new String[]{"Title"}); 
		
		if(titlePropertyIds==null)
			logger.error("Could not get id of property \"Title\" for \"Shape\" ole object");
		
		Variant titleVariant = shapeAutomation.getProperty(titlePropertyIds[0]);
		logger.debug("Invoking get property \"Title\" for \"Shape\" ole object returned variant: "+titleVariant);
		
		String title = titleVariant.getString();
		titleVariant.dispose();

		return title;
	}
	
	
	/**
	 * Set title to the given shape		
	 * @param shapeAutomation an OleAutomation that provides access to the "Shape" Ole object. It represents a single shape.
	 * @param title a string that represents the title to set
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean setShapeTitle(OleAutomation shapeAutomation, String title){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] titlePropertyIds = shapeAutomation.getIDsOfNames(new String[]{"Title"}); 
		if(titlePropertyIds==null)
			logger.error("Could not get id of property \"Title\" for \"Shape\" ole object");
		
		Variant titleVariant = new Variant(title);
		logger.debug("The title to set for \"Shape\" ole object is: "+title);
		
		boolean isSuccess = shapeAutomation.setProperty(titlePropertyIds[0], titleVariant);
		logger.debug("Was setting new \"Title\" for \"Shape\" ole object successfull? "+isSuccess);
		
		titleVariant.dispose();

		return isSuccess;
	}
	
	
	/**
	 * Get the name of the given shape		
	 * @param shapeAutomation an OleAutomation that provides access to the "Shape" Ole object. It represents a single shape.
	 * @return a string that represents the name of the shape
	 */
	public static String getShapeName(OleAutomation shapeAutomation){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] namePropertyIds = shapeAutomation.getIDsOfNames(new String[]{"Name"}); 
		if(namePropertyIds==null)
			logger.error("Could not get id of property \"Name\" for \"Shape\" ole object");
		
		Variant nameVariant = shapeAutomation.getProperty(namePropertyIds[0]);
		logger.debug("Invoking get property \"Name\" for \"Shape\" ole object returned variant: "+nameVariant);
		
		String name = nameVariant.getString();
		logger.debug("The value of the property \"Name\" for the \"Shape\" ole object is "+name);
		
		nameVariant.dispose();

		return name;
	}
	
	
	/**
	 * Set a name to the given shape		
	 * @param shapeAutomation an OleAutomation that provides access to the "Shape" Ole object. It represents a single shape.
	 * @param name a string that represents the name to set
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean setShapeName(OleAutomation shapeAutomation, String name){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] namePropertyIds = shapeAutomation.getIDsOfNames(new String[]{"Name"}); 
		if(namePropertyIds==null)
			logger.error("Could not get id of property \"Name\" for \"Shape\" ole object");
	
		Variant nameVariant = new Variant(name);
		logger.debug("The name to set for \"Shape\" ole object is: "+name);
		
		boolean isSuccess = shapeAutomation.setProperty(namePropertyIds[0], nameVariant);
		logger.debug("Setting property \"Name\" for \"Shape\" ole object returned: "+isSuccess);
		
		nameVariant.dispose();

		return isSuccess;
	}
	
	
	/**
	 * Get ID of the given shape		
	 * @param shapeAutomation an OleAutomation that provides access to the "Shape" Ole object. It represents a single shape.
	 * @return a long that represents the ID of the shape
	 */
	public static long getShapeID(OleAutomation shapeAutomation){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] shapeIdPropertyIds = shapeAutomation.getIDsOfNames(new String[]{"ID"}); 
		if(shapeIdPropertyIds==null)
			logger.error("Could not get the id of the property \"ID\" for the \"Shape\" ole object");
		
		Variant idVariant = shapeAutomation.getProperty(shapeIdPropertyIds[0]);
		logger.debug("Invoking get property \"ID\" for the \"Shape\" ole object returned variant "+idVariant);
		long id = idVariant.getLong();
		idVariant.dispose();

		return id;
	}
	
	
	/**
	 * Create a copy of the given shape
	 * @param shapeAutomation an OleAutomation that provides access to a Shape object. It represents an individual shape. 
	 * @return an OleAutomation that represents a copy of the given shape
	 */
	public static OleAutomation copyShape(OleAutomation shapeAutomation){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] copyMethodIds = shapeAutomation.getIDsOfNames(new String[]{"Copy"}); 
		if(copyMethodIds==null)
			logger.error("Could not get the ids of the method \"Copy\" for the \"Shape\" ole object");
		
		Variant result = shapeAutomation.invoke(copyMethodIds[0]);
		logger.debug("Invoking method \"Copy\" for the \"Shape\" ole object returned variant "+result);
		
		if(result==null){
			return null;
		}
		
		OleAutomation shapeCopy = result.getAutomation();
		result.dispose();
		
		return shapeCopy;
	}
	
	/**
	 * Duplicate of the given shape
	 * @param shapeAutomation an OleAutomation that provides access to a Shape object. It represents an individual shape. 
	 * @return an OleAutomation that represents the duplicate of the given shape
	 */
	public static OleAutomation duplicateShape(OleAutomation shapeAutomation){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] duplicateMethodIds = shapeAutomation.getIDsOfNames(new String[]{"Duplicate"}); 
		if(duplicateMethodIds==null)
			logger.error("Could not get the ids of the method \"Duplicate\" for the \"Shape\" ole object");
		
		Variant result = shapeAutomation.invoke(duplicateMethodIds[0]);
		logger.debug("Invoking method \"Duplicate\" for the \"Shape\" ole object returned variant "+result);
		
		if(result==null){
			return null;
		}
		
		OleAutomation shapeCopy = result.getAutomation();
		result.dispose();
		
		return shapeCopy;
	}
	
	/**
	 * Set the distance, in points, from the left edge of column A to the left edge of the shape.
	 * @param shapeAutomation an OleAutomation that provides access to a Shape object. It represents an individual shape. 
	 * @return true if the property was updated, false otherwise
	 */
	public static boolean setShapeLeftPosition(OleAutomation shapeAutomation, double points){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] leftPropertyIds = shapeAutomation.getIDsOfNames(new String[]{"Left"});
		if(leftPropertyIds==null)
			logger.error("Could not get the id of the property \"Left\" for the \"Shape\" ole object");
		
		Variant distance = new Variant(points);
		Boolean isSuccess = shapeAutomation.setProperty(leftPropertyIds[0], distance);
		
		logger.debug("Invoking set property \"Left\" for the \"Shape\" ole object returned: "+isSuccess);
		
		distance.dispose();
		return isSuccess;
	}

	/**
	 * Set the distance, in points, from the top edge of the worksheet to the top edge of the shape.
	 * @param shapeAutomation an OleAutomation that provides access to a Shape object. It represents an individual shape. 
	 * @return true if the property was updated, false otherwise
	 */
	public static boolean setShapeTopPosition(OleAutomation shapeAutomation, double points){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] topPropertyIds = shapeAutomation.getIDsOfNames(new String[]{"Top"});
		if(topPropertyIds==null)
			logger.error("Could not get the id of the property \"Top\" for the \"Shape\" ole object");
		
		Variant distance = new Variant(points);
		Boolean isSuccess = shapeAutomation.setProperty(topPropertyIds[0], distance);
		
		logger.debug("Invoking set property \"Top\" for the \"Shape\" ole object returned: "+isSuccess);
		
		distance.dispose();
		return isSuccess;
	}
	
	/**
	 * Set the height for the given shape
	 * @param shapeAutomation an OleAutomation that provides access to a Shape object. It represents an individual shape. 
	 * @return true if the property was updated, false otherwise
	 */
	public static boolean setShapeHeight(OleAutomation shapeAutomation, double points){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] heightPropertyIds = shapeAutomation.getIDsOfNames(new String[]{"Height"});
		if(heightPropertyIds==null)
			logger.error("Could not get the id of the property \"Height\" for the \"Shape\" ole object");
		
		Variant distance = new Variant(points);
		Boolean isSuccess = shapeAutomation.setProperty(heightPropertyIds[0], distance);
		
		logger.debug("Invoking set property \"Height\" for the \"Shape\" ole object returned: "+isSuccess);
		
		distance.dispose();
		return isSuccess;
	}
	
	
	/**
	 * Set the width for the given shape
	 * @param shapeAutomation an OleAutomation that provides access to a Shape object. It represents an individual shape. 
	 * @return true if the property was updated, false otherwise
	 */
	public static boolean setShapeWidth(OleAutomation shapeAutomation, double points){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] widthPropertyIds = shapeAutomation.getIDsOfNames(new String[]{"Width"});
		if(widthPropertyIds==null)
			logger.error("Could not get the id of the property \"Width\" for the \"Shape\" ole object");
		
		Variant distance = new Variant(points);
		Boolean isSuccess = shapeAutomation.setProperty(widthPropertyIds[0], distance);
		
		logger.debug("Invoking set property \"Width\" for the \"Shape\" ole object returned: "+isSuccess);
		
		distance.dispose();
		return isSuccess;
	}
	
	/**
	 * Set the visibility of the given shape
	 * @param shapeAutomation an OleAutomation that provides access to the "Shape" Ole object. It represents a single shape.
	 * @param visible true if the shape should be shown, false to hide the shape
	 * @return if the operation was successful the method will return true, otherwise it will return false
	 */
	public static boolean setShapeVisibility(OleAutomation shapeAutomation, boolean visible){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] visiblePropertyIds = shapeAutomation.getIDsOfNames(new String[]{"Visible"});
		if(visiblePropertyIds==null)
			logger.error("Could not get the id of the property \"Visible\" for the \"Shape\" ole object");
		
		Variant visibilityVariant = new Variant(visible);
		boolean isSuccess = shapeAutomation.setProperty(visiblePropertyIds[0], visibilityVariant);
		logger.debug("Invoking set property \"ID\" for the \"Shape\" ole object returned: "+isSuccess);
		visibilityVariant.dispose();
		
		return isSuccess;
	}
	
	/**
	 * Get FillFormat OleAutomation. This object can be used to change the format of the shape fill 
	 * @param shapeAutomation an OleAutomation that provides access to a Shape object. It represents an individual shape. 
	 * @return FillFormat OleAutomation for the specified shape. 
	 */
	public static OleAutomation getFillFormatAutomation(OleAutomation shapeAutomation){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] fillPropertyIds = shapeAutomation.getIDsOfNames(new String[]{"Fill"}); 
		if(fillPropertyIds==null)
			logger.error("Could not get the id of the property \"Fill\" for the \"Shape\" ole object");
		
		Variant fillFormatVariant = shapeAutomation.getProperty(fillPropertyIds[0]);
		logger.debug("Invoking get property \"Fill\" for the \"Shape\" ole object returned variant: "+fillFormatVariant);
		
		OleAutomation fillFormatAutomation = fillFormatVariant.getAutomation();
		fillFormatVariant.dispose();
		
		return fillFormatAutomation;
	}
	
	
	/**
	 * Get the TextFrame OleAutomation. This object can be used to manage text in a shape.
	 * @param shapeAutomation an OleAutomation that provides access to a Shape object. It represents an individual shape. 
	 * @return an automation to access the TextFrame functionalities. 
	 */
	public static OleAutomation getTextFrameAutomation(OleAutomation shapeAutomation){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] textFramePropertyIds = shapeAutomation.getIDsOfNames(new String[]{"TextFrame"}); 
		if(textFramePropertyIds==null)
			logger.error("Could not get the id of the property \"TextFrame\" for the \"Shape\" ole object");

		Variant textFrameVariant = shapeAutomation.getProperty(textFramePropertyIds[0]);
		logger.debug("Invoking get property \"TextFrame\" for the \"Shape\" ole object returned variant: "+textFrameVariant);
		
		OleAutomation textFrameAutomation = textFrameVariant.getAutomation();
		textFrameVariant.dispose();
		
		return textFrameAutomation;
	}
	
	
	/**
	 * Get the LineFormat OleAutomation. This object can be used to format the border of the shape.
	 * @param shapeAutomation an OleAutomation that provides access to a Shape object. It represents an individual shape. 
	 * @return an automation to access the LineFormat functionalities. 
	 */
	public static OleAutomation getLineFormatAutomation(OleAutomation shapeAutomation){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] lineFormatPropertyIds = shapeAutomation.getIDsOfNames(new String[]{"Line"}); 
		if(lineFormatPropertyIds==null)
			logger.error("Could not get the id of the property \"Line\" for the \"Shape\" ole object");
		
		Variant lineFormatVariant = shapeAutomation.getProperty(lineFormatPropertyIds[0]);
		logger.debug("Invoking get property \"Line\" for the \"Shape\" ole object returned variant: "+lineFormatVariant);
		
		OleAutomation lineFormatAutomation = lineFormatVariant.getAutomation();
		lineFormatVariant.dispose();
		
		return lineFormatAutomation;
	}
	
	/**
	 * Get the ShadowFormat OleAutomation. This object can be used to format the shadow of the shape.
	 * @param shapeAutomation an OleAutomation that provides access to a Shape object. It represents an individual shape. 
	 * @return an automation to access the ShadowFormat functionalities. 
	 */
	public static OleAutomation getShadowFormatAutomation(OleAutomation shapeAutomation){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] shadowFormatPropertyIds = shapeAutomation.getIDsOfNames(new String[]{"Shadow"}); 
		if(shadowFormatPropertyIds==null)
			logger.error("Could not get the id of the property \"Shadow\" for the \"Shape\" ole object");
		
		Variant shadowFormatVariant = shapeAutomation.getProperty(shadowFormatPropertyIds[0]);
		logger.debug("Invoking get property \"Shadow\" for the \"Shape\" ole object returned variant: "+shadowFormatVariant);
		
		OleAutomation shadowFormatAutomation = shadowFormatVariant.getAutomation();
		shadowFormatVariant.dispose();
		
		return shadowFormatAutomation;
	}
	
	/**
	 * Delete the given shape		
	 * @param shapeAutomation an OleAutomation that provides access to the "Shape" Ole object. It represents a single shape.
	 */
	public static boolean deleteShape(OleAutomation shapeAutomation){
		
		logger.debug("Is Shape OleAutomation null? ".concat(String.valueOf(shapeAutomation==null)));
		
		int[] deleteMethodIds = shapeAutomation.getIDsOfNames(new String[]{"Delete"}); 
		if(deleteMethodIds==null)
			logger.error("Could not get the ids of the method \"Delete\" for the \"Shape\" ole object");
		
		Variant result = shapeAutomation.invoke(deleteMethodIds[0]);
		logger.debug("Invoking the method \"Delete\" for the \"Shape\" ole object returned variant: "+result);
		
		if(result==null){
			return false;
		}
		
		result.dispose();
		return true;
	}
}
