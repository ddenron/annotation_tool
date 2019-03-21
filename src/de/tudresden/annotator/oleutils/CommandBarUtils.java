/**
 * 
 */
package de.tudresden.annotator.oleutils;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis
 *
 */
public class CommandBarUtils {
	
	/**
	 * Get the specific command bar automation by its name
	 * 
 	 * @param application
	 * @return
	 */
	public static OleAutomation getCommandBarByName(OleAutomation application, String commandBarName) {
		
		int[] commandBarsPropertyIds = application.getIDsOfNames(new String[]{"CommandBars"});
		if (commandBarsPropertyIds == null) {
			System.out.println("Property \"CommandBars\" of \"Application\" OLE Object is null!");
			return null;
		}
		
		Variant commandBarsVariant =  application.getProperty(commandBarsPropertyIds[0]);	
		if(commandBarsVariant == null){
			System.out.println("\"CommandBars\" variant is null!");
			return null;		
		}
		OleAutomation commandBarsAutomation = commandBarsVariant.getAutomation();
		commandBarsVariant.dispose();
			
		int[] itemPropertyIds = commandBarsAutomation.getIDsOfNames(new String[]{"Item"});
		if(itemPropertyIds == null){
			System.out.println("Property \"Item\" of \"CommandBars\" OLE object not found!");
			return null;
		}

		Variant[] parameters = new Variant[1];
		parameters[0] = new Variant(commandBarName);
		Variant cbVariant = commandBarsAutomation.getProperty(itemPropertyIds[0],parameters);
		parameters[0].dispose();
		
		if(cbVariant==null){
			System.out.println("There is no CommandBar named \""+commandBarName+"\"");
			return null;
		}
		OleAutomation commandBarAutomation = cbVariant.getAutomation();
		cbVariant.dispose();

		return commandBarAutomation;
	}
	
	
	/**
	 * Print list of command bars in excel application 
	 * 
	 * @param excelApplication
	 * @return
	 */
	public static boolean printListOfCommandBars(OleAutomation excelApplication) {
		
		int[] commandBarsObjectIds = excelApplication.getIDsOfNames(new String[]{"CommandBars"});
		if (commandBarsObjectIds == null) {
			System.out.println("Property \"CommandBars\" of \"Application\" OLE Object is null!");
			return false;
		}
		
		Variant commandBarsVariant =  excelApplication.getProperty(commandBarsObjectIds[0]);	
		if(commandBarsVariant == null){
			System.out.println("\"CommandBars\" variant is null!");
			return false;		
		}
		OleAutomation commandBarsAutomation = commandBarsVariant.getAutomation();
		commandBarsVariant.dispose();
		
		int[] countProperyIds = commandBarsAutomation.getIDsOfNames(new String[]{"Count"});
		if(countProperyIds == null){
			System.out.println("Property \"Count\" of \"CommandBars\" OLE object is null!");
			return false;
		}
				
		Variant countPropertyVariant =  commandBarsAutomation.getProperty(countProperyIds[0]);
		if(countPropertyVariant == null){
			System.out.println("\"Count\" variant is null!");
			return false;
		}				
				
		int count = countPropertyVariant.getInt();
		countPropertyVariant.dispose();
		
		int[] itemPropertyIds = commandBarsAutomation.getIDsOfNames(new String[]{"Item"});
		if(itemPropertyIds == null){
			System.out.println("Property \"Item\" of \"CommandBars\" is not found!");
			return false;
		}

		System.out.println("\nlist of command bars:\n".toUpperCase());
		for (int i = 1; i <= count; i++) {
			Variant[] args = new Variant[1];
			args[0] = new Variant(i);		
			Variant nextCommandBarVariant = commandBarsAutomation.getProperty(itemPropertyIds[0],args);	
			
			OleAutomation nextCommandBarAutomation = nextCommandBarVariant.getAutomation();
			System.out.println(getCommandBarName(nextCommandBarAutomation));
			
			nextCommandBarVariant.dispose();
			nextCommandBarAutomation.dispose();
			args[0].dispose();
		}
		return true;	
	}
	
	
	/**
	 * Get command bar name 
	 * 
	 * @param commandBarAutomation
	 * @return
	 */
	public static String getCommandBarName(OleAutomation commandBarAutomation){
		
		int[] namePropertyIds = commandBarAutomation.getIDsOfNames(new String[]{"Name"});
		if(namePropertyIds == null){
			System.out.println("Property \"Name\" of \"CommandBar\" is not found!");
			return null;
		}
		Variant nameVariant = commandBarAutomation.getProperty(namePropertyIds[0]);
		if(nameVariant == null){
			System.out.println("\"Name\" variant is null!");
			return null;
		}
		String name = nameVariant.getString();
		nameVariant.dispose();
		return name;
	}
	
	
	/**
	 * Get controls for the given command bar  
	 * 
	 * @param commandBarAutomation
	 * @return
	 */
	public static OleAutomation getCommandBarControls(OleAutomation commandBarAutomation){
		
		int[] controlsPropertyIds = commandBarAutomation.getIDsOfNames(new String[]{"Controls"});
		Variant controlsVariant = commandBarAutomation.getProperty(controlsPropertyIds[0]);
		OleAutomation contolsAutomation = controlsVariant.getAutomation();
		controlsVariant.dispose();
		
		return contolsAutomation;
	}
	
	
	/**
	 * Change the visibility of the given CommandBarContols
	 * 
	 * @param commandBarControls
	 * @param visible
	 * @return
	 */
	public static boolean setVisibilityOfControls(OleAutomation commandBarControls, boolean visible){

		int[] itemPropertyIds = commandBarControls.getIDsOfNames(new String[]{"Item"});
	
		Variant[] parameters = new Variant[1];
		parameters[0] = new Variant(1);
		Variant controlItemVariant = commandBarControls.getProperty(itemPropertyIds[0],parameters);
		parameters[0].dispose();
		
		int i=1;
		while (controlItemVariant!=null) {			
			OleAutomation controlItemAutomation = controlItemVariant.getAutomation();
			int[] visiblePropertyIds = controlItemAutomation.getIDsOfNames(new String[]{"Visible"});
			controlItemAutomation.setProperty(visiblePropertyIds[0],new Variant(visible));
			parameters[0] = new Variant(i++);
			controlItemVariant.dispose();
			controlItemVariant = commandBarControls.getProperty(itemPropertyIds[0],parameters);
			parameters[0].dispose();
		}

		return true;
	}
	
	/**
	 * Perform a temporary delete for the given controls 
	 * @param commandBarContols
	 * @return
	 */
	public static boolean deleteControlsTemporary(OleAutomation commandBarContols){

		int[] itemPropertyIds = commandBarContols.getIDsOfNames(new String[]{"Item"});
	
		Variant[] itemParams = new Variant[1];
		itemParams[0] = new Variant(1);
		Variant controlItemVariant = commandBarContols.getProperty(itemPropertyIds[0],itemParams);
		itemParams[0].dispose();
		
		int i=1;
		while (controlItemVariant!=null) {			
			OleAutomation controlItemAutomation = controlItemVariant.getAutomation();
			
			int[] deleteMethodIds = controlItemAutomation.getIDsOfNames(new String[]{"Delete"});
			Variant[]  args = new Variant[1];
			args[0] = new Variant(true);
			controlItemAutomation.invoke(deleteMethodIds[0],args);
			args[0].dispose();
			controlItemAutomation.dispose();
			
			itemParams[0] = new Variant(i++);
			controlItemVariant.dispose();
			controlItemVariant = commandBarContols.getProperty(itemPropertyIds[0],itemParams);
			itemParams[0].dispose();
		}

		return true;
	}
	
	/**
	 * Delete the custom controls that were created during the current (this) session of the application.
	 * 
	 * @param commandBarContols
	 * @param tag
	 * @return
	 */
	public static boolean deleteCustomControlsByTag(OleAutomation commandBarContols, String tag){

		int[] itemPropertyIds = commandBarContols.getIDsOfNames(new String[]{"Item"});
	
		Variant[] parameters = new Variant[1];
		parameters[0] = new Variant(1);
		Variant controlItemVariant = commandBarContols.getProperty(itemPropertyIds[0],parameters);
		parameters[0].dispose();
		
		int i=1;
		while (controlItemVariant!=null) {			
			OleAutomation controlItemAutomation = controlItemVariant.getAutomation();
			int[] tagPropertyIds = controlItemAutomation.getIDsOfNames(new String[]{"Tag"});
			Variant tagVariant = controlItemAutomation.getProperty(tagPropertyIds[0]);
			
			if(tagVariant.getString().compareToIgnoreCase(tag)==0){
				int[] deleteMethodIds = controlItemAutomation.getIDsOfNames(new String[]{"Delete"});
				controlItemAutomation.invoke(deleteMethodIds[0]);
			}
			tagVariant.dispose();
			
			parameters[0] = new Variant(i++);
			controlItemVariant.dispose();
			controlItemVariant = commandBarContols.getProperty(itemPropertyIds[0],parameters);
			parameters[0].dispose();
		}

		return true;
	}
	
	
	/**
	 * Set the visibility for the given command bar
 	 * @param application
	 * @return
	 */
	public static boolean setVisibilityForCommandBar(OleAutomation application, String commandBarName, boolean visible) {
		
		OleAutomation tabsCBAutomation = getCommandBarByName(application, commandBarName); 
		
		if(tabsCBAutomation==null)
			return false;
			
		int[] visiblePropertyIds = tabsCBAutomation.getIDsOfNames(new String[]{"Visible"});	
		boolean isSuccess = tabsCBAutomation.setProperty(visiblePropertyIds[0], new Variant(visible));
		return isSuccess;
	}
	
	/**
	 * Set the value of the enabled property for given command bar
 	 * @param application
	 * @return
	 */
	public static boolean setEnabledForCommandBar(OleAutomation application, String commandBarName, boolean enabled) {
		
		OleAutomation tabsCBAutomation = getCommandBarByName(application, commandBarName);
		
		if(tabsCBAutomation==null)
			return false;
		
		int[] enabledPropertyIds = tabsCBAutomation.getIDsOfNames(new String[]{"Enabled"});		
		boolean isSuccess = tabsCBAutomation.setProperty(enabledPropertyIds[0], new Variant(enabled));
		
		tabsCBAutomation.dispose();
		return isSuccess;
	}
	
	
	/**
	 * Hide the formula bar from Excel UI
	 * @param excelApplication
	 * @return
	 */
	public static boolean hideFormulaBar(OleAutomation excelApplication){
			
		int[] displayFormulaBarIds = excelApplication.getIDsOfNames(new String[]{"DisplayFormulaBar"});
		Boolean isUpdated = excelApplication.setProperty(displayFormulaBarIds[0],new Variant(false));
		
		return isUpdated;
	}
	
	
	/**
	 * Set whether to display or not floating menus on right click  
	 * 
	 * @param excelApplication
	 * @return
	 */
	public static boolean setShowMenuFloaties(OleAutomation excelApplication, boolean option){
		
		int[] showMenuFloatiesIds = excelApplication.getIDsOfNames(new String[]{"ShowMenuFloaties"});	
		if (showMenuFloatiesIds == null){			
			System.out.println("\"ShowMenuFloaties\" property not found for \"Application\" OLE object!");
			return false;
		}	
		
		boolean result = excelApplication.setProperty(showMenuFloatiesIds[0],new Variant(option));
		return result;
	}
	
}
