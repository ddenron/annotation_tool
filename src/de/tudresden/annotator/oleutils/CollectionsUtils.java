/**
 * 
 */
package de.tudresden.annotator.oleutils;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class CollectionsUtils {
	
	/**
	 * Get the item having the specified name from a OleAutomation object. The latter is a collection of OLE Objects. 
	 * This method will fail if the OleAutomation does not have the "Item" property.
	 * @param automation an OleAutomation of an OLE collection 
	 * @param itemName a string that represents the name of the item.
	 * @param useMethod if true use invoke method, otherwise use getProperty. For some OLE objects "Item" is a property, for others is a method. 
	 * @return
	 */
	public static OleAutomation getItemByName(OleAutomation automation, String itemName, boolean useMethod){
		
		int[] itemPropertyIds = automation.getIDsOfNames(new String[]{"Item"});
		if(itemPropertyIds == null){
			System.out.println("Property \"Item\" not found for the give Ole object");
			return null;
		}
		
		Variant args[] = new Variant[1];
		args[0] =  new Variant(itemName);
		
		Variant itemVariant; 
		if(!useMethod){
			itemVariant = automation.getProperty(itemPropertyIds[0],args);
		}else{
			itemVariant = automation.invoke(itemPropertyIds[0],args);
		}
		
		if(itemVariant==null){
			return null;
		}
		
		OleAutomation itemAutomation = itemVariant.getAutomation();
		
		args[0].dispose();
		itemVariant.dispose();
		
		return itemAutomation;
	}
	
	
	/**
	 * Get the item having the specified index from a OleAutomation object. The latter is a collection of OLE Objects. 
	 * This method will fail if the OleAutomation does not have the "Item" property.
	 * @param automation an OleAutomation of an OLE collection 
	 * @param index an integer that represents the index number of the item in the collection. 
	 * @param useMethod if true use invoke method, otherwise use getProperty. For some OLE objects "Item" is a property, for others is a method. 
	 * @return
	 */
	public static OleAutomation getItemByIndex(OleAutomation automation, int index, boolean useMethod){
		
		int[] itemPropertyIds = automation.getIDsOfNames(new String[]{"Item"});
		if(itemPropertyIds == null){
			if(!useMethod){
				System.out.println("Property \"Item\" not found for the give Ole object");
			}else{
				System.out.println("Method \"Item\" not found for the give Ole object");
			}
			return null;
		}
		
		Variant args[] = new Variant[1];
		args[0] =  new Variant(index);
		
		Variant itemVariant;
		if(!useMethod){
			itemVariant = automation.getProperty(itemPropertyIds[0],args);
		}else{
			itemVariant = automation.invoke(itemPropertyIds[0],args);
		}
		
		if(itemVariant==null){
			return null;
		}
		
		OleAutomation itemAutomation = itemVariant.getAutomation();
		
		args[0].dispose();
		itemVariant.dispose();
		
		return itemAutomation;
	}
	
	
	/**
	 * Get the number of items in OleAutomation that is (represents) a collection of OLE objects.
	 * This methods will fail if the given OleAutomation does not have the "Count" property.  
	 * @param automation
	 * @return
	 */
	public static int countItemsInCollection(OleAutomation automation){
		
		int[] countProperyIds = automation.getIDsOfNames(new String[]{"Count"});
		if(countProperyIds == null){
			System.out.println("Property \"Count\" not found for the given OleAutomation object!");
			return -1;
		}
				
		Variant countPropertyVariant =  automation.getProperty(countProperyIds[0]);
		if(countPropertyVariant == null){
			System.out.println("\"Count\" variant is null!");
			return -1;
		}				
				
		int count = countPropertyVariant.getInt();
		countPropertyVariant.dispose();
		
		return count;
	}

}
