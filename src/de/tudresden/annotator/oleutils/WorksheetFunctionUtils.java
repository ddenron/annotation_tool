/**
 * 
 */
package de.tudresden.annotator.oleutils;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class WorksheetFunctionUtils {
	
	/**
	 * Call an excel worksheet function (Ex. count, sum, min, max)
	 * @param worksheetFunction an OleAutomation that provides access to the worksheet functions
	 * @param name a string that represents the name of the function to call 
	 * @param args an array of variants that represent the arguments of the function to call 
	 * @return the result of the function
	 */
	public static Variant callFunction(OleAutomation worksheetFunction, String name, Variant[] args){
		
		int[] methodIds = worksheetFunction.getIDsOfNames(new String[]{name});
		if(methodIds==null){
			System.out.println("Method \""+name+"\" is not found for \"WorksheetFunction\" OLE Object!");
			return null;
		}
		
		Variant result = worksheetFunction.invoke(methodIds[0], args);
		 
		// invoke the unprotect method  
		return result;
	}
	
	
	public static int countBlankCells(OleAutomation application, OleAutomation rangeAuto){
		
		OleAutomation worksheetFunction = ApplicationUtils.getWorksheetFunctionAutomation(application);
		
		Variant[] args = new Variant[1];
		args[0] = new Variant(rangeAuto);
		
		Variant countBlankVariant = WorksheetFunctionUtils.callFunction(worksheetFunction, "COUNTBLANK", args);
		
		int countBlank = 0;
		if(countBlankVariant!=null){
			countBlank = countBlankVariant.getInt();
			countBlankVariant.dispose();
		}
		
		worksheetFunction.dispose();
		
		return countBlank;
	}

	
	public static int countNumericCells(OleAutomation application, OleAutomation rangeAuto){
		
		OleAutomation worksheetFunction = ApplicationUtils.getWorksheetFunctionAutomation(application);
		
		Variant[] args = new Variant[1];
		args[0] = new Variant(rangeAuto);
		
		Variant countNumericVariant = WorksheetFunctionUtils.callFunction(worksheetFunction, "COUNT", args);
		
		int countNumeric = 0;
		if(countNumericVariant!=null){
			countNumeric = countNumericVariant.getInt();
			countNumericVariant.dispose();
		}
		
		worksheetFunction.dispose();
		
		return countNumeric;
	}
	
	
	public static int countTextCells(OleAutomation application, OleAutomation rangeAuto){
		
		OleAutomation worksheetFunction = ApplicationUtils.getWorksheetFunctionAutomation(application);
		
		Variant[] args = new Variant[2];
		args[0] = new Variant(rangeAuto);
		args[1] = new Variant("*");
		
		Variant countTextVariant = WorksheetFunctionUtils.callFunction(worksheetFunction, "COUNTIF", args);
		args[1].dispose();
		
		int countText = 0;
		if(countTextVariant!=null){
			countText = countTextVariant.getInt();
			countTextVariant.dispose();
		}
		
		worksheetFunction.dispose();
		
		return countText;
	}
	
	
	public static int countNotEmptyCells(OleAutomation application, OleAutomation rangeAuto){
		
		OleAutomation worksheetFunction = ApplicationUtils.getWorksheetFunctionAutomation(application);
		
		Variant[] args = new Variant[1];
		args[0] = new Variant(rangeAuto);
		
		Variant countNotEmptyVariant = WorksheetFunctionUtils.callFunction(worksheetFunction, "COUNTA", args);
		
		int countNotEmpty = 0;
		if(countNotEmptyVariant!=null){
			countNotEmpty = countNotEmptyVariant.getInt();
			countNotEmptyVariant.dispose();
		}
		
		worksheetFunction.dispose();
		
		return countNotEmpty;
	}

	public static int countLogicalCells(OleAutomation application, OleAutomation rangeAuto){
		
		OleAutomation wf1 = ApplicationUtils.getWorksheetFunctionAutomation(application);	
		Variant[] args1 = new Variant[2];
		args1[0] = new Variant(rangeAuto);
		args1[1] = new Variant("FALSE");
		Variant countFalseVariant = WorksheetFunctionUtils.callFunction(wf1, "COUNTIF", args1);
		args1[1].dispose();
		wf1.dispose();
		
		int countFalse = 0;
		if(countFalseVariant!=null){
			countFalse = countFalseVariant.getInt();
			countFalseVariant.dispose();
		}
		
		OleAutomation wf2 = ApplicationUtils.getWorksheetFunctionAutomation(application);	
		Variant[] args2 = new Variant[2];
		args2[0] = new Variant(rangeAuto);
		args2[1] = new Variant("TRUE");
		Variant countTrueVariant = WorksheetFunctionUtils.callFunction(wf2, "COUNTIF", args2);
		args2[1].dispose();
		wf2.dispose();
		
		int countTrue = 0;
		if(countTrueVariant!=null){
			countTrue = countTrueVariant.getInt();
			countTrueVariant.dispose();
		}
		
		return countTrue + countFalse;
	}

//	
//	public static int countFormulaCells(OleAutomation application, OleAutomation rangeAuto){
//		
//		OleAutomation wf1 = ApplicationUtils.getWorksheetFunctionAutomation(application);		
//		OleAutomation wf2 = ApplicationUtils.getWorksheetFunctionAutomation(application);		
//		OleAutomation wf3 = ApplicationUtils.getWorksheetFunctionAutomation(application);
//			
//		Variant result = WorksheetFunctionUtils.callFunction(wf3, "SumProduct", new Variant[]{ WorksheetFunctionUtils.callFunction(wf2, "Ceiling_Math", 
//				new Variant[] {WorksheetFunctionUtils.callFunction(wf1, "IsFormula", new Variant[] {new Variant(rangeAuto)})})});
//		
//		wf1.dispose();
//		wf2.dispose();
//		wf3.dispose();
//		
//		int countFormula = 0; 
//		if(result!=null){
//			countFormula =result.getInt();
//			result.dispose();
//		}
//				
//		return countFormula;
//	}
	
}
