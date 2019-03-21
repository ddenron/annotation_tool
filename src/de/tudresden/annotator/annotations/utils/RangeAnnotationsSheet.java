/**
 * 
 */
package de.tudresden.annotator.annotations.utils;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;

import org.eclipse.swt.SWT;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.widgets.MessageBox;

import de.tudresden.annotator.annotations.RangeAnnotation;
import de.tudresden.annotator.annotations.WorkbookAnnotation;
import de.tudresden.annotator.annotations.WorksheetAnnotation;
import de.tudresden.annotator.annotations.format.AnnotationClass;
import de.tudresden.annotator.main.Launcher;
import de.tudresden.annotator.oleutils.ApplicationUtils;
import de.tudresden.annotator.oleutils.CollectionsUtils;
import de.tudresden.annotator.oleutils.RangeUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;
import de.tudresden.annotator.oleutils.WorksheetUtils;

/**
 * @author Elvis Koci
 */
public class RangeAnnotationsSheet {
	
	protected static final String name = "Range_Annotations_Data";
	private static String startColumn = "A";
	private static int startRow = 1; 
	
	/**
	 * This linked hash map stores the names of the fields used in the Annotation Data Sheet and their default order. 
	 */
	private static final LinkedHashMap<String, Integer> fields;
	static
    {
		fields = new LinkedHashMap<String,Integer>();
		fields.put("Sheet.Name", 0); // required 
		fields.put("Sheet.Index", 1); // required
		fields.put("Annotation.Label", 2); // required
		fields.put("Annotation.Name", 3); // required
		fields.put("Annotation.Range", 4); // required
		fields.put("Annotation.Parent", 5); // required
		fields.put("TotalCells", 6); // optional
		fields.put("EmptyCells", 7); // optional
		fields.put("ConstantCells", 8); // optional
		fields.put("FormulaCells", 9); // optional
		fields.put("HasMergedCells", 10); // optional
		fields.put("Rows", 11); // optional
		fields.put("Columns", 12); // optional
    }
	
	
	/**
	 * Save new annotation data
	 * @param workbookAutomation an OleAutomation to access the embedded workbook
	 * @param annotation a RangeAnnotation object that maintains (contains) the annotation data to be saved
	 */
	public static void saveRangeAnnotationData(OleAutomation workbookAutomation, RangeAnnotation annotation){
			
		OleAutomation annotationDataSheet =  WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		
		if(annotationDataSheet==null){		
			annotationDataSheet = createRangeAnnotationsSheet(workbookAutomation);	
		}
		
		OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationDataSheet);		
		String usedAddress = RangeUtils.getRangeAddress(usedRange);
		usedRange.dispose();
				
		String[] cells = usedAddress.split(":");		
		int endRow = Integer.valueOf(cells[1].replaceAll("[^0-9]+",""));
		int row = endRow + 1;
					
		writeNewDataRow(annotationDataSheet, row, annotation);		
		annotationDataSheet.dispose();
	}
	
	
	/**
	 * Save the data of many range annotations at once
	 * @param workbookAutomation an OleAutomation to access the embedded workbook
	 */
	public static void saveManyRangeAnnotations(OleAutomation workbookAutomation){
			
		OleAutomation rangeAnnotationsDataSheet =  WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		
		if(rangeAnnotationsDataSheet==null){		
			rangeAnnotationsDataSheet = createRangeAnnotationsSheet(workbookAutomation);
		}
			
		WorksheetUtils.unprotectWorksheet(rangeAnnotationsDataSheet);
		
		// delete all the existing data from the sheet. 
		// by removing all existing data we ensure that the "new" data will have
		// the right format. So, they are not effected by the existing data.
		OleAutomation usedRange = WorksheetUtils.getUsedRange(rangeAnnotationsDataSheet);		
		RangeUtils.deleteRange(usedRange);	
		usedRange.dispose();
		
		// re-create the header	
		OleAutomation rangeAuto = WorksheetUtils.getRangeAutomation(rangeAnnotationsDataSheet, startColumn+""+startRow, null);
		int colNum = RangeUtils.getFirstColumnIndex(rangeAuto);
		rangeAuto.dispose();
		
		Iterator<String> itr = fields.keySet().iterator();
		int i = 0;	
		while (itr.hasNext()) {
			OleAutomation cell = WorksheetUtils.getCell(rangeAnnotationsDataSheet, startRow, colNum+i);
			RangeUtils.setValue(cell, itr.next());
			i++;
		}
		
		// write the data for each range annotation
		WorksheetUtils.protectWorksheet(rangeAnnotationsDataSheet);	
		int j=startRow+1;
		for(RangeAnnotation ra: AnnotationHandler.getWorkbookAnnotation().getAllAnnotations()){
			AnnotationHandler.calculateStatistics(ra,workbookAutomation);
			writeNewDataRow(rangeAnnotationsDataSheet, j++, ra);				
		}
		rangeAnnotationsDataSheet.dispose();
	}
	
	/**
	 * Create the sheet that will store the annotation data 
	 * @param workbookAutomation an OleAutomation to access the embedded workbook
	 * @return the OleAutomation of the created worksheet
	 */
	private static OleAutomation createRangeAnnotationsSheet(OleAutomation workbookAutomation){
		
		WorkbookUtils.unprotectWorkbook(workbookAutomation);
		
		OleAutomation newWorksheet = WorkbookUtils.addWorksheetAsLast(workbookAutomation);
		WorksheetUtils.setWorksheetName(newWorksheet, name);
				
		OleAutomation rangeAuto = WorksheetUtils.getRangeAutomation(newWorksheet, startColumn+""+startRow, null);
		int colNum = RangeUtils.getFirstColumnIndex(rangeAuto);
		rangeAuto.dispose();
		
		Iterator<String> itr = fields.keySet().iterator();
		int i = 0;	
		while (itr.hasNext()) {
			OleAutomation cell = WorksheetUtils.getCell(newWorksheet, startRow, colNum+i);
			RangeUtils.setValue(cell, itr.next());
			i++;
		}
				
		//WorksheetUtils.setWorksheetVisibility(newWorksheet, false);		
		WorkbookUtils.protectWorkbook(workbookAutomation, true, false);
		
		return newWorksheet;
	}
	
	
	/**
	 * Write new row of annotation data
	 * @param annotationDataSheet an OleAutomation that provides access to the sheet that maintains the annotation data
	 * @param row an integer that represents the index of the row to write the data
	 * @param annotation a RangeAutomation object that maintains (contains) the annotation data to write  
	 */
	private static void writeNewDataRow(OleAutomation annotationDataSheet, int row, RangeAnnotation annotation){		
		
		WorksheetUtils.unprotectWorksheet(annotationDataSheet);
		
		OleAutomation topLeftCell = WorksheetUtils.getRangeAutomation(annotationDataSheet, startColumn+""+row, null);
		int colNum = RangeUtils.getFirstColumnIndex(topLeftCell);
		topLeftCell.dispose();
		
		Iterator<String> itr = fields.keySet().iterator();
		int i = 0;	
		while (itr.hasNext()) {
			OleAutomation cell = WorksheetUtils.getCell(annotationDataSheet, row, colNum+i);
			RangeUtils.formatCells(cell, "@");
			RangeUtils.setValue(cell, getFieldValue(itr.next(), annotation));	
			i++;
		}
		
		WorksheetUtils.protectWorksheet(annotationDataSheet);	
	}
	
	
	/**
	 * Get the value for the field from the corresponding attribute/s of the RangeAnnotation object 
	 * @param fieldName a string that represents the name of a field from the header row in the annotation data sheet 
	 * @param annotation a RangeAutomation object that maintains (contains) the annotation data to be retrieved  
	 * @return a string that represents the value of the specified (given) field
	 */
	private static String getFieldValue(String fieldName, RangeAnnotation annotation){
		
		String value = null;		
		switch (fieldName) {
			case "Sheet.Name"  : value = annotation.getSheetName(); break;
			case "Sheet.Index"  : value = String.valueOf(annotation.getSheetIndex()); break;
			case "Annotation.Label"  : value = annotation.getAnnotationClass().getLabel(); break;
			case "Annotation.Name"  : value = annotation.getName(); break;
			case "AnnotationTool.Name"  : value = annotation.getAnnotationClass().getAnnotationTool().name(); break;
			case "AnnotationTool.Code"  : value = String.valueOf(annotation.getAnnotationClass().getAnnotationTool().getCode()); break;
			case "Annotation.Range"  : value = annotation.getRangeAddress(); break;
			case "Annotation.Parent" : value = annotation.getParent() instanceof  RangeAnnotation ? 
											   ((RangeAnnotation) annotation.getParent()).getName() : annotation.getSheetName(); break;
			case "TotalCells" : value =  String.valueOf(annotation.getCells()); break;
			case "EmptyCells" : value =  String.valueOf(annotation.getEmptyCells()); break;
			case "ConstantCells" : value =  String.valueOf(annotation.getConstantCells()); break;
			case "FormulaCells" : value =  String.valueOf(annotation.getFormulaCells()); break;
			case "HasMergedCells" : value =  String.valueOf(annotation.containsMergedCells()); break;
			case "Rows" : value =  String.valueOf(annotation.getRows()); break;
			case "Columns": value =  String.valueOf(annotation.getColumns()); break;
			default: value = "Field not recognized!"; break;
		}
		return value;
	}
	
		
	/**
	 * Read all annotation data from the "Annotation Data" Sheet 
	 * @param workbookAutomation an OleAutomation that provides access to the embedded workbook
	 * @return null if the range annotations data could not be read, otherwise an array of 
	 * all recovered range annotations and their dependences 
	 */
	public static RangeAnnotation[] readRangeAnnotations(OleAutomation workbookAutomation){
		
		// get the OleAutomation object for the sheet that stores 
		// the annotation metadata (a.k.a. annotation data sheet) 
		OleAutomation annotationDataSheet = 
				WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		
		// workbook has no annotation data
		if(annotationDataSheet==null)
			return null;
		
		// TODO: Use special cells instead
		// get the range that has data. check that it is not empty
		OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationDataSheet);
		if(usedRange==null){
			int style = SWT.ICON_ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("Annotation data sheet is empty. Header row is missing! "
							+"Please delete the \""+name+"\" worksheet before proceeding with the annotation");
			message.open();
			return null;
		}
		
		// get used areas. check that there is only one such area
		OleAutomation usedAreas = RangeUtils.getAreas(usedRange);
		int countAreas = CollectionsUtils.countItemsInCollection(usedAreas);
		usedAreas.dispose();
		
		if(countAreas!=1){
			int style = SWT.ICON_ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("Could not read the annotation data from the sheet "+name+".\n" +
							"Data are not in the expected format. " +
							"There seem to be empty rows or empty columns between the data!");
			message.open();
			return null;
		}
		
		// get the bounding rows and columns 
		String usedRangeAddress = RangeUtils.getRangeAddress(usedRange);
		usedRange.dispose();
		
		String boundingCells[] = usedRangeAddress.split(":"); 
		String topLeftCell = boundingCells[0];
		String topLeftColumn = topLeftCell.replaceAll("[0-9\\$]+","");
		int topLeftRow = Integer.valueOf(topLeftCell.replaceAll("[^0-9]+",""));
		String downRightCell = boundingCells[1];
		String downRightColumn = downRightCell.replaceAll("[0-9\\$]+","");
		int downRightRow = Integer.valueOf(downRightCell.replaceAll("[^0-9]+",""));
		
		// ensure that the header row contains all the expected fields
		// if all required fields are present, save their order
		if(!validateHeaderRow(annotationDataSheet, topLeftRow, topLeftColumn, downRightColumn))
			return null;
		
		// read all the data rows and re-create the range annotations and their dependencies
		LinkedHashMap<String, RangeAnnotation> rangeAnnotations = new LinkedHashMap<String, RangeAnnotation>();	
		WorkbookAnnotation wa = AnnotationHandler.getWorkbookAnnotation();
		for (int i = (topLeftRow + 1); i <=downRightRow; i++) {
			
			String[] rangeAnnotationData = readDataRow(annotationDataSheet, i, topLeftColumn, downRightColumn);
			
			//re-create the range annotation object
			AnnotationClass annotationClass = ClassGenerator.getAnnotationClasses().get(rangeAnnotationData[2]);	
			RangeAnnotation annotation = new RangeAnnotation(rangeAnnotationData[0], Integer.valueOf(rangeAnnotationData[1]), 
									annotationClass, rangeAnnotationData[3], rangeAnnotationData[4]); 
			
			// set the parent annotation
			boolean hasParent = false;
			RangeAnnotation parentAnnotation = rangeAnnotations.get(rangeAnnotationData[5]);
			if(parentAnnotation != null){
				annotation.setParent(parentAnnotation);
				hasParent = true;
			}else{
				WorksheetAnnotation sa = wa.getWorksheetAnnotations().get(annotation.getSheetName());
				if(sa!=null){
					annotation.setParent(sa);
					hasParent = true;
				}
			}
			
			// if parent not found, discard the range annotation
			if(hasParent){
				rangeAnnotations.put(annotation.getName(), annotation);
			}
		}
			
		return rangeAnnotations.values().toArray(new RangeAnnotation[rangeAnnotations.size()]);
	}
	
	
	/**
	 * Validate header row. It should contain all the expected (predefined) fields  
	 * @param annotationDataSheet an OleAutomation to access the functionalities of the sheet that stores the annotation data
	 * @param topLeftRow an integer that represents the address of the top left row
	 * @param topLeftColumn a string that represents the address of the column on the top left
	 * @param downRightColumn a string that represents the address of the column on the down right
	 * @return true if the header row passes all checks, false if validation fails. 
	 */
	protected static boolean validateHeaderRow(OleAutomation annotationDataSheet, int topLeftRow, String topLeftColumn, String downRightColumn){
		
		// get all the values from the range that represents the header
		String topLeftCell = topLeftColumn+topLeftRow;
		String downRightCell = downRightColumn+topLeftRow;		
		OleAutomation rangeAutomation = WorksheetUtils.getRangeAutomation(annotationDataSheet, topLeftCell, downRightCell);
		String values[] = RangeUtils.getRangeValues(rangeAutomation);
		rangeAutomation.dispose();
		
		// check if the number of fields in the sheet match with the pre-defined (expected) ones.  
		if(values.length > fields.size()){
			int style = SWT.ICON_ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("The number of fields in the annotation data sheet is larger than the declared (expected) fields");
			message.open();
			return false;
		}

		// check that the header row contains recognizable fields. update their order
		for (int i = 0; i< values.length; i++) {
			
			String val = values[i];
			if( val==null || val.compareTo("")==0){
				int style = SWT.ICON_ERROR;
				MessageBox message = Launcher.getInstance().createMessageBox(style);
				message.setMessage("There are empty cells in the header row! Each field has to have a name.");
				message.open();
				return false;
			}
			
			if(!fields.containsKey(val)){
				int style = SWT.ICON_ERROR;
				MessageBox message = Launcher.getInstance().createMessageBox(style);
				message.setMessage("Field \""+val+"\" is not recognized. It is not part of the pre-defined (expected) fields");
				message.open();
				return false;
			}
			
			fields.put(val, i);
		} 
		
		// update the start column and row (e.i., the address of the first cell) of the range that contains the annotation data
		startColumn = topLeftColumn;
		startRow = topLeftRow;
				
		return true;
	}
	
	
	/**
	 * Read a row of range annotation data.
	 * @param annotationDataSheet an OleAutomation to access the functionalities of the worksheet that stores the annotation data
	 * @param row an integer that represents the row to be read
	 * @param startColumn a string that represents the column to start the reading
	 * @param endColumn a string that represents the column to end the reading
	 * @return an array of string values that represent the range annotation data in the standard order.
	 */
	protected static String[] readDataRow(OleAutomation annotationDataSheet, int row, String startColumn, String endColumn){		
		
		// get the values of the cells in the given row, bounded by the start and the end column.   
		String topLeftCell = startColumn+row;
		String downRightCell = endColumn+row;		
		OleAutomation rangeAutomation = WorksheetUtils.getRangeAutomation(annotationDataSheet, topLeftCell, downRightCell);
		String values[] = RangeUtils.getRangeValues(rangeAutomation);
		rangeAutomation.dispose();
	
		// create an array of values in the standard order
		String [] rangeAnnotationData = new String[fields.size()];
		
		rangeAnnotationData[0] = values[fields.get("Sheet.Name")];
		rangeAnnotationData[1] = values[fields.get("Sheet.Index")];
		rangeAnnotationData[2] = values[fields.get("Annotation.Label")];
		rangeAnnotationData[3] = values[fields.get("Annotation.Name")];
		rangeAnnotationData[4] = values[fields.get("Annotation.Range")];
		rangeAnnotationData[5] = values[fields.get("Annotation.Parent")];
		
		return rangeAnnotationData;
	
	}
	
	/**
	 * Delete the annotation data for the sheet with the given name 
	 * This method will clear all the rows in the annotation (meta-)data sheet 
	 * that are associated with the specified sheet  
	 * @param workbookAutomation an OleAumation for accessing the functionalities of the embedded workbook
	 * @param sheetName the name of the sheet where the annotation are placed (drawn)
	 * @param permanentDelete if true the annotation data will be deleted permanently, otherwise they will just be hidden  
	 */
	public static void deleteRangeAnnotationDataFromSheet(OleAutomation workbookAutomation, String sheetName){
		deleteDataRows(workbookAutomation, "Sheet.Name", sheetName);
	}
	
	/**
	 * Delete the data row from the sheet for the specified range annotation 
	 * @param workbookAutomation an OleAumation for accessing the functionalities of the embedded workbook
	 * @param annotation the range annotation for which the data need to be deleted  
	 * @param permanentDelete if true the annotation data will be deleted permanently, otherwise they will just be hidden  
	 */
	public static void deleteRangeAnnotationData(OleAutomation workbookAutomation, RangeAnnotation annotation){		
		deleteDataRows(workbookAutomation, "Annotation.Name", annotation.getName());
	}
	
	/**
	 * Delete the filtered rows of data
	 * @param workbookAutomation an OleAumation for accessing the functionalities of the embedded workbook
	 * @param fieldToFilter a string the represents the name of the field filter will apply
	 * @param value a string that represents the value to filter
	 * @param permanentDelete true to permanently delete the row, false to just hide it
	 */
	private static void deleteDataRows(OleAutomation workbookAutomation, String fieldToFilter, String value){
					
		// get the OleAutomation object for the sheet that stores the annotation metadata (a.k.a. annotation data sheet) 
		OleAutomation annotationDataSheetBeforeFilter = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		
		// unprotect the annotation data sheet
		WorksheetUtils.unprotectWorksheet(annotationDataSheetBeforeFilter);
		
		// determine the position of the field that represents the name of the annotation
		OleAutomation topLeftCellAuto = WorksheetUtils.getRangeAutomation(annotationDataSheetBeforeFilter, startColumn+""+startRow, null);
		int columnIndex = RangeUtils.getFirstColumnIndex(topLeftCellAuto);
		topLeftCellAuto.dispose();
		int fieldRelativePosition = fields.get(fieldToFilter);
		int fieldIndex = columnIndex + fieldRelativePosition;
		
		// get the range that contains the annotation data together with the header row
		OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationDataSheetBeforeFilter);
		annotationDataSheetBeforeFilter.dispose();	// TODO: It seems this line occasionally throws SWTError native exception: 0xc0000005
		
		// filter this range to get only those rows relevant to the specified sheet 
		RangeUtils.filterRange(usedRange, fieldIndex, value); 
				
		// get the range OleAutomation for the filtered results. This range still includes the header row.  
		OleAutomation filteredRange = RangeUtils.getSpecialCells(usedRange, 12); // xlCellTypeVisible = 12  (visible cells)
		usedRange.dispose();
		
		// get all the areas in the filtered range 
		OleAutomation areasAuto = RangeUtils.getAreas(filteredRange);
		filteredRange.dispose();
		int countAreas = CollectionsUtils.countItemsInCollection(areasAuto);
				
		// get all rows (indices) in the filtered range. exclude the header row 
		ArrayList<Integer> filteredRows = new ArrayList<Integer>();
		for (int j = 1; j <=countAreas; j++) {
			OleAutomation area = CollectionsUtils.getItemByIndex(areasAuto, j, false);
			OleAutomation rowsAuto = RangeUtils.getRangeRows(area);
			area.dispose();
			int countRows = CollectionsUtils.countItemsInCollection(rowsAuto);
			for (int i = 1; i <=countRows; i++) {				
				if(!(j==1 && i==1)){ // j==1 && i==1 corresponds to the header row
					OleAutomation row =  CollectionsUtils.getItemByIndex(rowsAuto, i, false);
					int rowIndex = RangeUtils.getFirstRowIndex(row);
					row.dispose();
					filteredRows.add(rowIndex);
				}
			}
			rowsAuto.dispose();
		}
		areasAuto.dispose();
		
		// Create OleAutomation after filtering the range 
		OleAutomation annotationDataSheetAfterFilter = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		// Show before deleting
		WorksheetUtils.showAllWorksheetData(annotationDataSheetAfterFilter);
		// Delete filtered rows 
		if(filteredRows.size()>0){			
			OleAutomation applicationAuto = WorkbookUtils.getApplicationAutomation(workbookAutomation);
			OleAutomation filteredRowsAuto = null;
			//String multiSelectionRange = "";
			for (int i = 0; i < filteredRows.size(); i++) {
				int rowIndex = filteredRows.get(i);
				//String rowAddress = "$"+rowIndex+":$"+rowIndex;
				
				OleAutomation nextFilteredRowAuto = WorksheetUtils.getRow(annotationDataSheetAfterFilter, rowIndex);
				if(filteredRowsAuto!=null){
					filteredRowsAuto = ApplicationUtils.getUnion(applicationAuto, filteredRowsAuto, nextFilteredRowAuto);
				}else{
					filteredRowsAuto = nextFilteredRowAuto;
				}
				//multiSelectionRange = multiSelectionRange.concat(rowAddress+",");
			}
			applicationAuto.dispose();
			
			//multiSelectionRange = multiSelectionRange.substring(0, (multiSelectionRange.length()-1));	
			//OleAutomation filteredRowsAuto = WorksheetUtils.getMultiSelectionRangeAutomation(annotationDataSheetAfterFilter, multiSelectionRange);
			RangeUtils.deleteRange(filteredRowsAuto);
			filteredRowsAuto.dispose();
		}
		
		// protect the worksheet from further user alterations 
		WorksheetUtils.protectWorksheet(annotationDataSheetAfterFilter);
		annotationDataSheetAfterFilter.dispose();
	}

	/**
	 * Delete all annotation data. This method clears all the values from the sheet 
	 * where the annotation data are stored, except of the header row. 
	 * @param workbookAutomation
	 */
	public static void deleteAllRangeAnnotationData(OleAutomation workbookAutomation){
		
		// get the OleAutomation object for the sheet that stores the annotation metadata (a.k.a. annotation data sheet) 
		OleAutomation annotationDataSheetBeforeDelete = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		
		// unprotect the worksheet in order to perform the following actions
		boolean isUnprotected= WorksheetUtils.unprotectWorksheet(annotationDataSheetBeforeDelete);
		if(!isUnprotected){
			int style = SWT.ICON_ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("Annotation Data Sheet could not be unprotected!");
			message.open();
			return;
		}
		
		// find the last row that contains data
		OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationDataSheetBeforeDelete);
		String address = RangeUtils.getRangeAddress(usedRange);
		int position = address.indexOf(":");
		String downRightCell = address.substring(position+1).replace("$", "");
		usedRange.dispose();
		
		// delete all the rows except of the one header
		OleAutomation rangeToDelete = WorksheetUtils.getRangeAutomation(annotationDataSheetBeforeDelete, startColumn+""+(startRow+1), downRightCell);
		RangeUtils.deleteRange(rangeToDelete);
		rangeToDelete.dispose();
		
		annotationDataSheetBeforeDelete.dispose();
		
		// protect the worksheet from further user manipulation 
		OleAutomation annotationDataSheetAfterDelete = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		WorksheetUtils.protectWorksheet(annotationDataSheetAfterDelete);
		annotationDataSheetAfterDelete.dispose();		
	}
	
	
	/**
	 * Export annotation metadata as CSV
	 * @param workbookAutomation
	 * @param directoryPath
	 * @param fileName
	 * @return
	 */
	public static boolean exportRangeAnnotationsAsCSV(OleAutomation workbookAutomation, String directoryPath, String fileName){
		
		// get the OleAutomation object for the worksheet where the annotation data are stored
		OleAutomation annotationDataSheet = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		
		// check if the annotation data sheet exists
		if(annotationDataSheet==null){		
			//System.out.println("Annotation Data Sheet not found!");
			return false;
		}
		
		// check if there are annotation data 
		OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationDataSheet);
		if(usedRange==null){
			//System.out.println("Used range is null!");
			return false;
		}
		
		// there should be more than two rows of annotation data, 
		// otherwise will export only the header
		OleAutomation rows = RangeUtils.getRangeRows(usedRange);
		if(CollectionsUtils.countItemsInCollection(rows)<2) {
			//System.out.println("There are no annotation data, just the header!");
			return false;
		}
		usedRange.dispose();
		rows.dispose();
		
		// unprotect worksheet in order allow export
		WorksheetUtils.unprotectWorksheet(annotationDataSheet);
		
		// export annotation data
		int index = fileName.lastIndexOf('.');		
		String nameWithoutExtension = fileName.substring(0, index);
		String annotationDataFile = directoryPath+"\\"+nameWithoutExtension+"_annotation_data";
		boolean isSuccess = WorksheetUtils.saveAsCSV(annotationDataSheet, annotationDataFile);
		
		// protect worksheet from further user manipulation
		WorksheetUtils.protectWorksheet(annotationDataSheet);
		
		return isSuccess;
	}
	
	
	/**
	 * Hide/Show the worksheet that stores the annotation (meta-)data
	 * @param embeddedWorkbook an OleAutomation that is used to access the functionalities of the workbook that is currently embedded by the application
	 * @param visible true to show the worksheet, false to hide it
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean setVisibility(OleAutomation embeddedWorkbook, boolean visible){
		
		OleAutomation rangeAnnotationsDataSheet = WorkbookUtils.getWorksheetAutomationByName(embeddedWorkbook, name);
		
		if(rangeAnnotationsDataSheet==null)
			return false; 
		
		boolean result = WorksheetUtils.setWorksheetVisibility(rangeAnnotationsDataSheet, visible);
		rangeAnnotationsDataSheet.dispose();
		return result;
	}
	
	
	/**
	 * Protect the annotation data sheet
	 * @param embeddedWorkbook an OleAutomation that is used to access the functionalities of the workbook that is currently embedded by the application
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean protect(OleAutomation embeddedWorkbook){
		
		OleAutomation rangeAnnotationsDataSheet = WorkbookUtils.getWorksheetAutomationByName(embeddedWorkbook, name);
		
		if(rangeAnnotationsDataSheet==null)
			return false; 
		
		boolean result = WorksheetUtils.protectWorksheet(rangeAnnotationsDataSheet);
		rangeAnnotationsDataSheet.dispose();
		return result;
	}
	
	/**
	 * Unprotect the annotation data sheet
	 * @param embeddedWorkbook an OleAutomation that is used to access the functionalities of the workbook that is currently embedded by the application
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean unprotect(OleAutomation embeddedWorkbook){
		
		OleAutomation rangeAnnotationsDataSheet = WorkbookUtils.getWorksheetAutomationByName(embeddedWorkbook, name);
		
		if(rangeAnnotationsDataSheet==null)
			return false; 
		
		boolean result = WorksheetUtils.unprotectWorksheet(rangeAnnotationsDataSheet);
		rangeAnnotationsDataSheet.dispose();
		return result;
	}

	
	public static boolean delete(OleAutomation embeddedWorkbook){
		
		OleAutomation rangeAnnotationsDataSheet = WorkbookUtils.getWorksheetAutomationByName(embeddedWorkbook, name);
		
		if(rangeAnnotationsDataSheet==null)
			return false; 
		
		boolean result = WorksheetUtils.deleteWorksheet(rangeAnnotationsDataSheet);
		rangeAnnotationsDataSheet.dispose();
		return result;	
	}
	

	/**
	 * @return the name
	 */
	public static String getName() {
		return name;
	}


	/**
	 * @return the startColumn
	 */
	public static String getStartColumn() {
		return startColumn;
	}


	/**
	 * @return the startRow
	 */
	public static int getStartRow() {
		return startRow;
	}


	/**
	 * @return the fields
	 */
	public static LinkedHashMap<String, Integer> getFields() {
		return fields;
	}
	
}
