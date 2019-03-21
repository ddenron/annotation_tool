/**
 * 
 */
package de.tudresden.annotator.annotations.utils;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

import org.eclipse.swt.SWT;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.widgets.MessageBox;

import de.tudresden.annotator.annotations.NotApplicableStatus;
import de.tudresden.annotator.annotations.WorkbookAnnotation;
import de.tudresden.annotator.annotations.WorksheetAnnotation;
import de.tudresden.annotator.main.Launcher;
import de.tudresden.annotator.oleutils.CollectionsUtils;
import de.tudresden.annotator.oleutils.RangeUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;
import de.tudresden.annotator.oleutils.WorksheetUtils;

/**
 * @author Elvis Koci
 */
public class AnnotationStatusSheet {
	
	protected static final String name = "Annotation_Status_Data";
	private static final String startColumnChar = "A";
	private static final int startColumnIndex = 1;
	private static final int lastColumnIndex = startColumnIndex + 3 + (NotApplicableStatus.values().length - 1); 
	private static final int startRow = 1; 
	
	private static List<NotApplicableStatus> naStatusFields = new ArrayList<NotApplicableStatus>(); 
	private static boolean hasAmbiguousField = false;
	
	
	/**
	 * Save the annotation status for the workbook and worksheets
	 * @param workbookAutomation an OleAutomation to access the embedded workbook
	 */
	public static void saveAnnotationStatuses(OleAutomation workbookAutomation){
			
		OleAutomation annotationStatusSheet =  WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		
		if(annotationStatusSheet==null){		
			annotationStatusSheet = createAnnotationStatusSheet(workbookAutomation);	
		}else{
			
			WorksheetUtils.unprotectWorksheet(annotationStatusSheet);
			
			// delete all the existing data from the worksheet. 
			// by removing all existing data we ensure that the "new" data will have
			// the right format. So, they are not effected by the existing data.
			OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationStatusSheet);		
			RangeUtils.deleteRange(usedRange);	
			usedRange.dispose();
			
			// re-create the header
			createHeaderRow(annotationStatusSheet);
			WorksheetUtils.protectWorksheet(annotationStatusSheet);	
		}
		
		writeStatuses(annotationStatusSheet);		
		annotationStatusSheet.dispose();
	}
	
	/**
	 * Read the annotation status data stored in the sheet.
	 * @param workbookAutomation an OleAutomation to access the embedded workbook functionalities
	 */
	public static void readAnnotationStatuses(OleAutomation workbookAutomation){
		
		// get the automation for the annotation_status sheet
		OleAutomation annotationStatusSheet =  WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		
		// create the base in memory structure to save the annotation objects 
		AnnotationHandler.createBaseAnnotations(workbookAutomation);
		
		// if the sheet that stores the annotation statuses does not exist, return 
		// there are no data to read. the default status applies
		if(annotationStatusSheet==null){		
			return;
		}
		
		// ensure that the annotation_status sheet has the expected format
		if(!validateAnnotationStatusSheet(annotationStatusSheet)){
			return;
		}
		
		// read the workbook annotation status data 
		if(!readWorkbookAnnotationStatus(annotationStatusSheet)){
			return;
		}
		
		// read the annotation status data for each sheet
		readWorksheetAnnotationsStatuses(annotationStatusSheet);

	}
	
	/**
	 * Create the sheet that will store the status of the annotation for the workbook and the individual worksheets 
	 * @param workbookAutomation an OleAutomation to access the embedded workbook functionalities
	 * @return the OleAutomation of the created Annotation_Status sheet
	 */
	private static OleAutomation createAnnotationStatusSheet(OleAutomation workbookAutomation){
		
		WorkbookUtils.unprotectWorkbook(workbookAutomation);
		
		OleAutomation annotationStatusSheet = WorkbookUtils.addWorksheetAsLast(workbookAutomation);
		WorksheetUtils.setWorksheetName(annotationStatusSheet, name);
		
		createHeaderRow(annotationStatusSheet);
		
		WorksheetUtils.setWorksheetVisibility(annotationStatusSheet, false);
		WorkbookUtils.protectWorkbook(workbookAutomation, true, false);
		
		return annotationStatusSheet;
	}
	
	
	/**
	 * Create (write) the header row that contains the field names 
	 * @param annotationStatusSheet  an OleAutomation that provides access to the sheet that maintains annotation status data
	 */
	private static void createHeaderRow(OleAutomation annotationStatusSheet){
		
		/*
		 * The first three column headers are required 
		 * for backwards compatibility with previous versions
		 * of the SpreadsheetAnnotator tool.
		 */
		OleAutomation entryName = WorksheetUtils.getCell(annotationStatusSheet, startRow, startColumnIndex);
		RangeUtils.formatCells(entryName, "@");
		RangeUtils.setValue(entryName, "Name");
		entryName.dispose();
		
		OleAutomation isComplete = WorksheetUtils.getCell(annotationStatusSheet, startRow, startColumnIndex+1);
		RangeUtils.formatCells(isComplete, "@");
		RangeUtils.setValue(isComplete, "Completed");
		isComplete.dispose();
		
		OleAutomation isNotApplicable = WorksheetUtils.getCell(annotationStatusSheet, startRow, startColumnIndex+2);
		RangeUtils.formatCells(isNotApplicable, "@");
		RangeUtils.setValue(isNotApplicable, "NotApplicable");
		isNotApplicable.dispose();
				
		int columnIx = startColumnIndex+3; 
		for(NotApplicableStatus naStatus: NotApplicableStatus.values()){		
			if (naStatus != NotApplicableStatus.NONE){
				OleAutomation nextField = WorksheetUtils.getCell(annotationStatusSheet, startRow, columnIx);
				RangeUtils.formatCells(nextField, "@");
				RangeUtils.setValue(nextField, naStatus.getStatusName());
				nextField.dispose();
				columnIx++;
			}
		}
		
		OleAutomation isAmbiguous = WorksheetUtils.getCell(annotationStatusSheet, startRow, columnIx);
		RangeUtils.formatCells(isAmbiguous, "@");
		RangeUtils.setValue(isAmbiguous, "Ambiguous");
		isAmbiguous.dispose();
	}
	
	
	/**
	 * Write the status data of the Workbook and Worksheet annotations for the embedded excel file 
	 * @param annotationStatusSheet an OleAutomation that provides access to the sheet that maintains the annotation status data
	 */
	private static void writeStatuses(OleAutomation annotationStatusSheet){
		
		WorksheetUtils.unprotectWorksheet(annotationStatusSheet);
		
		// write the annotation status (data) of the workbook annotation
		WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();

		OleAutomation bookName = WorksheetUtils.getCell(annotationStatusSheet, startRow+1, startColumnIndex);
		RangeUtils.formatCells(bookName, "@");
		RangeUtils.setValue(bookName, "Workbook");
		bookName.dispose();
		
		
		if(workbookAnnotation.isCompleted()){
			OleAutomation isBookCompleted = WorksheetUtils.getCell(annotationStatusSheet, startRow+1, startColumnIndex+1);
			RangeUtils.setValue(isBookCompleted, "true");
			isBookCompleted.dispose();
		}
		
		if(workbookAnnotation.isNotApplicable()){
			OleAutomation isBookNotApplicable = WorksheetUtils.getCell(annotationStatusSheet, startRow+1, startColumnIndex+2);
			RangeUtils.setValue(isBookNotApplicable, "true");
			isBookNotApplicable.dispose();
		}
			
		writeNotApplicableStatuses(annotationStatusSheet, workbookAnnotation.getNotApplicableStatus(), startRow+1, startColumnIndex+3);
		if(workbookAnnotation.isAmbiguous()){
			OleAutomation isBookAmbiguous = WorksheetUtils.getCell(annotationStatusSheet, startRow+1, lastColumnIndex);
			RangeUtils.setValue(isBookAmbiguous, "true");
			isBookAmbiguous.dispose();
		}
		
		// write the annotation status (data) of each worksheet annotation
		Collection<WorksheetAnnotation> collection = workbookAnnotation.getWorksheetAnnotations().values(); 
		WorksheetAnnotation[] worksheetAnnotations = collection.toArray(new WorksheetAnnotation[collection.size()]);
		
		int rowIndex = startRow+2;
		for (WorksheetAnnotation worksheetAnnotation : worksheetAnnotations) {
			
			OleAutomation sheetName = WorksheetUtils.getCell(annotationStatusSheet, rowIndex, startColumnIndex);
			RangeUtils.formatCells(sheetName, "@");
			RangeUtils.setValue(sheetName, worksheetAnnotation.getSheetName());
			sheetName.dispose();
			
			
			if(worksheetAnnotation.isCompleted()){
				OleAutomation isSheetCompleted = WorksheetUtils.getCell(annotationStatusSheet, rowIndex, startColumnIndex+1);
				RangeUtils.setValue(isSheetCompleted, "true");
				isSheetCompleted.dispose();
			}
			
			if(worksheetAnnotation.isNotApplicable()){
				OleAutomation isSheetNotApplicable = WorksheetUtils.getCell(annotationStatusSheet, rowIndex, startColumnIndex+2);
				RangeUtils.setValue(isSheetNotApplicable, "true");
				isSheetNotApplicable.dispose();
			}
			
			writeNotApplicableStatuses(annotationStatusSheet, worksheetAnnotation.getNotApplicableStatus(), rowIndex, startColumnIndex+3);
			if(worksheetAnnotation.isAmbiguous()){
					OleAutomation isSheetAmbiguous = WorksheetUtils.getCell(annotationStatusSheet, rowIndex, lastColumnIndex);
					RangeUtils.setValue(isSheetAmbiguous, "true");
					isSheetAmbiguous.dispose();
			}			
			rowIndex++;
		}
		WorksheetUtils.protectWorksheet(annotationStatusSheet);
	}
	
	
	/**
	 * 
	 * @param annotationStatusSheet
	 * @param activeStatus
	 * @param rowIx
	 * @param columnIx
	 */
	private static void writeNotApplicableStatuses(OleAutomation annotationStatusSheet, NotApplicableStatus activeStatus, int rowIx, int columnIx){	
		int currentColumnIx = columnIx; 
		for(NotApplicableStatus naStatus: NotApplicableStatus.values()){		
			if (naStatus != NotApplicableStatus.NONE){		
				if(activeStatus==naStatus){
					OleAutomation nextField = WorksheetUtils.getCell(annotationStatusSheet, rowIx, currentColumnIx);
					RangeUtils.setValue(nextField, "true");
					nextField.dispose();	
				}		
				currentColumnIx++;
			}
		}
	}
	
	
	/**
	 * Validate the "Annotation_Status" sheet to ensure that the data are in the expected format. 
	 * @param annotationStatusSheet an OleAutomation that provides access to the sheet that maintains the annotation status data
	 * @return true if the status data are in the expected format, false otherwise
	 */
	private static boolean validateAnnotationStatusSheet(OleAutomation annotationStatusSheet){
		
		OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationStatusSheet);
		if(usedRange==null){
			int style = SWT.ICON_ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("Could not recover the annotation status from the previous session. "
					+ "The \"Annotation_Status\" sheet is empty.");
			message.open();
			return false;
		}
		
		
		OleAutomation rangeColumns = RangeUtils.getRangeColumns(usedRange);
		int countColumns = CollectionsUtils.countItemsInCollection(rangeColumns);
		rangeColumns.dispose();
		
		int defaultNrColumns = 3;
		if(countColumns < defaultNrColumns){
			usedRange.dispose();
						
			int style = SWT.ICON_ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("Could not recover the annotation status from the previous session. "
					+ "The annotation status data are not in the expected format. "
					+ "The expected number of columns is at least "+defaultNrColumns+". "
					+ "Instead, the application found "+countColumns+" column/s.");
			message.open();	
			return false;
		}
				
		OleAutomation rangeRows = RangeUtils.getRangeRows(usedRange);
		OleAutomation headerRow = CollectionsUtils.getItemByIndex(rangeRows, 1, false);
		rangeRows.dispose();
		usedRange.dispose();
		
		String values[] = RangeUtils.getRangeValues(headerRow);	
		boolean hasDefaultHeaders =	(values[0].compareToIgnoreCase("Name")==0 &&
				values[1].compareToIgnoreCase("Completed")==0 &&
				values[2].compareToIgnoreCase("NotApplicable")==0);
		
		
		if(!hasDefaultHeaders){		
			int style = SWT.ICON_ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("Could not recover the annotation status from the previous session. "
					+ "The header row does not contain the expected default fields.");
			message.open();
			return false;
		}
		
		if(values.length>3){
			for(int i=3; i<values.length; i++){
				if(values[i].compareToIgnoreCase("Ambiguous")==0){
					if(i==values.length-1){
						hasAmbiguousField = true;
					}else{
						int style = SWT.ICON_ERROR;
						MessageBox message = Launcher.getInstance().createMessageBox(style);
						message.setMessage("Could not recover the annotation status from the previous session. "
								+ "The annotation status data are not in the expected format. "
								+ "The field \"Ambiguous\" is expected to be the last one.");
						message.open();					
						return false;
					}			
				}else{
					NotApplicableStatus status = NotApplicableStatus.getByName(values[i]);
					if(status != null){
						naStatusFields.add(status);
					}else{
						int style = SWT.ICON_WARNING;
						MessageBox message = Launcher.getInstance().createMessageBox(style);
						message.setMessage("Could not recognize the status field: \"" + values[i]+"\"");
						message.open();		
					}
				}
			}
		}			
		return true;
	}
	
	
	/**
	 * Read the workbook annotation status data
	 * @param annotationStatusSheet an OleAutomation that provides access to the sheet that maintains the annotation status data
	 * @return true if the status data were successfully read, false otherwise
	 */
	private static boolean readWorkbookAnnotationStatus(OleAutomation annotationStatusSheet){
		
		WorkbookAnnotation wa = AnnotationHandler.getWorkbookAnnotation();
		String startCell = "$"+startColumnChar+"$"+(startRow+1);
		OleAutomation endCellAuto = WorksheetUtils.getCell(annotationStatusSheet, (startRow+1), lastColumnIndex);
		String endCell = RangeUtils.getRangeAddress(endCellAuto);
		endCellAuto.dispose();
		
		OleAutomation workbookStatusRow = WorksheetUtils.getRangeAutomation(annotationStatusSheet, startCell, endCell);
		String[] values= RangeUtils.getRangeValues(workbookStatusRow);
		workbookStatusRow.dispose();
		
		if(values[0].compareToIgnoreCase("Workbook")!=0){
			int style = SWT.ICON_ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("Could not recover the annotation status from the previous session. "
					+ "Could not recognize the values of 1st field. Expected \"Workbook\".");
			message.open();
		}
		
		if(!validateRowData(values)){
			return false;
		}
		
		// some version of  excel translate boolean TRUE as integer -1
		if(values[1].compareToIgnoreCase("-1")==0 || values[1].compareToIgnoreCase("true")==0){
			wa.setCompleted(true);	
		}

		if(values[2].compareToIgnoreCase("-1")==0 || values[2].compareToIgnoreCase("true")==0){
			wa.setNotApplicable(true);
			
			int len = values.length;
			if(hasAmbiguousField){
				len = len-1;
			}
			 
			int otherIx = 0;
			for(int i=3; i<len; i++){
				if(values[i].compareToIgnoreCase("-1")==0 || values[i].compareToIgnoreCase("true")==0){		
					wa.setNaStatus(naStatusFields.get(otherIx));
				}
				otherIx++;
			}
		}
		
		if(hasAmbiguousField){
			if(values[values.length-1].compareToIgnoreCase("-1")==0 || values[values.length-1].compareToIgnoreCase("true")==0){		
				wa.setAmbiguous(true);
			}
		}
			
		return true;	
	}

	/**
	 * Read the status data for the worksheet annotations
	 * @param annotationStatusSheet an OleAutomation that provides access to the sheet that maintains the annotation status data
	 * @return true if the status data were successfully read, false otherwise
	 */
	private static boolean readWorksheetAnnotationsStatuses(OleAutomation annotationStatusSheet){
		
		OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationStatusSheet);		
		OleAutomation rangeRows = RangeUtils.getRangeRows(usedRange);
		int countRows = CollectionsUtils.countItemsInCollection(rangeRows);		
		usedRange.dispose();
		
		OleAutomation topRightCellAuto = WorksheetUtils.getCell(annotationStatusSheet, (startRow), lastColumnIndex);
		String topRightCell= RangeUtils.getRangeAddress(topRightCellAuto);
		topRightCellAuto.dispose();
		String endColumnChar = topRightCell.replaceAll("[0-9\\$]+",""); 
		
		int rowIndex = startRow+2;

		WorkbookAnnotation wa = AnnotationHandler.getWorkbookAnnotation();
		while(rowIndex<=countRows) {
			
			String startCell = "$"+startColumnChar+"$"+rowIndex;
			String endCell = "$"+endColumnChar+"$"+rowIndex;
			OleAutomation sheetStatusRow = WorksheetUtils.getRangeAutomation(annotationStatusSheet, startCell, endCell);
			String[] values= RangeUtils.getRangeValues(sheetStatusRow);

			if(!validateRowData(values)){
				return false;
			}
			
			
			String sheetName = values[0];
			WorksheetAnnotation sheetAnnotation = wa.getWorksheetAnnotations().get(sheetName);
			if(sheetAnnotation!=null){
				
				// some version of  excel translate boolean TRUE as integer -1
				if(values[1].compareToIgnoreCase("-1")==0 || values[1].compareToIgnoreCase("true")==0){
					sheetAnnotation.setCompleted(true);
				}

				if(values[2].compareToIgnoreCase("-1")==0 || values[2].compareToIgnoreCase("true")==0){
					sheetAnnotation.setNotApplicable(true);
					int len = values.length;
					if(hasAmbiguousField){
						len = len-1;
					}
					 
					int otherIx = 0;
					for(int i=3; i<len; i++){
						if(values[i].compareToIgnoreCase("-1")==0 || values[i].compareToIgnoreCase("true")==0){		
							sheetAnnotation.setNaStatus(naStatusFields.get(otherIx));
						}
						otherIx++;
					}
				}
				
				if(hasAmbiguousField){
					if(values[values.length-1].compareToIgnoreCase("-1")==0 || values[values.length-1].compareToIgnoreCase("true")==0){		
						sheetAnnotation.setAmbiguous(true);
					}
				}								
			}
			rowIndex++;		
		}		
		return true;
	}
	
	/**
	 * Validate the row data. 
	 * @param data an array of string values that were read from the worksheet range - row
	 * @return true if the data pass all the validation tests, false otherwise.
	 */
	private static boolean validateRowData(String data[]){
		
		if(data.length<3){
			int style = SWT.ICON_ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("Could not recover the annotation status from the previous session. "
					+ "The data are not in the expected format!");
			message.open();
			return false;
		}
		
		int countTrue = 0;
		boolean isCompleteTrue = false;
		boolean isNotApplicableTrue = false;
		boolean isAmbiguousTrue = false;
		for (int i=1; i<data.length; i++){
			
			if(data[i]==null || data[i]==""){
				continue;
			}
			
			if(!(data[i].compareToIgnoreCase("0")==0  || data[i].compareToIgnoreCase("-1")==0 
					|| data[i].compareToIgnoreCase("true")==0 || data[i].compareToIgnoreCase("false")==0)){
		
				int style = SWT.ICON_ERROR;
				MessageBox message = Launcher.getInstance().createMessageBox(style);
				message.setMessage("Could not recover the annotation status from the previous session. "
						+ "Could not recognize one or more values of the "+(i==2?"2nd":(i==3?"3rd":i+"th"))+" field.");
				message.open();
				return false;
			}else{			
				if(data[i].compareToIgnoreCase("-1")==0 || data[i].compareToIgnoreCase("true")==0){	
					countTrue++;
					
					if(i==1){
						isCompleteTrue = true;
					}else if(i==2){
						isNotApplicableTrue = true;
					}else if(i>2 && i==data.length-1){
						isAmbiguousTrue = true;
					}
				}
			}
		}
		
		if(((isNotApplicableTrue && !isCompleteTrue && !isAmbiguousTrue) && countTrue>2) || (!isNotApplicableTrue && countTrue>1)){
			int style = SWT.ICON_ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("Only one of the annotation statuses can be true at a time! "
					+ "Instead, found "+countTrue+" true statuses!");
			message.open();
		}
		
		return true;
	}
	
	
	/**
	 * Hide/Show the sheet that stores the annotation statuses
	 * @param embeddedWorkbook an OleAutomation that is used to access the functionalities of the workbook that is currently embedded by the application
	 * @param visible true to show the sheet, false to hide it
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean setVisibility(OleAutomation embeddedWorkbook, boolean visible){
		
		OleAutomation annotationStatusSheet = WorkbookUtils.getWorksheetAutomationByName(embeddedWorkbook, name);
		
		if(annotationStatusSheet==null)
			return false; 
		
		boolean result = WorksheetUtils.setWorksheetVisibility(annotationStatusSheet, visible);
		annotationStatusSheet.dispose();
		return result;
	}
	
	
	/**
	 * Protect the sheet that stores the annotation status
	 * @param embeddedWorkbook an OleAutomation that is used to access the functionalities of the workbook that is currently embedded by the application
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean protect(OleAutomation embeddedWorkbook){
		
		OleAutomation annotationStatusSheet = WorkbookUtils.getWorksheetAutomationByName(embeddedWorkbook, name);
		
		if(annotationStatusSheet==null)
			return false; 
		
		boolean result = WorksheetUtils.protectWorksheet(annotationStatusSheet);
		annotationStatusSheet.dispose();
		return result;
	}
	
	/**
	 * Unprotect the sheet that stores the annotation status
	 * @param embeddedWorkbook an OleAutomation that is used to access the functionalities of the workbook that is currently embedded by the application
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean unprotect(OleAutomation embeddedWorkbook){
		
		OleAutomation annotationStatusSheet = WorkbookUtils.getWorksheetAutomationByName(embeddedWorkbook, name);
		
		if(annotationStatusSheet==null)
			return false; 
		
		boolean result = WorksheetUtils.unprotectWorksheet(annotationStatusSheet);
		annotationStatusSheet.dispose();
		return result;
	}
	
	
	/**
	 * Delete the sheet that stores the annotation statuses
	 * @param embeddedWorkbook an OleAutomation that is used to access the functionalities of the workbook that is currently embedded by the application
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean delete(OleAutomation embeddedWorkbook){
		
		OleAutomation annotationStatusSheet = WorkbookUtils.getWorksheetAutomationByName(embeddedWorkbook, name);
		
		if(annotationStatusSheet==null)
			return false; 
		
		boolean result = WorksheetUtils.deleteWorksheet(annotationStatusSheet);
		annotationStatusSheet.dispose();
		return result;	
	}
	
	/**
	 * @return the name
	 */
	public static String getName() {
		return name;
	}
		
}
