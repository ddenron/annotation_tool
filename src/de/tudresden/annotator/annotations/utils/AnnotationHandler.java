/**
 * 
 */
package de.tudresden.annotator.annotations.utils;

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.TreeMap;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.util.CellRangeAddress;
import org.eclipse.swt.SWT;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.widgets.MessageBox;

import de.tudresden.annotator.annotations.RangeAnnotation;
import de.tudresden.annotator.annotations.WorkbookAnnotation;
import de.tudresden.annotator.annotations.WorksheetAnnotation;
import de.tudresden.annotator.annotations.format.AnnotationClass;
import de.tudresden.annotator.annotations.format.AnnotationShape;
import de.tudresden.annotator.main.GUIListeners;
import de.tudresden.annotator.main.Launcher;
import de.tudresden.annotator.oleutils.ApplicationUtils;
import de.tudresden.annotator.oleutils.CharactersUtils;
import de.tudresden.annotator.oleutils.CollectionsUtils;
import de.tudresden.annotator.oleutils.ColorFormatUtils;
import de.tudresden.annotator.oleutils.FillFormatUtils;
import de.tudresden.annotator.oleutils.FontUtils;
import de.tudresden.annotator.oleutils.LineFormatUtils;
import de.tudresden.annotator.oleutils.RangeUtils;
import de.tudresden.annotator.oleutils.ShadowFormatUtils;
import de.tudresden.annotator.oleutils.ShapeUtils;
import de.tudresden.annotator.oleutils.TextFrameUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;
import de.tudresden.annotator.oleutils.WorksheetFunctionUtils;
import de.tudresden.annotator.oleutils.WorksheetUtils;

/**
 * This class handles all the operations related to the creation and management of annotations 
 * @author Elvis Koci
 */
public class AnnotationHandler {

	/**
	 * This object stores and provides access to all annotations that are created in the embedded workbook  
	 */
	private static final WorkbookAnnotation workbookAnnotation = new WorkbookAnnotation();	
	private static int oldWorkbookAnnotationHash;

	/**
	 * Maintains the list of all range annotations that can be undone
	 */
	private static final LinkedList<RangeAnnotation> undoList = new LinkedList<RangeAnnotation>();
	
	/**
	 * Maintains the list of all range annotations that can be re-done after they have first been undone
	 */
	private static final LinkedList<RangeAnnotation> redoList = new LinkedList<RangeAnnotation>();
	
	
	private static final Logger logger = LogManager.getLogger(GUIListeners.class.getName());
	
	/**
	 * Set up the base structure for keeping the annotation data in memory. This means having a WorkbookAnnotation 
	 * object for the embedded workbook, as well as an WorksheetAnnotation object for each sheet in the workbook.  
	 * @param workbookAutomation an OleAutomation for accessing the functionalities of the embedded workbook
	 */
	public static void createBaseAnnotations(OleAutomation workbookAutomation){
		
		// save on memory metadata about the annotation
		String workbookName = WorkbookUtils.getWorkbookName(workbookAutomation);
		workbookAnnotation.setWorkbookName(workbookName);
		
		workbookAnnotation.setCompleted(false);
		workbookAnnotation.setNotApplicable(false);
		workbookAnnotation.setAmbiguous(false);
		
		OleAutomation worksheetsAutomation = WorkbookUtils.getWorksheetsAutomation(workbookAutomation);
		if(worksheetsAutomation==null)
			return;
		
		int i = 1;
		while (true) {
			OleAutomation worksheet = CollectionsUtils.getItemByIndex(worksheetsAutomation, i++, false);		
			if(worksheet==null)
				break;
			
			String name = WorksheetUtils.getWorksheetName(worksheet);
			int index = WorksheetUtils.getWorksheetIndex(worksheet);
			
			if(name.compareToIgnoreCase(RangeAnnotationsSheet.getName())==0 ||
			   name.compareToIgnoreCase(AnnotationStatusSheet.getName())==0 ||
			   name.compareToIgnoreCase(UsageLogSheet.getName())==0){
				continue;
			}
				
			WorksheetAnnotation wa = new WorksheetAnnotation(name, index);
			workbookAnnotation.getWorksheetAnnotations().put(name, wa);
		}		
	}
	
	/**
	 * Create again (reproduce) the given list of range annotations and add them in the in-memory structure
	 * @param workbookAutomation an OleAutomation for accessing the functionalities of the embedded workbook
	 * @return true if all the range annotations were successfully re-created, false otherwise
	 */
	public static void recreateRangeAnnotations(OleAutomation workbookAutomation, RangeAnnotation[] rangeAnnotations){	
		
		WorkbookUtils.unprotectAllWorksheets(workbookAutomation);			
		
		for (int i=0; i< rangeAnnotations.length; i++) {		
			
			boolean result = false;
			try{
				result = AnnotationHandler.drawRangeAnnotation(workbookAutomation, rangeAnnotations[i], true);
			}catch (Exception ex){
				logger.error("Generic error on drawing range annotation object", ex);
			}
			
			if(result){
				workbookAnnotation.addRangeAnnotation(rangeAnnotations[i]);
			}
		}
		
		WorkbookUtils.protectAllWorksheets(workbookAutomation);
	}
	
	/**
	 * Annotate the selected ranges (areas) of cells 
	 * @param workbookAutomation an OleAutomation for accessing the functionalities of the embedded workbook
	 * @param worksheetFunction an OleAutomation that provide access to the various excel functions (Ex. sum, count)
	 * @param sheetName the name of the sheet that contains the selected range 
	 * @param sheetIndex the index of the sheet that contains the selected range
	 * @param selectedAreas an array of strings that represent all the selected ranges (areas)   
	 * @param annotationClass the class to use for the annotation
	 * @return an AnnotationResult object that holds information about the result from the execution of this method
	 */
	public static void annotate(OleAutomation workbookAutomation, String sheetName, int sheetIndex,
								String selectedAreas[], AnnotationClass annotationClass) {
	
		if(selectedAreas==null){
			MessageBox messageBox = Launcher.getInstance().createMessageBox(SWT.ICON_ERROR);
            messageBox.setMessage( "You have not selected a range! "
					+ "Please, first do a selection.");
            messageBox.open();
            return;
		}
		
		// get the OleAutomation object for the worksheet using its name
		OleAutomation sheetAutoBeforeUnprotect = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, sheetName);
					
		// unprotect the worksheet in order to create the annotations
		WorksheetUtils.unprotectWorksheet(sheetAutoBeforeUnprotect);
		sheetAutoBeforeUnprotect.dispose();
				
		// for each area in the range create an annotation
		for (String selectedArea : selectedAreas) {
			
			// get the OleAutomation object for the worksheet using its name
			OleAutomation sheetAutomation = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, sheetName);
			
			// create annotation object			
			String classLabel = annotationClass.getLabel();
			String annotationName = generateRangeAnnotationName(sheetName, classLabel, selectedArea);
			RangeAnnotation ra = new RangeAnnotation(sheetName, sheetIndex, annotationClass, annotationName, selectedArea);
			
			// validate annotation before creation
			boolean annotationResult =validateRangeAnnotation(workbookAutomation, sheetAutomation, ra);
			if(!annotationResult){
			    WorksheetUtils.protectWorksheet(sheetAutomation);
				WorksheetUtils.makeWorksheetActive(sheetAutomation);
				sheetAutomation.dispose();
				return;
			}
			
			// range automation has to re-created here because was disposed when checked if range is empty 
			OleAutomation rangeAutomation = WorksheetUtils.getRangeAutomation(sheetAutomation, selectedArea);
			
			// draw annotation
			drawRangeAnnotation(sheetAutomation, rangeAutomation, annotationClass, annotationName);
			
			// calculate statistics about the contents of the annotated range
			calculateStatistics(ra, workbookAutomation);
			
			// save on the AnnotationDataSheet metadata about the annotation 
			RangeAnnotationsSheet.saveRangeAnnotationData(workbookAutomation, ra);	
			
			// add the annotation object in memory data structure
			workbookAnnotation.addRangeAnnotation(ra);
			addToUndoList(ra);
		}		
		
		
		// get the OleAutomation object for the worksheet using its name
	    OleAutomation sheetAutomationAfterAnnotating = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, sheetName);
				
		// protect the worksheet to prevent user from modifying the annotations
		WorksheetUtils.protectWorksheet(sheetAutomationAfterAnnotating);
		WorksheetUtils.makeWorksheetActive(sheetAutomationAfterAnnotating);
		sheetAutomationAfterAnnotating.dispose();
	}
	
	/**
	 * Calculate stats for the annotated range
	 * @param ra a RangeAnnotation object that contains information about the annotated range
	 * @param workbookAuto an OleAutomation that provides access to the functionalities of the embedded workbook
	 */
	public static void calculateStatistics(RangeAnnotation ra, OleAutomation workbookAuto){
		
		try {
			
			OleAutomation sheetAuto = WorkbookUtils.getWorksheetAutomationByName(workbookAuto, ra.getSheetName());
			OleAutomation rangeAuto = WorksheetUtils.getRangeAutomation(sheetAuto, ra.getRangeAddress());
			OleAutomation application = WorkbookUtils.getApplicationAutomation(workbookAuto);
			
			// count all cells in the range 
			int count = RangeUtils.count(rangeAuto);
			ra.setCells(count);
			
			// count all columns in the range 
			OleAutomation columns = RangeUtils.getRangeColumns(rangeAuto);
			int countColumns = 0;
			if(columns!=null){
				countColumns= CollectionsUtils.countItemsInCollection(columns);
				columns.dispose();
			}
			ra.setColumns(countColumns);
			
			// count all rows in the range 
			OleAutomation rows = RangeUtils.getRangeRows(rangeAuto);
			int countRows = 0;
			if(rows!=null){
				countRows = CollectionsUtils.countItemsInCollection(rows);
				rows.dispose();
			}
			ra.setRows(countRows);
			
			// count all blank cells in the range 		
			int countBlank = WorksheetFunctionUtils.countBlankCells(application, rangeAuto);
			ra.setEmptyCells(countBlank);
			
			// count all formula cells in the range 
			int countFormulas = 0;
			if(RangeUtils.getMergeCells(rangeAuto)==-1){
				
				ra.setContainsMergedCells(false);
				
				OleAutomation  formulas = RangeUtils.getSpecialCells(rangeAuto, -4123); // xlCellTypeFormulas = -4123
				if(formulas!=null){			
					countFormulas = RangeUtils.count(formulas);
					formulas.dispose();
				}
				
			}else if(RangeUtils.getMergeCells(rangeAuto)==0){
				
				ra.setContainsMergedCells(true);
				
				boolean mergedFormulas = false;
				OleAutomation  formulaCells = RangeUtils.getSpecialCells(rangeAuto, -4123); // xlCellTypeFormulas = -4123
				if(formulaCells!=null){			
					countFormulas = RangeUtils.count(formulaCells);
					mergedFormulas = (RangeUtils.getMergeCells(formulaCells)>=0);
				}
				
				if(mergedFormulas){
					
					int countConstants = 0;
					boolean mergedConstants = false;
					OleAutomation  constantCells = RangeUtils.getSpecialCells(rangeAuto, 2); // xlCellTypeConstants = 2
					if(constantCells!=null){			
						countConstants = RangeUtils.count(constantCells);
						mergedConstants = (RangeUtils.getMergeCells(constantCells)>=0);
						
					}
					
					if(mergedConstants){
						
						int countBlankInFormulas = 0;
						
						OleAutomation formulaAreas = RangeUtils.getAreas(formulaCells);
						int k=1;
						while(true){
							OleAutomation area = CollectionsUtils.getItemByIndex(formulaAreas, k++, false);
							if(area==null)
								break;
							
							countBlankInFormulas += WorksheetFunctionUtils.countBlankCells(application, area);
						}
						countFormulas = countFormulas - countBlankInFormulas;
						
					}else{
						countFormulas = count-countConstants-countBlank;
					}
					
					if(constantCells!=null)
						constantCells.dispose();
				}
				
				if(formulaCells!=null)
					formulaCells.dispose();
			
			}else{
				ra.setContainsMergedCells(true);
				
				if(RangeUtils.hasFormula(rangeAuto)==0){
					countFormulas = 1;
				}
			}
			ra.setFormulaCells(countFormulas);
			
			
			// calculate constant cells in the range
			ra.setConstantCells(count - countBlank - countFormulas);
			
			rangeAuto.dispose();
			sheetAuto.dispose();
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Generate the name of the range annotation 
	 * @param sheetName the name of the worksheet that contains the annotation
	 * @param classLabel the label of the class the annotation is member of
	 * @param rangeAddress a string that represents the address of the selected range
	 * @return a string that represents the annotation name
	 */
	public static String generateRangeAnnotationName(String sheetName, String classLabel, String rangeAddress){	
		 
		 String startOfAnnotationName = getStartOfRangeAnnotationName(sheetName);
		 String endofAnnotationName = classLabel+"_"+rangeAddress.replace("$", "").replace(":", "_");
		 String formatedName = startOfAnnotationName+"_"+endofAnnotationName;
		 
		 return formatedName.toUpperCase();
	}
	
	
	/**
	 * Get the string that the names of all range annotations from the same worksheet begin with.    
	 * @param sheetName the name of the worksheet that contains the annotation
	 * @return  a string that is used as the beginning of the annotation name
	 */
	public static String getStartOfRangeAnnotationName(String sheetName){
		
		String name = sheetName.replace(" ", "_")+"_Annotation"; 
		return  name.toUpperCase(); 
	}
	
	
	/**
	 * Validate the range annotation to prevent inconsistencies. 
	 * @param annotation an object that represents a range annotation  
	 * @return true the annotation range passed all the validation tests, false otherwise
	 */
	public static boolean validateRangeAnnotation(OleAutomation  embeddedWorkbook, 
														OleAutomation sheetAutomation, RangeAnnotation annotation){
		
		// check if the range is valid (e.i., the OleAutomation can be created for this range)
		OleAutomation selectedAreaAuto = WorksheetUtils.getRangeAutomation(sheetAutomation, annotation.getRangeAddress());
		if(selectedAreaAuto==null){					
			MessageBox messageBox = Launcher.getInstance().createMessageBox(SWT.ICON_ERROR);
            messageBox.setMessage("Can not annotate range "+annotation.getRangeAddress()+". "
					+ "Please, avoid selecting entire rows or columns for annotation.");
            messageBox.open();           
			return false;
		}
		
		// ensure that the range contains data (i.e., range not empty)
		OleAutomation applicationAuto = WorkbookUtils.getApplicationAutomation(embeddedWorkbook);			
		double notEmpty = WorksheetFunctionUtils.countNotEmptyCells(applicationAuto, selectedAreaAuto);
		if(notEmpty==0){
			MessageBox messageBox = Launcher.getInstance().createMessageBox(SWT.ICON_ERROR);
            messageBox.setMessage("The selected range does not contain any value!");
            messageBox.open();
			return false;
		}
		
		// ensure that the range annotation satisfies the dependencies and containment constrains  
		AnnotationClass annotationClass =  annotation.getAnnotationClass();
		String sheetName = annotation.getSheetName();
		
		if(annotationClass.isDependent()){
						
			AnnotationClass  containerClass = annotationClass.getContainer();
			String containerLabel =  containerClass.getLabel();
			
			Collection<RangeAnnotation> collection = workbookAnnotation.getSheetAnnotationsByClass(sheetName, containerLabel);
			RangeAnnotation containerAnnotation = getSmallestParent(collection, annotation.getRangeAddress());
			
			if(containerAnnotation!=null){
				annotation.setParent(containerAnnotation);
			 
				Collection<RangeAnnotation> containersCollection = containerAnnotation.getAllAnnotations();
				if(containersCollection!=null){
					ArrayList<RangeAnnotation> dependentAnnotations = new ArrayList<RangeAnnotation>(containersCollection);
					boolean result = checkForOverlaps(dependentAnnotations, annotation.getRangeAddress(), false);
					if(result){
						 String classLabel =annotationClass.getLabel();
						 MessageBox messageBox = Launcher.getInstance().createMessageBox(SWT.ICON_ERROR);
			             messageBox.setMessage( "Could not create the \""+classLabel+"\" annotation in \""+sheetName+"\" sheet!\n"
								+ "The selected range \""+annotation.getRangeAddress()+"\" overlaps with an existing annotation.");
			             messageBox.open();
			             return false;
					}			 
					return true;
				}
				
			}else{
				String classLabel = annotationClass.getLabel();
				MessageBox messageBox = Launcher.getInstance().createMessageBox(SWT.ICON_ERROR);
	            messageBox.setMessage( "Could not create the \""+classLabel+"\" annotation in \""+sheetName+"\" sheet!\n"
						+ "The range it is not inside the borders of a "+containerLabel+" annotation range");
	            messageBox.open();
			
				return false;
			}
						
		}else{
			
			if(annotationClass.isContainable()){
				
				 Collection<RangeAnnotation> collection = workbookAnnotation.getAllRangeAnnotationsForSheet(sheetName);
				 boolean result = checkForOverlaps(collection, annotation.getRangeAddress(), true);
				 if(!result){
					 RangeAnnotation smallestParent = getSmallestParent(collection, annotation.getRangeAddress());
					 if(smallestParent!=null){
							annotation.setParent(smallestParent);
					 }
					 return true;
				 }
				 
				 String classLabel = annotationClass.getLabel();
				 MessageBox messageBox = Launcher.getInstance().createMessageBox(SWT.ICON_ERROR);
	             messageBox.setMessage( "Could not create the \""+classLabel+"\" annotation in \""+sheetName+"\" sheet! \n"
						+ "The selected range \""+annotation.getRangeAddress()+"\" overlaps with an existing annotation.");
	             messageBox.open();
		            
	             return false; 
	             
			}else{
				 Collection<RangeAnnotation> collection = workbookAnnotation.getAllRangeAnnotationsForSheet(sheetName);
				 boolean result = checkForOverlaps(collection, annotation.getRangeAddress(), false);
				 
				 if(result){
					 String classLabel = annotationClass.getLabel();
					 MessageBox messageBox = Launcher.getInstance().createMessageBox(SWT.ICON_ERROR);
		             messageBox.setMessage( "Could not create the \""+classLabel+"\" annotation in \""+sheetName+"\" sheet!\n"
							+ "The selected range \""+annotation.getRangeAddress()+"\" overlaps with an existing annotation.");
		             messageBox.open();
		             return false;
				 }				 
				 return true;
			}	
		}		
		return true;
	}
	
	
	/**
	 * Check if the new range annotation overlaps with existing range annotations
	 * @param collection the new annotation will be compared with each element of this collection for overlaps
	 * @param newAnnotationRange a string that represents the range address of the new annotation  
	 * @param ignoreContainers true to skip annotations that are members of classes that are marked as containers, false to consider them. 
	 * This argument was included especially for the cases where annotations can be contained, but do not have a specific parent annotation class.
	 * @return true if there are overlaps, false otherwise
	 */
	public static boolean checkForOverlaps(Collection<RangeAnnotation> collection, String newAnnotationRange, boolean ignoreContainers){
		
		if(collection==null){
			 return false;
		}
		
		ArrayList<RangeAnnotation> annotations = new ArrayList<RangeAnnotation>(collection);
		for (int i = 0; i < annotations.size(); i++) {
			String annotatedRange = annotations.get(i).getRangeAddress();
			boolean isPartialContainment = RangeUtils.checkForPartialContainment(annotatedRange, newAnnotationRange);
			if(isPartialContainment){
				if(ignoreContainers){
					boolean isContainer = annotations.get(i).getAnnotationClass().isContainer();
					if(isContainer)
						continue;
				}
				             
				return true; 
			}
		}
		return  false;
	}
	
	
	/**
	 * Get the smallest annotated range that contains completely the given range (annotation). 
	 * There might be several existing annotations that contain the range. These annotations 
	 * have to be members of AnnotationClasses that are defined as containers (i.e., can contain other annotations). 
	 * As annotations can not overlap with each other we are certain to find the smallest parent, which is also the direct
	 * (in hierarchy) parent of the given range (annotation).
	 * @param collection the collection of annotations to search for the smallest parent. 
	 * @param newAnnotationRange the range address of the new annotation 
	 * @return an annotation object that represents the smallest parent 
	 */
	public static RangeAnnotation getSmallestParent(Collection<RangeAnnotation> collection, String newAnnotationRange){
		
		if(collection==null){
			 return null;
		}
		
		ArrayList<RangeAnnotation> annotations = new ArrayList<RangeAnnotation>(collection);
		ArrayList<RangeAnnotation> parents = new ArrayList<RangeAnnotation>();
		for (int i = 0; i < annotations.size(); i++) {
			String annotatedRange = annotations.get(i).getRangeAddress();
			boolean isFullContainment = RangeUtils.checkForContainment(annotatedRange, newAnnotationRange);
			if(isFullContainment){
				RangeAnnotation annotation = annotations.get(i);
				boolean isContainer = annotation.getAnnotationClass().isContainer();
				if(isContainer)
					parents.add(annotation);
			}	
		}
		
		if(parents.isEmpty())
			return null;
		
		RangeAnnotation smallestParent = parents.get(0); 
		for (int j = 1; j < parents.size(); j++){
			String smallestRange = smallestParent.getRangeAddress();
			RangeAnnotation candidate = parents.get(j);
			String candidateRange = candidate.getRangeAddress(); 
			
			boolean result = RangeUtils.checkForContainment(smallestRange, candidateRange);		
			if(result)
				smallestParent = candidate;
		}
		
		return smallestParent;
	}
	
	
	
	/**
	 * Get the range of cells from the sheet that are annotated 
	 * @param workbookAutomation an OleAutomation that provides access to the functionalities of the embedded workbook
	 * @param sheetName a string the represents the name of the sheet
	 * @return an OleAutomation of a Range object that represents the areas that are annotated 
	 */
	public static OleAutomation getAnnotatedRanges(OleAutomation workbookAutomation, String sheetName){
		
		WorksheetAnnotation sheetAnnotation = workbookAnnotation.getWorksheetAnnotations().get(sheetName);
		if(sheetAnnotation == null)
			return null;
		
		Collection<RangeAnnotation> annotations = sheetAnnotation.getAllAnnotations();
		if(annotations == null  || annotations.isEmpty())
			return null;
		
		OleAutomation annotatedRanges = null;
		try{	
				
			String strConcatAddresses = null;
			for (RangeAnnotation ra : annotations) {		
				if(!ra.getAnnotationClass().isContainer()){			
					if (strConcatAddresses == null){
						strConcatAddresses = ra.getRangeAddress();
					}else{
						strConcatAddresses = strConcatAddresses.concat(",").concat(ra.getRangeAddress());
					}		
				}
			}
			
			if(!(strConcatAddresses==null || strConcatAddresses.compareTo("")==0)){
				OleAutomation sheetAuto = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, sheetName);
				annotatedRanges = WorksheetUtils.getMultiSelectionRangeAutomation(sheetAuto, strConcatAddresses);
				sheetAuto.dispose();	
			}
			
		}catch (Exception ex){
			logger.error("Genereric exception on check for unannotated ranges!", ex);
		}
		
		return annotatedRanges;
	}
	
	/**
	 * Check if the sheet has cells that are not annotated yet 
	 * @param workbookAutomation an OleAutomation that provides access to the functionalities of the embedded workbook
	 * @param sheetName a string the represents the name of the sheet
	 * @return true if the are not annotated cells, false otherwise 
	 */
	public static boolean hasUnannotatedRanges(OleAutomation workbookAutomation, String sheetName){
		OleAutomation unannotatedRanges = getUnannotatedRanges( workbookAutomation, sheetName);
		
		if(unannotatedRanges!=null)
			return true;
		
		return false;
	}
	
	/**
	 * Get the range of cells from the sheet that are not annotated yet 
	 * @param workbookAutomation an OleAutomation that provides access to the functionalities of the embedded workbook
	 * @param sheetName a string the represents the name of the sheet
	 * @return an OleAutomation of a Range object that represents the cells that are not annotated yet 
	 */
	public static OleAutomation getUnannotatedRanges(OleAutomation workbookAutomation, String sheetName){
			
		WorksheetAnnotation sheetAnnotation = workbookAnnotation.getWorksheetAnnotations().get(sheetName);
		if(sheetAnnotation == null)
			return null;
		
		if(sheetAnnotation.getAllAnnotations() == null  ||  sheetAnnotation.getAllAnnotations().isEmpty())
			return null;
		
		try{
						
			// collect all annotated ranges
			TreeMap<String, CellRangeAddress> annotatedRanges = new TreeMap<String, CellRangeAddress>();		
 			for (RangeAnnotation ra : sheetAnnotation.getAllAnnotations()) {				
				if(!ra.getAnnotationClass().isContainer()){
					CellRangeAddress cra = CellRangeAddress.valueOf(ra.getRangeAddress());					
					annotatedRanges.put("R"+cra.getFirstRow()+"C"+cra.getFirstColumn(), cra);
				}
			}
			if(annotatedRanges.isEmpty()){
				return null;
			}
			
			
			OleAutomation sheetAuto = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, sheetName);
			WorksheetUtils.unprotectWorksheet(sheetAuto);
			
			// identify the used areas (of cells) in the sheet
			OleAutomation usedRangeAuto = WorksheetUtils.getUsedRange(sheetAuto);
			sheetAuto.dispose();
					
			//identify visible cells (i.e., not hidden) in used  
			OleAutomation usedVisible = RangeUtils.getSpecialCells(usedRangeAuto, 12); // xlCellTypeVisible = 12  (visible cells)
			usedRangeAuto.dispose();	
			if(usedVisible==null){
				return null;
			}
			
			
			//identify empty cells in used  
			OleAutomation emptyUsed = RangeUtils.getSpecialCells(usedVisible, 4); // xlCellTypeEmpty = 4
			// collect all empty ranges in used
			TreeMap<String, CellRangeAddress> emptyRanges = new TreeMap<String, CellRangeAddress>();
			if(emptyUsed!=null){
				OleAutomation emptyAreas = RangeUtils.getAreas(emptyUsed);			
				int countEmpty = CollectionsUtils.countItemsInCollection(emptyAreas);
				for(int index=1; index <= countEmpty; index++){
					OleAutomation area = CollectionsUtils.getItemByIndex(emptyAreas, index, false);					 
					CellRangeAddress cra = CellRangeAddress.valueOf(RangeUtils.getRangeAddress(area));	
					emptyRanges.put("R"+cra.getFirstRow()+"C"+cra.getFirstColumn(), cra);
					area.dispose();
				}
				emptyAreas.dispose();
				emptyUsed.dispose();
			}
					
			// collect row ranges containing not annotated cells having a value (i.e., not empty)  
			TreeMap<Integer, ArrayList<CellRangeAddress>> notAnnotatedRowRanges = new TreeMap<Integer, ArrayList<CellRangeAddress>>();
			
			OleAutomation visibleUsedAreas = RangeUtils.getAreas(usedVisible);		
			int countVisible = CollectionsUtils.countItemsInCollection(visibleUsedAreas);
			
			CellRangeAddress prevCra = null;
			for(int index=1; index <= countVisible; index++){
				OleAutomation area = CollectionsUtils.getItemByIndex(visibleUsedAreas, index, false);
				CellRangeAddress visibleRng = CellRangeAddress.valueOf(RangeUtils.getRangeAddress(area));
				area.dispose();
				for (int m = visibleRng.getFirstRow(); m<visibleRng.getLastRow()+1; m++){
					for(int n = visibleRng.getFirstColumn(); n<visibleRng.getLastColumn()+1; n++){
						
						boolean inEmptyRange = false;
						CellRangeAddress toRemoveEmpty = null; 
						for(CellRangeAddress er: emptyRanges.values()){
							if(er.isInRange(m,n)){
								inEmptyRange = true;
								
								if(er.getLastRow()==m && er.getLastColumn()==n){
									toRemoveEmpty = er;
								}
								break;
							}
						}
						if(toRemoveEmpty!=null){
							//System.out.println("Remove empty region: "+toRemoveEmpty.formatAsString()+", size = "+emptyRanges.size());
							emptyRanges.remove("R"+toRemoveEmpty.getFirstRow()+"C"+toRemoveEmpty.getFirstColumn());
							toRemoveEmpty = null;
						}
						
						
						
						boolean inAnnotatedRange = false;
						CellRangeAddress toRemoveAnnotated = null; 
						for(CellRangeAddress ar: annotatedRanges.values()){
							if(ar.isInRange(m,n)){
								inAnnotatedRange = true;
								if(ar.getLastRow()==m && ar.getLastColumn()==n){
									toRemoveAnnotated = ar;
								}
								break;
							}
						}
						if(toRemoveAnnotated!=null){
							//System.out.println("Remove annotated region: "+toRemoveAnnotated.formatAsString()+", size = "+annotatedRanges.size());
							annotatedRanges.remove("R"+toRemoveAnnotated.getFirstRow()+"C"+toRemoveAnnotated.getFirstColumn());
							toRemoveAnnotated = null;
						}
											
						
						if(!inAnnotatedRange && !inEmptyRange){							
							if(prevCra==null){
								prevCra = new CellRangeAddress(m, m, n, n);
							}else{
								
								if(!(prevCra.getLastRow() == m && ((n - prevCra.getLastColumn()) == 1))){
									if(!notAnnotatedRowRanges.containsKey(prevCra.getLastRow())){
										notAnnotatedRowRanges.put(prevCra.getLastRow(), new ArrayList<CellRangeAddress>());
									}
									notAnnotatedRowRanges.get(prevCra.getLastRow()).add(prevCra);
									prevCra = new CellRangeAddress(m, m, n, n);
									
								}else{
									prevCra.setLastColumn(n);		
								}
							}	 
						}						
					}
				}			
			}
			if(prevCra!=null){
				if(!notAnnotatedRowRanges.containsKey(prevCra.getLastRow())){
					notAnnotatedRowRanges.put(prevCra.getLastRow(), new ArrayList<CellRangeAddress>());
				}
				notAnnotatedRowRanges.get(prevCra.getLastRow()).add(prevCra);
			}
			visibleUsedAreas.dispose();
			
	
			// merge the row ranges, to form larger ones (i.e., expanding multiple rows)
			ArrayList<CellRangeAddress> frontier = null;
			ArrayList<CellRangeAddress> mergedNotAnnotated = new ArrayList<CellRangeAddress>();
			int prevRow = -1;
			for(Integer rowNum: notAnnotatedRowRanges.navigableKeySet()){				
				if(frontier==null || rowNum - prevRow>1){					
					if(frontier!=null) mergedNotAnnotated.addAll(frontier);
					
					frontier = new ArrayList<CellRangeAddress>();
					for(CellRangeAddress cra: notAnnotatedRowRanges.get(rowNum))
						frontier.add(cra.copy());
					
					prevRow = rowNum;
					continue;
				}
				
				ArrayList<CellRangeAddress> currentRow = new ArrayList<CellRangeAddress>(notAnnotatedRowRanges.get(rowNum));
				for(CellRangeAddress fr: new ArrayList<CellRangeAddress>(frontier)){
					CellRangeAddress extension = null;
					for(CellRangeAddress rng: currentRow){ 
						if(fr.getFirstColumn()==rng.getFirstColumn() && fr.getLastColumn()==rng.getLastColumn()){
							extension = rng;
							break;
						}
					}
					
					if(extension==null){
						frontier.remove(fr);
						mergedNotAnnotated.add(fr);
					}else{
						fr.setLastRow(rowNum);
						currentRow.remove(extension);
					}
				}		
				
				for (CellRangeAddress rng: currentRow)
					frontier.add(rng.copy());
				prevRow = rowNum;
			}		
			if(frontier!=null)
				mergedNotAnnotated.addAll(frontier);
			
			
			OleAutomation sheetAutoToProtect = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, sheetName);
			OleAutomation applicationAuto = WorkbookUtils.getApplicationAutomation(workbookAutomation);
			OleAutomation visibleNotAnnotated = null;
			
			StringBuilder sb = new StringBuilder();
			for(int i=0; i<mergedNotAnnotated.size(); i++){			
				
				if(sb.length()>0){
					sb.append(", ");
				}				
				sb.append(mergedNotAnnotated.get(i).formatAsString());					
				
				// There seems to be a limitation with the multi selection addresses. maximum 18 supported.
				if((i+1) % 18 == 0){
					if(visibleNotAnnotated==null){
						visibleNotAnnotated = WorksheetUtils.getMultiSelectionRangeAutomation(sheetAutoToProtect, sb.toString());
					}else{
						OleAutomation multiSelectionRange = WorksheetUtils.getMultiSelectionRangeAutomation(sheetAutoToProtect, sb.toString());
						visibleNotAnnotated = ApplicationUtils.getUnion(applicationAuto, visibleNotAnnotated, multiSelectionRange);
					}		
					sb = new StringBuilder();
				}
			}
			if(sb.length()>0){
				if(visibleNotAnnotated==null){
					visibleNotAnnotated = WorksheetUtils.getMultiSelectionRangeAutomation(sheetAutoToProtect, sb.toString());
				}else{
					OleAutomation multiSelectionRange = WorksheetUtils.getMultiSelectionRangeAutomation(sheetAutoToProtect, sb.toString());
					visibleNotAnnotated = ApplicationUtils.getUnion(applicationAuto, visibleNotAnnotated, multiSelectionRange);
				}
			}
								
//			OleAutomation notEmpty = null;
//			if(visibleNotAnnotated!=null){			
//				System.out.println("Visible Not Annotated: "+RangeUtils.getRangeAddress(visibleNotAnnotated));
//				
//				// xlCellTypeConstants  = 2  (constant cells)
//				OleAutomation constants = RangeUtils.getSpecialCells(visibleNotAnnotated, 2); 
//				// xlCellTypeFormulas   = -4123  (formulas cells)
//				OleAutomation formulas =  RangeUtils.getSpecialCells(visibleNotAnnotated, -4123); 
//				
//				if(formulas==null){
//					notEmpty = constants;
//					System.out.println("Constants: "+RangeUtils.getRangeAddress(constants));
//				}else if(constants==null){
//					notEmpty = formulas;
//					System.out.println("Formulas: "+RangeUtils.getRangeAddress(formulas));
//				}else{
//					notEmpty = ApplicationUtils.getUnion(applicationAuto, constants, formulas);
//					System.out.println("Constants: "+RangeUtils.getRangeAddress(constants));
//					System.out.println("Formulas: "+RangeUtils.getRangeAddress(formulas));
//				}
//				visibleNotAnnotated.dispose();
//			}
					
			WorksheetUtils.protectWorksheet(sheetAutoToProtect);
			sheetAutoToProtect.dispose();
			
			return visibleNotAnnotated;
			
		}catch (Exception ex){
			logger.error("Genereric exception on check for unannotated ranges!", ex);
		}
		
		return null;
	}
		
	
	/**
	 * This method calls the function to draw the range annotation based on the annotation tool specified in the annotation class
	 * @param sheetAutomation an OleAutomation for accessing the active worksheet functionalities
	 * @param rangeAutomation an OleAutomation for accessing the selected range functionalities
	 * @param annotationClass the (annotation) class that will be used  for the annotation
	 * @param annotationName the name of the annotation to create
	 */
	public static void drawRangeAnnotation(OleAutomation sheetAutomation, OleAutomation rangeAutomation, 
															AnnotationClass annotationClass, String annotationName){
		switch (annotationClass.getAnnotationTool()) {
		case SHAPE  : annotateWithShape(sheetAutomation, rangeAutomation, annotationClass, annotationName); break;
		case TEXTBOX  :  annotateWithShape(sheetAutomation, rangeAutomation, annotationClass, annotationName); break;
		case BORDERAROUND: annotateByBorderAround(rangeAutomation, annotationClass, annotationName); break;
		default: logger.fatal("Option "+annotationClass.getAnnotationTool()+" not recognized.");break;
		}	
	}
	
	
	/**
	 * Draw the annotation on top of the selected range. 
	 * The annotation is drawn based on the properties specified in the annotation class 
	 * @param workbookAutomation an OleAutomation for accessing the functionalities of the embedded workbook
	 * @param ra the RangeAnnotation object to draw
	 */
	public static boolean drawRangeAnnotation(OleAutomation workbookAutomation, RangeAnnotation ra, boolean validate){
		OleAutomation sheetAutomation = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, ra.getSheetName());
		OleAutomation rangeAutomation = WorksheetUtils.getRangeAutomation(sheetAutomation, ra.getRangeAddress());
		
		if(validate){
			if(!validateRangeAnnotation(workbookAutomation, sheetAutomation, ra)){
		         return false;
			}
		}
	
		drawRangeAnnotation(sheetAutomation, rangeAutomation, ra.getAnnotationClass(), ra.getName());
		rangeAutomation.dispose();
		sheetAutomation.dispose();
		return true;
	}
	
	
	/**
	 * Draw many range annotations at once
	 * @param workbookAutomation an OleAutomation for accessing the functionalities of the embedded workbook
	 * @param rangeAnnotations the list of range annotations to draw
	 */
	public static void drawManyRangeAnnotations(OleAutomation workbookAutomation, RangeAnnotation[] rangeAnnotations, boolean validate){
		
		WorkbookUtils.unprotectAllWorksheets(workbookAutomation);			
		
		for (int i=0; i< rangeAnnotations.length; i++) {	
			try{
				AnnotationHandler.drawRangeAnnotation(workbookAutomation, rangeAnnotations[i], validate);
			}catch (Exception ex){
				logger.error("Generic exception on draw range annotation object", ex);
			}
		}
		
		WorkbookUtils.protectAllWorksheets(workbookAutomation);
	}
	
	
	/**
	 * @deprecated
	 * Draw many range annotations at once
	 * @param workbookAutomation an OleAutomation for accessing the functionalities of the embedded workbook
	 * @param annotations the list of range annotations to draw
	 * @param true to validate the annotation before drawing, false to skip validation
	 */
	public static void drawManyRangeAnnotationsOptimized(OleAutomation workbookAutomation, RangeAnnotation[] annotations, boolean validate){
		
		HashMap<String, OleAutomation> shapesAutomations = new HashMap<String, OleAutomation>();

		for (int i=0; i< annotations.length; i++) {	
						
			String currentSheetName = annotations[i].getSheetName();
			
			OleAutomation sheetAuto = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, currentSheetName);
			OleAutomation rangeAuto = WorksheetUtils.getRangeAutomation(sheetAuto, annotations[i].getRangeAddress());
			
			if(!shapesAutomations.keySet().contains(currentSheetName)){			
				OleAutomation shapesAuto = WorksheetUtils.getWorksheetShapes(sheetAuto);
				shapesAutomations.put(currentSheetName, shapesAuto);
				WorksheetUtils.unprotectWorksheet(sheetAuto);
				sheetAuto.dispose();
			}
					
			drawAnnotationShape(shapesAutomations.get(currentSheetName), rangeAuto,
					annotations[i].getAnnotationClass(), annotations[i].getName());
		}
		
		for (String sheetName : shapesAutomations.keySet()) {
			shapesAutomations.get(sheetName).dispose();
			OleAutomation sheetAutomation = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, sheetName);
			WorksheetUtils.protectWorksheet(sheetAutomation);			
		}
	}
	
	
	/**
	 * Annotate the selected range by drawing a border around it 
	 * @param rangeAutomation rangeAutomation an OleAutomation for accessing the selected range functionalities
	 * @param annotationClass the (annotation) class that will be used  for the annotation
	 * @param annotationName the name of the annotation to create
	 */
	public static void annotateByBorderAround(OleAutomation rangeAutomation, AnnotationClass annotationClass, String annotationName){
	
		long color = annotationClass.getLineColor();
		if(annotationClass.getLineColor()<0){
			color = annotationClass.getColor();
		}
			
		RangeUtils.drawBorderAroundRange(rangeAutomation, annotationClass.getLineStyle(), annotationClass.getLineWeight(), color);	
		rangeAutomation.dispose();
	}
	
	
	/**
	 * Annotate the selected range of cells (area) using a shape object
	 * @param sheetAutomation an OleAutomation for accessing the active worksheet functionalities
	 * @param rangeAutomation rangeAutomation an OleAutomation for accessing the selected range functionalities
	 * @param annotationClass the (annotation) class that will be used  for the annotation
	 * @param annotationName the name of the annotation to create 
	 */
	public static void annotateWithShape(OleAutomation sheetAutomation, OleAutomation rangeAutomation, 
																		AnnotationClass annotationClass, String annotationName){
		
		double left = RangeUtils.getRangeLeftPosition(rangeAutomation);  
		double top = RangeUtils.getRangeTopPosition(rangeAutomation);
		double width = RangeUtils.getRangeWidth(rangeAutomation);
		double height = RangeUtils.getRangeHeight(rangeAutomation);
		rangeAutomation.dispose();
		
		String currentSheetName = WorksheetUtils.getWorksheetName(sheetAutomation);
		
		OleAutomation shapesAutomation = WorksheetUtils.getWorksheetShapes(sheetAutomation);
				
		boolean result = hasShapeAnnotationsWithLabel(shapesAutomation, currentSheetName, annotationClass.getLabel());
		
		Collection<RangeAnnotation> annotationsCollection = workbookAnnotation.getSheetAnnotationsByClass(currentSheetName, annotationClass.getLabel());	
				
		if(result && annotationsCollection!=null && !annotationsCollection.isEmpty()){
			
			ArrayList<RangeAnnotation> annotations= new ArrayList<RangeAnnotation>(annotationsCollection);
			RangeAnnotation ra = annotations.get(0);				
			
			OleAutomation shapeAutomation = CollectionsUtils.getItemByName(shapesAutomation, ra.getName(), true);
			OleAutomation copyShape = ShapeUtils.duplicateShape(shapeAutomation);
			shapeAutomation.dispose();
					
			ShapeUtils.setShapeLeftPosition(copyShape, left);
			ShapeUtils.setShapeTopPosition(copyShape, top);
			ShapeUtils.setShapeHeight(copyShape, height);
			ShapeUtils.setShapeWidth(copyShape, width);
			ShapeUtils.setShapeName(copyShape, annotationName);
			copyShape.dispose();
				
		}else{
			
			if(annotationClass.getAnnotationTool()==AnnotationShape.TEXTBOX){	
				
				OleAutomation textboxAutomation = ShapeUtils.drawTextBox(shapesAutomation, left, top, width, height); 
				setAnnotationProperties(textboxAutomation, annotationClass);
				ShapeUtils.setShapeName(textboxAutomation, annotationName);
				textboxAutomation.dispose();
				
			}	
			
			if(annotationClass.getAnnotationTool()==AnnotationShape.SHAPE){		
				
				OleAutomation shapeAuto = ShapeUtils.drawShape(shapesAutomation, annotationClass.getShapeType(), left, top, width, height);
			
				setAnnotationProperties(shapeAuto, annotationClass);  
				ShapeUtils.setShapeName(shapeAuto, annotationName);
				shapeAuto.dispose();	
			}			
		}
		shapesAutomation.dispose();
	}
	
	
	/**
	 * 
	 * @param shapesAutomation
	 * @param rangeAutomation
	 * @param annotationClass
	 * @param annotationName
	 */
	public static void drawAnnotationShape(OleAutomation shapesAutomation, OleAutomation rangeAutomation, 
													AnnotationClass annotationClass, String annotationName){

		double left = RangeUtils.getRangeLeftPosition(rangeAutomation);  
		double top = RangeUtils.getRangeTopPosition(rangeAutomation);
		double width = RangeUtils.getRangeWidth(rangeAutomation);
		double height = RangeUtils.getRangeHeight(rangeAutomation);
		
		if(annotationClass.getAnnotationTool()==AnnotationShape.TEXTBOX){	
		
		OleAutomation textboxAutomation = ShapeUtils.drawTextBox(shapesAutomation, left, top, width, height); 
		setAnnotationProperties(textboxAutomation, annotationClass);
		ShapeUtils.setShapeName(textboxAutomation, annotationName);
		textboxAutomation.dispose();
		}
		
		if(annotationClass.getAnnotationTool()==AnnotationShape.SHAPE){		
		
		OleAutomation shapeAuto = ShapeUtils.drawShape(shapesAutomation, annotationClass.getShapeType(), left, top, width, height);
		setAnnotationProperties(shapeAuto, annotationClass);  
		ShapeUtils.setShapeName(shapeAuto, annotationName);
		shapeAuto.dispose();
		}
	}


	/**
	 * Check if there the sheet contains shape annotations with the given label
	 * @param shapesAutomation an OleAutomation that provides access to the functionalities of the sheet Shapes 
	 * @param sheetName a string that represents the name of the sheet to search for shape annotations
	 * @param label a string that represents the annotation_class label
	 * @return
	 */
	public static boolean hasShapeAnnotationsWithLabel(OleAutomation shapesAutomation, String sheetName, String label){
		
		boolean hasShapeAnnotations =false;
		int i=1;
		while(true){
			OleAutomation shape = CollectionsUtils.getItemByIndex(shapesAutomation, i++, true);
			
			if(shape==null){
				break;
			}
			
			String shapeName = ShapeUtils.getShapeName(shape);
			
			if(shapeName.startsWith(getStartOfRangeAnnotationName(sheetName)) && 
					shapeName.toLowerCase().indexOf("_"+label.toLowerCase()+"_")>0){
				hasShapeAnnotations = true;
				break;
			}
		}	
		return hasShapeAnnotations;
	}
	
	/**
	 * Format the annotation object (shape, textbox, etc) 
	 * @param annotation an OleAutomation to access the functionalities of the annotation object
	 * @param annotationClass an object that represents an annotation class
	 */
	public static void setAnnotationProperties(OleAutomation annotation, AnnotationClass annotationClass ){
		
		// set fill
		OleAutomation fillFormatAutomation = ShapeUtils.getFillFormatAutomation(annotation);	
		if(annotationClass.hasFill()){
			//TODO: Handle specific cases  
			ColorFormatUtils.setBackColor(fillFormatAutomation, annotationClass.getColor());
			ColorFormatUtils.setForeColor(fillFormatAutomation, annotationClass.getColor());
			FillFormatUtils.setFillTransparency(fillFormatAutomation, annotationClass.getFillTransparency());
		}else{
			FillFormatUtils.setFillVisibility(fillFormatAutomation, false);  	
		}
		
		// set border/line
		OleAutomation lineFormatAutomation = ShapeUtils.getLineFormatAutomation(annotation);
		if(annotationClass.useLine()){					
			LineFormatUtils.setLineStyle(lineFormatAutomation, annotationClass.getLineStyle());
			LineFormatUtils.setLineWeight(lineFormatAutomation, annotationClass.getLineWeight());
			ColorFormatUtils.setForeColor(lineFormatAutomation, annotationClass.getLineColor());
			LineFormatUtils.setLineTransparency(lineFormatAutomation, annotationClass.getLineTransparency());	
			
		}else{
			LineFormatUtils.setLineVisibility(lineFormatAutomation, false);
		}
		
		// set shadow
		OleAutomation shadowFormatAutomation = ShapeUtils.getShadowFormatAutomation(annotation);
		if(annotationClass.useShadow()){
			ShadowFormatUtils.setShadowType(shadowFormatAutomation, annotationClass.getShadowType()); 
			ShadowFormatUtils.setShadowStyle(shadowFormatAutomation, annotationClass.getShadowStyle());
			ShadowFormatUtils.setShadowSize(shadowFormatAutomation, annotationClass.getShadowSize());
			ShadowFormatUtils.setShadowBlur(shadowFormatAutomation, annotationClass.getShadowBlur());
			ColorFormatUtils.setForeColor(shadowFormatAutomation, annotationClass.getColor());
			ShadowFormatUtils.setShadowTransparency(shadowFormatAutomation, annotationClass.getShadowTransparency());				
		}else{
			ShadowFormatUtils.setShadowVisibility(shadowFormatAutomation, false);
		}
	
		// set text properties
		if(annotationClass.useText()){	
			
			OleAutomation textFrameAutomation = ShapeUtils.getTextFrameAutomation(annotation);
			TextFrameUtils.setHorizontalAlignment(textFrameAutomation, annotationClass.getTextHAlignment());  
			TextFrameUtils.setVerticalAlignment(textFrameAutomation, annotationClass.getTextVAlignment());
			
			OleAutomation charactersAutomation = TextFrameUtils.getCharactersAutomation(textFrameAutomation);
			CharactersUtils.setText(charactersAutomation, annotationClass.getText());
			
			OleAutomation fontAutomation = CharactersUtils.getFontAutomation(charactersAutomation);
			FontUtils.setFontColor(fontAutomation, annotationClass.getTextColor());
			FontUtils.setBoldFont(fontAutomation, annotationClass.isBoldText()); 
			FontUtils.setFontSize(fontAutomation, annotationClass.getFontSize()); 
			// TODO: font size should be relative to the size of the range ? 
			
			fontAutomation.dispose();
			charactersAutomation.dispose();
			textFrameAutomation.dispose();
		}

		shadowFormatAutomation.dispose();
		lineFormatAutomation.dispose();
		fillFormatAutomation.dispose();
	}
	
	
	/**
	 * Hide/Show all shapes in the embedded workbook that are used for annotating ranges of cells
	 * @param workbookAutomation an OleAutomation to access the functionalities of the embedded workbook
	 * @param visible true to make annotation shapes visible, false to hide them
	 */
	public static void setVisilityForAllAnnotations(OleAutomation workbookAutomation, boolean visible){
		
		OleAutomation worksheets = WorkbookUtils.getWorksheetsAutomation(workbookAutomation);
		int count = CollectionsUtils.countItemsInCollection(worksheets);
		
		// indexing starts from 1 
		for (int i = 1; i < count; i++) {
			OleAutomation sheet = CollectionsUtils.getItemByIndex(worksheets, i, false);
			setVisibilityForAnnotationsInSheet(sheet, visible);
		}		
	}
	
	
	/**
	 * Hide/Show all the shapes used for annotations in the sheet having the given name. 
	 * This method will loop through the collection of shapes in the sheet and set their visibility to false or true,
	 * if they are used for annotations. The other shapes, that were present in the original file, will not be affected. 
	 * @param workbookAutomation an OleAutomation to access the functionalities of the embedded workbook
	 * @param sheetName the name of the sheet for which the action will be performed
	 * @param visible true to make annotation shapes visible, false to hide them
	 */
	public static void setVisibilityForAnnotationsInSheet(OleAutomation workbookAutomation, String sheetName, boolean visible ){
		
		// worksheet that stores the annotation (meta-)data does not have annotation shapes 
		if(sheetName.compareTo(RangeAnnotationsSheet.name)==0)
			return;
		
		// get the OleAutomation object for the worksheet using the given name
		OleAutomation worksheetAutomation = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, sheetName);
		setVisibilityForAnnotationsInSheet(worksheetAutomation, visible);
	}
	
	
	/**
	 * Hide/Show all the shapes used for annotations in the sheet having the given name. 
	 * This method will loop through the collection of shapes in the sheet and set their visibility to false or true,
	 * if they are used for annotations. The other shapes, that were present in the original file, will not be affected.
	 * @param worksheetAutomation an OleAutoamtion to access the functionalities of the sheet that action will be applied on
	 * @param visible true to make annotation shapes visible, false to hide them
	 */
	public static void setVisibilityForAnnotationsInSheet(OleAutomation worksheetAutomation,  boolean visible){
		
		String sheetName = WorksheetUtils.getWorksheetName(worksheetAutomation);
		
		// unprotect the worksheet in order to change the visibility of the shapes
		boolean isUnprotected= WorksheetUtils.unprotectWorksheet(worksheetAutomation);
		if(!isUnprotected){
			int style = SWT.ICON_ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("ERROR: "+sheetName+" could not be unprotected!");
			message.open();
			return;
		}
		
		// get the collection of shapes in the worksheet
		OleAutomation shapesAutomation = WorksheetUtils.getWorksheetShapes(worksheetAutomation);
		
		// all shapes that are used for annotations have names that start with the following string pattern 
		String startString = getStartOfRangeAnnotationName(sheetName);
		
		int count = CollectionsUtils.countItemsInCollection(shapesAutomation);	
		for (int i = 1; i <= count; i++) {
			 OleAutomation shapeAutomation = CollectionsUtils.getItemByIndex(shapesAutomation, i, true);	 
			 String name = ShapeUtils.getShapeName(shapeAutomation);
			 if(name.indexOf(startString)== 0){
				 ShapeUtils.setShapeVisibility(shapeAutomation, visible);
			 }
		}
				
		shapesAutomation.dispose();
		
		// protect the worksheet from further user manipulation 
		WorksheetUtils.protectWorksheet(worksheetAutomation);
		worksheetAutomation.dispose();	
	}
	
	
	/**
	 * Delete all shapes in the embedded workbook that are used for annotating ranges of cells
	 * @param workbookAutomation an OleAutomation to access the functionalities of the embedded workbook
	 */
	public static void deleteAllShapeAnnotations(OleAutomation workbookAutomation){
		
		OleAutomation worksheets = WorkbookUtils.getWorksheetsAutomation(workbookAutomation);	
		int i =1; 
		while(true){
			OleAutomation sheet = CollectionsUtils.getItemByIndex(worksheets, i++, false);
			if(sheet==null)
				break;
			deleteShapeAnnotationsInSheet(sheet);
		}
	}
		
	/**
	 * Delete all the shapes used for annotations in the specified sheet 
	 * @param workbookAutomation an OleAutomation to access the functionalities of the embedded workbook
	 * @param sheetName the name of the sheet for which the annotation shapes will be deleted
	 */
	public static void deleteShapeAnnotationsInSheet(OleAutomation workbookAutomation, String sheetName ){
		
		// can not apply this method to the worksheet that stores the annotation (meta-)data, as it does not contain any shapes 
		if(sheetName.compareTo(RangeAnnotationsSheet.name)==0 || sheetName.compareTo(AnnotationStatusSheet.name)==0 || sheetName.compareTo(UsageLogSheet.name)==0)
			return;
		
		// get the OleAutomation object for the active worksheet using its name
		OleAutomation worksheetAutomation = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, sheetName);
		deleteShapeAnnotationsInSheet(worksheetAutomation);
	}
	
	/**
	 * Delete all the shapes used for annotations in the given sheet 
	 * @param worksheetAutomation an OleAutoamtion to access the functionalities of the sheet that action will be applied on
	 */
	public static void deleteShapeAnnotationsInSheet(OleAutomation worksheetAutomation){
		
		String sheetName = WorksheetUtils.getWorksheetName(worksheetAutomation);
		
		// unprotect the worksheet
		boolean isUnprotected= WorksheetUtils.unprotectWorksheet(worksheetAutomation);
		if(!isUnprotected){
			int style = SWT.ICON_ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("ERROR: "+sheetName+" could not be unprotected!");
			message.open();
			return;
		}
		
		// delete all shapes that are used for annotating ranges of cells
		OleAutomation shapesAutomation = WorksheetUtils.getWorksheetShapes(worksheetAutomation);	
	
		// all shapes that are used for annotating have names that start with the following string pattern 
		String startString =  getStartOfRangeAnnotationName(sheetName);
		
		int count = CollectionsUtils.countItemsInCollection(shapesAutomation);	
		int processed = 0; 
		int i = 1;
		while (processed!=count){
			 OleAutomation shapeAutomation = CollectionsUtils.getItemByIndex(shapesAutomation, i, true);	
			 if(shapeAutomation==null){ // it seems that it considers comments as shapes. although, it should not
				 processed++;
				 continue;
			 }
			 
			 String name = ShapeUtils.getShapeName(shapeAutomation);
			 
			 if(name.indexOf(startString)==0){
				 ShapeUtils.deleteShape(shapeAutomation);
			 }else{
				 i++;
			 }
			 
			 processed++;
		}			
		shapesAutomation.dispose();
		
		// protect the worksheet from further user manipulation 
		WorksheetUtils.protectWorksheet(worksheetAutomation);
		worksheetAutomation.dispose();	
	}
	
	/**
	 * Delete the specified range annotation from the sheet 
	 * @param worksheetAutomation an OleAutomation to access the functionalities of the sheet 
	 * @param annotation a RangeAnnotation object 
	 * @return true if deletion was successful, false otherwise
	 */
	public static boolean deleteShapeAnnotation(OleAutomation worksheetAutomation, RangeAnnotation annotation){
				
		OleAutomation shapesAutomation = WorksheetUtils.getWorksheetShapes(worksheetAutomation);
		OleAutomation shapeAutomation = CollectionsUtils.getItemByName(shapesAutomation, annotation.getName(), true);	 
		
		boolean result = false; 
		if(shapeAutomation!=null)
			result = ShapeUtils.deleteShape(shapeAutomation);
				
		shapeAutomation.dispose();
		shapesAutomation.dispose();
		
		return result;
	}
	
	
	public static void addToUndoList(RangeAnnotation annotation){
		undoList.add(annotation);
		if(undoList.size()>10)
			undoList.removeFirst();
	}
		
	public static void addToRedoList(RangeAnnotation annotation){
		redoList.add(annotation);
		if(redoList.size()>10)
			redoList.removeFirst();
	}
	
	public static void removeLastFromUndoList(){
		undoList.removeLast();
	}
	
	public static RangeAnnotation getLastFromUndoList(){
		if(!undoList.isEmpty()){
			return undoList.getLast();
		}else{
			return null;
		}
	}
	
	public static void removeLastFromRedoList(){
		redoList.removeLast();
	}
		
	public static RangeAnnotation getLastFromRedoList(){
		if(!redoList.isEmpty()){
			return redoList.getLast();
		}else{
			return null;
		}
	}

	public static void clearRedoList(){
		redoList.clear();
	}
	
	public static void clearUndoList(){
		undoList.clear();
	}
	
	
	/**
	 * @return the workbookannotation
	 */
	public static WorkbookAnnotation getWorkbookAnnotation() {
		return workbookAnnotation;
	}
	
	/**
	 * @return the oldWorkbookAnnotationHash
	 */
	public static int getOldWorkbookAnnotationHash() {
		return oldWorkbookAnnotationHash;
	}

	/**
	 * @param oldWorkbookAnnotationHash the oldWorkbookAnnotationHash to set
	 */
	public static void setOldWorkbookAnnotationHash(int oldWorkbookAnnotationHash) {
		AnnotationHandler.oldWorkbookAnnotationHash = oldWorkbookAnnotationHash;
	}
}
