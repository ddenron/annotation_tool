/**
 * 
 */
package de.tudresden.annotator.annotations;

import java.util.Collection;
import java.util.HashMap;

import de.tudresden.annotator.annotations.format.AnnotationClass;

/**
 * 
 * @author Elvis Koci
 */
public class WorkbookAnnotation extends Annotation<RangeAnnotation>{
	
	/**
	 * The name of the workbook that is embedded in the application
	 * This name can be different from the file name;
	 */
	private String workbookName;
	
	/*
	 * This two attributes mentain the status of the annotation for the workbook
	 */
	private boolean isCompleted = false;
	private boolean isNotApplicable = false;
	private NotApplicableStatus naStatus = NotApplicableStatus.NONE;
	private boolean isAmbiguous = false;
	
	/**
	 * This hashmap is used to manage worksheet annotations
	 * For the moment the key is the name of the worksheet
	 * The WorksheetAnnotation objects acts as value 
	 */
	private HashMap<String, WorksheetAnnotation> worksheetAnnotations;
	
	
	/**
	 * @param workbookName
	 */
	public WorkbookAnnotation() {
		this.worksheetAnnotations = new  HashMap<String, WorksheetAnnotation>();
	}
	
	
	/**
	 * @param workbookName
	 */
	public WorkbookAnnotation(String workbookName) {
		this.workbookName = workbookName;
		this.worksheetAnnotations = new  HashMap<String, WorksheetAnnotation>();
	}

	
	/**
	 * @param workbookName
	 * @param worksheetAnnotations
	 */
	public WorkbookAnnotation(String workbookName, HashMap<String, WorksheetAnnotation> worksheetAnnotations) {
		this.workbookName = workbookName;
		this.worksheetAnnotations = worksheetAnnotations;
	}
	
	
	/**
	 * Add a new RangeAnnotation
	 * @param sheetName the name of the worksheet where the RangeAnnotation is placed 
	 * @param sheetIndex the index of the worksheet where the RangeAnnotation is placed 
	 * @param annotationClass the AnnotationClass that this RangeAnnotation is member of
	 * @param name a string that represents the name of the RangeAnnotation
	 * @param rangeAddress the address of the range that was annotated 
	 */
	public void addRangeAnnotation(String sheetName, int sheetIndex, AnnotationClass annotationClass, String name, String rangeAddress ){
		
		RangeAnnotation rangeAnnotation= new RangeAnnotation(sheetName, sheetIndex, annotationClass, name, rangeAddress);
		addRangeAnnotation(rangeAnnotation);
	}
	
	
	/**
	 * Add a RangeAnnotation 
	 * @param rangeAnnotation an object that represents a RangeAnnotation
	 */
	public void addRangeAnnotation(RangeAnnotation rangeAnnotation){
		
		String sheetKey = rangeAnnotation.getSheetName();
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		
		if(sheetAnnotation==null){
			sheetAnnotation = new WorksheetAnnotation(rangeAnnotation.getSheetName(), rangeAnnotation.getSheetIndex());
			sheetAnnotation.setParent(this);
			this.worksheetAnnotations.put(sheetKey, sheetAnnotation);
			
		}
		
		DependentAnnotation<?> parent = rangeAnnotation.getParent();
		if(parent!=null){
			parent.addAnnotationToBucket(rangeAnnotation.getAnnotationClass().getLabel(), rangeAnnotation.getName(), rangeAnnotation);
			parent.addAnnotation(rangeAnnotation.getName(), rangeAnnotation);
		}else{
			rangeAnnotation.setParent(sheetAnnotation);
		}
		
		sheetAnnotation.addAnnotation(rangeAnnotation.getName(), rangeAnnotation);
		sheetAnnotation.addAnnotationToBucket(rangeAnnotation.getAnnotationClass().getLabel(), rangeAnnotation.getName(), rangeAnnotation);
		
		this.addAnnotation(rangeAnnotation.getName(), rangeAnnotation);
		this.addAnnotationToBucket(rangeAnnotation.getAnnotationClass().getLabel(), rangeAnnotation.getName(), rangeAnnotation);
	}
	
	
	/**
	 * Get the RangeAnnotation based on the worksheet key and annotation key
	 * @param sheetKey a string that represents the id (key) of the worksheet where the RangeAnnotation is placed 
	 * @param annotationKey a string that is used as key for the annotation object 
	 * @return the RangeAnnotation object that corresponds to the given arguments  
	 */
	public RangeAnnotation getRangeAnnotation(String sheetKey, String annotationKey){	
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		
		if(sheetAnnotation==null)
			return null;
		
		return sheetAnnotation.getAnnotation(annotationKey);
	}
	
	
	/**
	 * Get the collection of RangeAnnotations for the given Worksheet key and AnnotationClass label
	 * @param sheetKey a string that represents the id (key) of the worksheet where the RangeAnnotation is placed
	 * @param classLabel the label of the AnnotationClass that this RangeAnnotation is member of
	 * @return a collection of RangeAnnotations that correspond to the given arguments
	 */
	public Collection<RangeAnnotation> getSheetAnnotationsByClass(String sheetKey, String classLabel){
		
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		
		if(sheetAnnotation==null)
			return null;
		
		return sheetAnnotation.getAnnotationsByClass(classLabel);		
	}
	
	
	/**
	 * Get all annotation objects for the specified worksheet
	 * @param sheetKey a string that represents the key (id) of the worksheet 
	 * @return a collection of annotations objects
	 */
	public Collection<RangeAnnotation> getAllRangeAnnotationsForSheet(String sheetKey){
		
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		
		if(sheetAnnotation==null)
			return null;
		
		return sheetAnnotation.getAllAnnotations();
	}
	
	
	/**
	 * Remove a RangeAnnotation 
	 * @param rangeAnnotation an object that represents a RangeAnnotation
	 */
	public void removeRangeAnnotation(RangeAnnotation rangeAnnotation){
		String sheetKey = rangeAnnotation.getSheetName(); 
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		
		if(sheetAnnotation==null)
			return;
		
		String classLabel = rangeAnnotation.getAnnotationClass().getLabel();
		
		sheetAnnotation.removeAnnotation(rangeAnnotation.getName());
		sheetAnnotation.removeAnnotationFromBucket(classLabel, rangeAnnotation.getName());
		
		if(rangeAnnotation.getParent() instanceof RangeAnnotation){
			rangeAnnotation.getParent().removeAnnotation(rangeAnnotation.getName());
			rangeAnnotation.getParent().removeAnnotationFromBucket(classLabel, rangeAnnotation.getName());
		}
		
		this.removeAnnotation(rangeAnnotation.getName());
		this.removeAnnotationFromBucket(classLabel, rangeAnnotation.getName());
	}
	
	
	/**
	 * Remove a RangeAnnotation
	 * @param sheetKey a string that represents the id (key) of the worksheet where the RangeAnnotation is placed
	 * @param classLabel the label of the AnnotationClass that this RangeAnnotation is member of
	 * @param annotationKey a string that is used as key for the annotation object 
	 */
	public void removeRangeAnnotation(String sheetKey, String classLabel, String annotationKey){
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		
		if(sheetAnnotation==null)
			return;
		
		DependentAnnotation<?> parent = sheetAnnotation.getAnnotation(annotationKey).getParent();
		
		if(parent instanceof RangeAnnotation){	
			parent = (RangeAnnotation) parent;
			parent.removeAnnotation(annotationKey);
			parent.removeAnnotationFromBucket(classLabel, annotationKey);
		}
		
		sheetAnnotation.removeAnnotation(annotationKey);
		sheetAnnotation.removeAnnotationFromBucket(classLabel, annotationKey);
		
		this.removeAnnotation(annotationKey);
		this.removeAnnotationFromBucket(classLabel, annotationKey);
	}
	
	
	/**
	 * Remove a RangeAnnotation
	 * @param sheetKey a string that represents the id (key) of the worksheet where the RangeAnnotation is placed
	 * @param rangeAnnotationKey a string that is used as key for the annotation object 
	 */
	public void removeRangeAnnotation(String sheetKey, String rangeAnnotationKey){
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		
		if(sheetAnnotation==null)
			return;
		
		String classLabel = sheetAnnotation.getAnnotation(rangeAnnotationKey).getAnnotationClass().getLabel();
		DependentAnnotation<?> parent = sheetAnnotation.getAnnotation(rangeAnnotationKey).getParent();
		
		if(parent instanceof RangeAnnotation){	
			RangeAnnotation parentAnnotation = (RangeAnnotation) parent;
			parentAnnotation.removeAnnotation(rangeAnnotationKey);
			parentAnnotation.removeAnnotationFromBucket(classLabel, rangeAnnotationKey);
		}
		
		sheetAnnotation.removeAnnotation(rangeAnnotationKey);
		sheetAnnotation.removeAnnotationFromBucket(classLabel, rangeAnnotationKey);
		
		this.removeAnnotation(rangeAnnotationKey);
		this.removeAnnotationFromBucket(classLabel, rangeAnnotationKey);
	}
	
	
	/**
	 * Remove all RangeAnnotations belonging to the specified Workbook and AnnotationClass
	 * @param sheetKey a string that represents the id (key) of the worksheet where the RangeAnnotation is placed
	 * @param classLabel the label of the AnnotationClass that this RangeAnnotation is member of
	 * 
	 */
	public void emptySheetAnnotationBucket(String sheetKey, String classLabel){
		
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		
		if(sheetAnnotation==null)
			return;
		
		RangeAnnotation[] rangeAnnotations = 
				sheetAnnotation.getAnnotationsByClass(classLabel).toArray(
						new RangeAnnotation[sheetAnnotation.getAllAnnotations().size()]);
			
		for (RangeAnnotation rangeAnnotation : rangeAnnotations) {
			this.removeRangeAnnotation(rangeAnnotation);
		}
		
		sheetAnnotation.removeAllAnnotationsOfClass(classLabel);
	}
	
	
	/**
	 * Remove all RangeAnnotations belonging to the worksheet with the given key
	 * @param sheetKey a string that represents the id (key) of the worksheet where the RangeAnnotation is placed
	 */
	public void removeAllRangeAnnotationsFromSheet(String sheetKey){
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		RangeAnnotation[] rangeAnnotations = 
				sheetAnnotation.getAllAnnotations().toArray(
						new RangeAnnotation[sheetAnnotation.getAllAnnotations().size()]);
		
		for (RangeAnnotation rangeAnnotation : rangeAnnotations) {
			this.removeRangeAnnotation(rangeAnnotation);
		}
	}
	
	/**
	 * Remove all RangeAnnotations belonging to the worksheet with the given key
	 * @param sheetAnnotation the WorksheetAnnotation object that contains all the range annotations to delete
	 */
	public void removeAllRangeAnnotationsFromSheet(WorksheetAnnotation sheetAnnotation){
		
		RangeAnnotation[] rangeAnnotations = 
				sheetAnnotation.getAllAnnotations().toArray(
						new RangeAnnotation[sheetAnnotation.getAllAnnotations().size()]);
		
		for (RangeAnnotation rangeAnnotation : rangeAnnotations) {
			this.removeRangeAnnotation(rangeAnnotation);
		}
	}
	
	@Override
	/**
	 * Remove all annotations 
	 */
	public void removeAllAnnotations(){
		this.allAnnotations.clear();
		this.annotationsByClass.clear();
		this.worksheetAnnotations.clear();
		this.isCompleted = false;
		this.isNotApplicable = false;
		this.naStatus = NotApplicableStatus.NONE;
		this.isAmbiguous = false;
	}
	
	
	/**
	 * @return the workbookName
	 */
	public String getWorkbookName() {
		return workbookName;
	}

	/**
	 * @param workbookName the workbookName to set
	 */
	public void setWorkbookName(String workbookName) {
		this.workbookName = workbookName;
	}
	
		
	/**
	 * @return the worksheetAnnotations
	 */
	public HashMap<String, WorksheetAnnotation> getWorksheetAnnotations() {
		return worksheetAnnotations;
	}
	

	/**
	 * @return the isCompleted
	 */
	public boolean isCompleted() {
		return isCompleted;
	}


	/**
	 * @param isCompleted the isCompleted to set
	 */
	public void setCompleted(boolean isCompleted) {
		this.isCompleted = isCompleted;
		if(isCompleted){
			this.isNotApplicable = false;
			this.naStatus = NotApplicableStatus.NONE;
			this.isAmbiguous = false;
		}
	}


	/**
	 * @return the isIrrelevant
	 */
	public boolean isNotApplicable() {
		return isNotApplicable;
	}


	/**
	 * @param isNotApplicable the isNotApplicable to set
	 */
	public void setNotApplicable(boolean isNotApplicable) {
		this.isNotApplicable = isNotApplicable;
		if(isNotApplicable){
			this.isCompleted = false;
			this.isAmbiguous = false;
		}else{
			this.naStatus = NotApplicableStatus.NONE;
		}
	}

	
	/**
	 * @return the isAmbiguous
	 */
	public boolean isAmbiguous() {
		return isAmbiguous;
	}


	/**
	 * @param isAmbiguous the isAmbiguous to set
	 */
	public void setAmbiguous(boolean isAmbiguous) {
		this.isAmbiguous = isAmbiguous;
		if(isAmbiguous){
			this.isCompleted = false;
			this.isNotApplicable = false;
			this.naStatus = NotApplicableStatus.NONE;
		}
	}

	
	/**
	 * 
	 * @return
	 */
	public boolean hasActiveStatus(){
		return this.isCompleted || this.isAmbiguous || this.isNotApplicable;
	}
	
	/**
	 * 
	 * @return
	 */
	public boolean hasActiveWorksheetStatus(){
		for(WorksheetAnnotation sha : worksheetAnnotations.values()){
			if(sha.hasActiveStatus()){
				return true;
			}
		}
		return false;
	}

	/**
	 * @return the naStatus
	 */
	public NotApplicableStatus getNotApplicableStatus() {
		return naStatus;
	}


	/**
	 * @param naStatus the naStatus to set
	 */
	public void setNaStatus(NotApplicableStatus naStatus) {
		this.naStatus = naStatus;
		if (naStatus!=NotApplicableStatus.NONE){
			setNotApplicable(true);
		}else{
			this.isNotApplicable = false;
		}
	}


	@Override
	public String toString() {
		return this.worksheetAnnotations.values().toString();
	}
	
	@Override
	public int hashCode() {	
		int hash = this.getWorkbookName().hashCode() + (this.isCompleted?1:0) + 
				(this.isNotApplicable?1:0) + (this.naStatus.getStatusId()) + (this.isAmbiguous?1:0);
		
		for (WorksheetAnnotation val : this.getWorksheetAnnotations().values()) {
			hash = hash + val.hashCode();
		}		
		return hash;
	}

	@Override
	public boolean equals(Annotation<RangeAnnotation> annotation) {
		
		if(!(annotation instanceof WorkbookAnnotation))
			return false;
		
		WorkbookAnnotation wa = (WorkbookAnnotation) annotation;
				
		if(this.workbookName.compareTo(wa.getWorkbookName())!=0)
			return false;
		
		if(!this.worksheetAnnotations.equals(wa.getWorksheetAnnotations()))
			return false;
		
		if(wa.isCompleted()!=this.isCompleted)
			return false;
			
		if(wa.isNotApplicable()!=this.isNotApplicable)
			return false;
		
		if(wa.getNotApplicableStatus()!=this.naStatus)
			return false;
		
		if(wa.isAmbiguous()!=this.isAmbiguous)
			return false;
		
		return true;	
	}
	
}
