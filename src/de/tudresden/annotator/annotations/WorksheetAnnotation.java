/**
 * 
 */
package de.tudresden.annotator.annotations;

/**
 * @author Elvis Koci
 */
public class WorksheetAnnotation extends DependentAnnotation<WorkbookAnnotation> {

	private String workbookName;
	private String sheetName;
	private int sheetIndex;
	private boolean isCompleted = false;
	private boolean isNotApplicable = false;
	private NotApplicableStatus naStatus = NotApplicableStatus.NONE;
	private boolean isAmbiguous = false;

	/**
	 * @param workbookName
	 * @param sheetName
	 * @param sheetIndex
	 */
	public WorksheetAnnotation(String workbookName, String sheetName, int sheetIndex) {
		this.workbookName = workbookName;
		this.sheetIndex = sheetIndex;
		this.sheetName = sheetName;
	}
		
	/**
	 * @param sheetName
	 * @param sheetIndex
	 */
	public WorksheetAnnotation(String sheetName, int sheetIndex) {
		this.sheetIndex = sheetIndex;
		this.sheetName = sheetName;
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
	 * @return the sheetName
	 */
	public String getSheetName() {
		return sheetName;
	}

	/**
	 * @param sheetName the sheetName to set
	 */
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	/**
	 * @return the sheetIndex
	 */
	public int getSheetIndex() {
		return sheetIndex;
	}

	/**
	 * @param sheetIndex the sheetIndex to set
	 */
	public void setSheetIndex(int sheetIndex) {
		this.sheetIndex = sheetIndex;
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
	 * @return the NotApplicable
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
		if(naStatus!=NotApplicableStatus.NONE){
			setNotApplicable(true);
		}else{
			this.isNotApplicable = true;
		}
		
	}
	
	@Override 
	public String toString() {
		return this.getSheetName()+" = "+this.allAnnotations.values(); 
	}

	@Override
	public boolean equals(Annotation<RangeAnnotation> annotation) {
		
		if(!(annotation instanceof WorksheetAnnotation))
			return false;
		
		WorksheetAnnotation sa = (WorksheetAnnotation) annotation;
		
		if(sa.getSheetName().compareTo(this.sheetName)!=0)
			return false;
		
		if(!(sa.getAllAnnotations().equals(this.allAnnotations)))
			return false;
		
		if(sa.isCompleted()!=this.isCompleted)
			return false;
			
		if(sa.isNotApplicable()!=this.isNotApplicable)
			return false;
		
		if(sa.getNotApplicableStatus()!=this.naStatus)
			return false;
		
		if(sa.isAmbiguous()!=this.isAmbiguous)
			return false;
		
		return true;
	}

	@Override
	public int hashCode() {
		int hash = this.getSheetName().hashCode() + (this.isCompleted?1:0) + 
				(this.isNotApplicable?1:0) + (this.naStatus.getStatusId()) + (this.isAmbiguous?1:0);
		
		for (RangeAnnotation val : this.allAnnotations.values()) {
			hash = hash + val.hashCode();
		}
		
		return hash;
	}
}
