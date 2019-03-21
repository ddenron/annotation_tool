/**
 * 
 */
package de.tudresden.annotator.annotations;

import de.tudresden.annotator.annotations.format.AnnotationClass;

/**
 * @author Elvis Koci
 */
public class RangeAnnotation extends DependentAnnotation<DependentAnnotation<?>> {

	private String sheetName;
	private int sheetIndex;
	private AnnotationClass annotationClass; 	
	private String name;
	private String rangeAddress;
	
	private int cells;
	private int emptyCells;
	private int constantCells;
	private int formulaCells;
	private int rows;
	private int nonEmptyRows;
	private int columns; 
	private int nonEmptyColumns;
	
	private boolean containsMergedCells;
	
	/**
	 * Create a new RangeAnnotation
	 * @param sheetName the name of the sheet where the RangeAnnotation is placed 
	 * @param sheetIndex the index of the sheet where the RangeAnnotation is placed 
	 * @param annotationClass the AnnotationClass that this RangeAnnotation is member of
	 * @param name a string that represents the name of the RangeAnnotation
	 * @param rangeAddress the address of the range that was annotated 
	 */
	public RangeAnnotation(String sheetName, int sheetIndex, AnnotationClass annotationClass, String name, String rangeAddress ) {
		this.annotationClass = annotationClass;
		this.name = name;
		this.rangeAddress = rangeAddress;
		this.sheetName = sheetName;
		this.sheetIndex = sheetIndex;	
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
	 * @return the annotationClass
	 */
	public AnnotationClass getAnnotationClass() {
		return annotationClass;
	}

	/**
	 * @param annotationClass the annotationClass to set
	 */
	public void setAnnotationClass(AnnotationClass annotationClass) {
		this.annotationClass = annotationClass;
	}

	/**
	 * @return the name
	 */
	public String getName() {
		return name;
	}

	/**
	 * @param name the name to set
	 */
	public void setName(String name) {
		this.name = name;
	}

	/**
	 * @return the rangeAddress
	 */
	public String getRangeAddress() {
		return rangeAddress;
	}

	/**
	 * @param rangeAddress the rangeAddress to set
	 */
	public void setRangeAddress(String rangeAddress) {
		this.rangeAddress = rangeAddress;
	}

	/**
	 * @return the cells
	 */
	public int getCells() {
		return cells;
	}

	/**
	 * @param cells the cells to set
	 */
	public void setCells(int cells) {
		this.cells = cells;
	}

	/**
	 * @return the emptyCells
	 */
	public int getEmptyCells() {
		return emptyCells;
	}

	/**
	 * @param emptyCells the emptyCells to set
	 */
	public void setEmptyCells(int emptyCells) {
		this.emptyCells = emptyCells;
	}

	/**
	 * @return the constantCells
	 */
	public int getConstantCells() {
		return constantCells;
	}

	/**
	 * @param constantCells the constantCells to set
	 */
	public void setConstantCells(int constantCells) {
		this.constantCells = constantCells;
	}

	/**
	 * @return the formulaCells
	 */
	public int getFormulaCells() {
		return formulaCells;
	}

	/**
	 * @param formulaCells the formulaCells to set
	 */
	public void setFormulaCells(int formulaCells) {
		this.formulaCells = formulaCells;
	}

	/**
	 * @return the rows
	 */
	public int getRows() {
		return rows;
	}

	/**
	 * @param rows the rows to set
	 */
	public void setRows(int rows) {
		this.rows = rows;
	}

	/**
	 * @return the nonEmptyRows
	 */
	public int getNonEmptyRows() {
		return nonEmptyRows;
	}

	/**
	 * @param nonEmptyRows the nonEmptyRows to set
	 */
	public void setNonEmptyRows(int nonEmptyRows) {
		this.nonEmptyRows = nonEmptyRows;
	}

	/**
	 * @return the columns
	 */
	public int getColumns() {
		return columns;
	}

	/**
	 * @param columns the columns to set
	 */
	public void setColumns(int columns) {
		this.columns = columns;
	}

	/**
	 * @return the nonEmptyColumns
	 */
	public int getNonEmptyColumns() {
		return nonEmptyColumns;
	}

	/**
	 * @param nonEmptyColumns the nonEmptyColumns to set
	 */
	public void setNonEmptyColumns(int nonEmptyColumns) {
		this.nonEmptyColumns = nonEmptyColumns;
	}

	/**
	 * @return the hasMergedCells
	 */
	public boolean containsMergedCells() {
		return containsMergedCells;
	}

	/**
	 * @param hasMergedCells the hasMergedCells to set
	 */
	public void setContainsMergedCells(boolean hasMergedCells) {
		this.containsMergedCells = hasMergedCells;
	}

	@Override 
	public String toString() {
			return this.name+" = "+this.allAnnotations.values().toString();
	}

	@Override
	public boolean equals(Annotation<RangeAnnotation> annotation) {
		
		if(!(annotation instanceof RangeAnnotation))
			return false;
		
		RangeAnnotation ra = (RangeAnnotation) annotation;
				
		if(this.name.compareTo(ra.getName())!=0)
			return false;
		
		if(!(this.allAnnotations.equals(ra.getAllAnnotations())))
			return false;
		
		return true;	
	}

	
	@Override
	public int hashCode() {
		
		int hash = this.name.hashCode();
		
		for (RangeAnnotation val : this.allAnnotations.values()) {
			hash = hash + val.hashCode();
		}
		
		return hash;
	}
}
