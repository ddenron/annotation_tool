/**
 * 
 */
package de.tudresden.annotator.oleutils;

import java.util.Arrays;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class RangeUtils {
	
	
	/**
	 * Get the address of the range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return a string that represents the address of the given range. 
	 */
	public static String getRangeAddress(OleAutomation rangeAutomation){
		
		int[] addressIds = rangeAutomation.getIDsOfNames(new String[]{"Address"}); 
		Variant addressVariant = rangeAutomation.getProperty(addressIds[0]);	
		String address = addressVariant.getString();
		addressVariant.dispose();
		
		return address;
	}
	
	
	/**
	 * Get the number of cells in the range. This method can handle very large selection of cells 
	 * @param rangeAutomation an automation that provides access to Range OLE object
	 * @return a long that represents the number of cells in the range
	 */
	public static long countLarge(OleAutomation rangeAutomation){
		
		int[] countLargeProperyIds = rangeAutomation.getIDsOfNames(new String[]{"CountLarge"});
				
		Variant countLargePropertyVariant =  rangeAutomation.getProperty(countLargeProperyIds[0]);			
				
		long countLarge = countLargePropertyVariant.getLong();
		countLargePropertyVariant.dispose();
		
		return countLarge;
	}
	
	
	/**
	 * Get the number of cells in the range. This method can handle ranges having up to 2,147,483,647 cells
	 * @param rangeAutomation an automation that provides access to Range OLE object
	 * @return an integer that represents the number of cells in the range
	 */
	public static int count(OleAutomation rangeAutomation){
		
		int[] countProperyIds = rangeAutomation.getIDsOfNames(new String[]{"Count"});
				
		Variant countPropertyVariant =  rangeAutomation.getProperty(countProperyIds[0]);			
				
		int count = countPropertyVariant.getInt();
		countPropertyVariant.dispose();
		
		return count;
	}
	
	
	/**
	 * Get areas from a range that represents a multi area selection  
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return an OleAutomation that provides access to the collection of areas. 
	 */
	public static OleAutomation getAreas(OleAutomation rangeAutomation){
		
		int[] areasPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Areas"}); 
		Variant areasVariant = rangeAutomation.getProperty(areasPropertyIds[0]);	
		OleAutomation areasAutomation = areasVariant.getAutomation();
		areasVariant.dispose();
		
		return areasAutomation;
	}
	
	
	/**
	 * Get a collection of cells from a range  
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return an OleAutomation that provides access to the collection of cells. 
	 */
	public static OleAutomation getCells(OleAutomation rangeAutomation){
		
		int[] cellsPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Cells"}); 
		Variant cellsVariant = rangeAutomation.getProperty(cellsPropertyIds[0]);	
		OleAutomation cellsAutomation = cellsVariant.getAutomation();
		cellsVariant.dispose();
		
		return cellsAutomation;
	}
	
	
	/**
	 * Get range value. This method assumes single cell ranges.
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return a string that represents the value of the range
	 */
	public static String getValue(OleAutomation rangeAutomation){
		int[] valuePropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Value"});
		
		Variant valueVariant =  rangeAutomation.getProperty(valuePropertyIds[0]);
			
		String value = "";
		if(valueVariant!=null && valueVariant.getType() != 0){
			value= valueVariant.getString();
			valueVariant.dispose();
		}		
		
		return value;
	}
	
	
	/**
	 * Set a value to the range. This method assumes single cell ranges.
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @param value the string to set as value 
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean setValue(OleAutomation rangeAutomation, String value){
		
		int[] valuePropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Value"});
		
		Variant valueVariant = new Variant(value);
		boolean isSuccess = rangeAutomation.setProperty(valuePropertyIds[0], valueVariant);
		valueVariant.dispose();
		
		return isSuccess;
	}
	
	
	/**
	 * Get the value of each cell in the range. 
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return an array of strings. each string represent a cell value.
	 */
	public static String[] getRangeValues(OleAutomation rangeAutomation){
		
		OleAutomation cells = RangeUtils.getCells(rangeAutomation);
		int countCells = CollectionsUtils.countItemsInCollection(cells); 
		
		String[] values = new String[countCells];
		int i = 1; 
		while(i<=countCells){
			OleAutomation cellAutomation = CollectionsUtils.getItemByIndex(cells, i, false);
			values[i-1]= RangeUtils.getValue(cellAutomation);
			i++;
		}
		
		return values;
	}
	
	
	/**
	 * Get the number of the first column in the first area in the specified range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return the number of the first column in the first area in the specified range
	 */
	public static int getFirstColumnIndex(OleAutomation rangeAutomation){
		
		int[] columnPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Column"}); 
		Variant columnPropertyVariant = rangeAutomation.getProperty(columnPropertyIds[0]);	
		int column = columnPropertyVariant.getInt();
		columnPropertyVariant.dispose();
		
		return column;
	}
	
	
	/**
	 * Get collection of columns in the range 
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return an OleAutomation that provides access to the collection of columns in the range 
	 */
	public static OleAutomation getRangeColumns(OleAutomation rangeAutomation){
				
		int[] columnsPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Columns"}); 
		Variant columnsPropertyVariant = rangeAutomation.getProperty(columnsPropertyIds[0]);	
		
		OleAutomation columnsAutomation =  columnsPropertyVariant.getAutomation();
		columnsPropertyVariant.dispose();
		
		return columnsAutomation;
	}
	
	/**
	 * Get a specific column in the range 
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return an OleAutomation that provides access to the range that represents the entire specified column 
	 */
	public static OleAutomation getRangeColumn(OleAutomation rangeAutomation, String column){
				
		int[] columnsPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Columns"}); 
		
		Variant[] args= new Variant[1];
		args[0] = new Variant(column);
		
		Variant columnPropertyVariant = rangeAutomation.getProperty(columnsPropertyIds[0], args);	
		args[0].dispose();
		
		OleAutomation columnAutomation =  columnPropertyVariant.getAutomation();
		columnPropertyVariant.dispose();
		
		return columnAutomation;
	}
	
	
	/**
	 * Get the number of the first row in the first area in the specified range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return the number of the first row in the first area in the specified range
	 */
	public static int getFirstRowIndex(OleAutomation rangeAutomation){
		
		int[] rowPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Row"}); 
		Variant rowPropertyVariant = rangeAutomation.getProperty(rowPropertyIds[0]);	
		int row = rowPropertyVariant.getInt();
		rowPropertyVariant.dispose();
		
		return row;
	}
	
	/**
	 * Get collection of rows in the range 
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return an OleAutomation that provides access to the collection of rows in the range 
	 */
	public static OleAutomation getRangeRows(OleAutomation rangeAutomation){
		
		int[] rowsPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Rows"}); 
		Variant rowsPropertyVariant = rangeAutomation.getProperty(rowsPropertyIds[0]);	
		OleAutomation rowsAutomation = rowsPropertyVariant.getAutomation();
		rowsPropertyVariant.dispose();
		
		return rowsAutomation;
	}
	
	
	/**
	 * Get a specific row from the range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return an OleAutomation that provides access to the range that represents the entire specified row 
	 */
	public static OleAutomation getRangeRow(OleAutomation rangeAutomation, int row){
		
		int[] rowsPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Rows"}); 

		Variant[] args= new Variant[1];
		args[0] = new Variant(row);
		
		Variant rowPropertyVariant = rangeAutomation.getProperty(rowsPropertyIds[0], args);
		args[0].dispose();
		
		OleAutomation rowAutomation = rowPropertyVariant.getAutomation();
		rowPropertyVariant.dispose();
		
		return rowAutomation;
	}
	
	
	/**
	 * Hide/Show the given range. This method applies only to ranges that represent entire rows or columns
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 */
	public static void setRangeVisibility(OleAutomation rangeAutomation, boolean visible){
		
		int[] hiddenPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Hidden"});
		Variant hiddenVariant = new Variant(!visible);

		rangeAutomation.setProperty(hiddenPropertyIds[0], hiddenVariant);
		hiddenVariant.dispose();
	}
	
	
	/**
	 * Get the distance, in points, from the left edge of column A to the left edge of the range.
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return the distance, in points, from the left edge of column A to the left edge of the range
	 */
	public static double getRangeLeftPosition(OleAutomation rangeAutomation){

		int[] leftPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Left"});
		Variant leftVariant=rangeAutomation.getProperty(leftPropertyIds[0]);
		double left = leftVariant.getDouble();
		leftVariant.dispose();
		return left;
	}
	
	
	/**
	 * Get the distance, in points, from the top edge of row 1 to the top edge of the range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return the distance, in points, from the top edge of row 1 to the top edge of the range
	 */
	public static double getRangeTopPosition(OleAutomation rangeAutomation){
		
		int[] topPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Top"});
		Variant topVariant=rangeAutomation.getProperty(topPropertyIds[0]);
		double top = topVariant.getDouble();
		topVariant.dispose();
		
		return top;
	}
	
	
	/**
	 * Get the height, in units, of the range.
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return the height, in units, of the range.
	 */
	public static double getRangeHeight(OleAutomation rangeAutomation){
		
		int[] heightPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Height"});
		Variant heightVariant=rangeAutomation.getProperty(heightPropertyIds[0]);
		double height = heightVariant.getDouble();
		heightVariant.dispose();
		
		return height;
	}
	
	
	/**
	 * Get the width, in units, of the range.
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return the width, in units, of the range.
	 */
	public static double getRangeWidth(OleAutomation rangeAutomation){
		
		int[] widthPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Width"});
		
		Variant widthVariant=rangeAutomation.getProperty(widthPropertyIds[0]);
		double width = widthVariant.getDouble();
		widthVariant.dispose();
		
		return width;
	}
	
	
	/**
	 * Select the given range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 */
	public static boolean selectRange(OleAutomation rangeAutomation){
		
		int[] selectMethodIds = rangeAutomation.getIDsOfNames(new String[]{"Select"});
		Variant result = rangeAutomation.invoke(selectMethodIds[0]);
		
		if(result==null)
			return false;
		
		boolean isSuccess = result.getBoolean();
		result.dispose();
		
		return isSuccess;
	}
	
	
	/**
	 * Filter range based on a criteria applied to a single field  
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return true if the operation succeeded, false otherwise
	 */
	public static boolean filterRange(OleAutomation rangeAutomation, int field, String criteria1){
		
		int[] autoFilterMethodIds = rangeAutomation.getIDsOfNames(new String[]{"AutoFilter", "Field", "Criteria1"});
		
		Variant[] args = new Variant[2];
		args[0] = new Variant(field);
		args[1] = new Variant(criteria1);
		int argsIds[] = Arrays.copyOfRange(autoFilterMethodIds, 1, autoFilterMethodIds.length);
		
		Variant result = rangeAutomation.invoke(autoFilterMethodIds[0], args, argsIds);
		for (Variant arg : args) {
			arg.dispose();
		}
		
		if(result == null){
			return false;
		}
		
		result.dispose();
		return true;
	}
	
	
	/**
	 * Format the cells in the given range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @param format a string that represents the format that the cells should have 
	 * @return true if the operation succeeded, false otherwise
	 */
	public static boolean formatCells(OleAutomation rangeAutomation, String format){
		
		int[] numberFormatMethodIds = rangeAutomation.getIDsOfNames(new String[]{"NumberFormat"});
		
		Variant formatVariant = new Variant(format);
		boolean isSuccess = rangeAutomation.setProperty(numberFormatMethodIds[0], formatVariant);
		formatVariant.dispose();
		
		return isSuccess;
	}
	
	/**
	 * Check if the given range contains merged cells 
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return 1 if all cells in the range are merged, 0 if the range contains merged cells, 
	 * -1 if the range does not contain any merged cells
	 */
	public static int getMergeCells(OleAutomation rangeAutomation){
		
		int[] mergeCellsPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"MergeCells"});
		
		Variant result = rangeAutomation.getProperty(mergeCellsPropertyIds[0]);
				
		if(result==null)
			return -1;
		
		if(result.getType()==1){
			result.dispose();
			return 0;
		}
		
		boolean hasMergedCells = result.getBoolean();
		result.dispose();
		
		return hasMergedCells? 1 : -1;
	}
	
	/**
	 * Check if the given range is part of a merged area 
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return an OleAutomation of a Range object that represents a merged area
	 */
	public static OleAutomation getMergeArea(OleAutomation rangeAutomation){
		
		int[] mergeAreaPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"MergeArea"});
		
		Variant result = rangeAutomation.getProperty(mergeAreaPropertyIds[0]);
		
		if(result==null)
			return null;
		
		OleAutomation mergeArea = result.getAutomation();
		result.dispose();
		
		return mergeArea;
	}
	
	
	/**
	 * Check if the given range contains formulas
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 *  @return 1 if all cells in the range have formulas, 0 if the range contains at least one 
	 *  cell that has formula, -1 if the range does not contain any formula
	 */
	public static int hasFormula(OleAutomation rangeAutomation){
		
		int[] hasFormulaPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"HasFormula"});
		
		Variant result = rangeAutomation.getProperty(hasFormulaPropertyIds[0]);
		
		if(result==null)
			return -1;
		
		if(result.getType()==1){
			result.dispose();
			return 0;
		}
			
		boolean hasFormula = result.getBoolean();
		result.dispose();
		
		return hasFormula? 1 : -1;
	}
	
	/**
	 * Get special cells from the given range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @param type an integer that represents the type of special cells to get. For more info see XlCellType enumeration. 
	 * @return an OleAutomation to access the range of special cells
	 */
	public static OleAutomation getSpecialCells(OleAutomation rangeAutomation, int type){
		
		int[] specialCellsMethodIds = rangeAutomation.getIDsOfNames(new String[]{"SpecialCells", "Type"});
		
		Variant[] args = new Variant[1];
		args[0] = new Variant(type);
		int argsIds[] = Arrays.copyOfRange(specialCellsMethodIds, 1, specialCellsMethodIds.length);
		
		Variant result = rangeAutomation.invoke(specialCellsMethodIds[0], args, argsIds);
		for (Variant arg : args)
			arg.dispose();
		
		if(result==null)
			return null; 
		
		OleAutomation specialCells = result.getAutomation();
		result.dispose();
		
		return specialCells;
	}
	
	
	/**
	 * Draw a border around the range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @param lineStyle one of the constants of XlLineStyle 
	 * @param weight one of the constants of XlBorderWeight
	 * @param colorIndex the border color, as an index into the current color palette or as an XlColorIndex constant.
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean  drawBorderAroundRange(OleAutomation rangeAutomation, int lineStyle, double weight, long color){
		
		int[] borderAroundMethodIds = rangeAutomation.getIDsOfNames(new String[]{"BorderAround","LineStyle", "Weight", "Color"}); // "ColorIndex" 
		Variant methodParams[] = new Variant[3];
		methodParams[0] = new Variant(lineStyle); // line style (e.g., continuous, dashed ) 
		methodParams[1] = new Variant(weight); // border weight  (e.g., thick, thin )
		methodParams[2] = new Variant(color); // RGB color as a long value 
	
		int[] paramIds = Arrays.copyOfRange(borderAroundMethodIds, 1, borderAroundMethodIds.length);
		Variant result = rangeAutomation.invoke(borderAroundMethodIds[0], methodParams, paramIds);
		
		for (Variant v : methodParams) {
			v.dispose();
		}	
		
		if(result==null){
			return false;
		}
		
		result.dispose();
		return true;
	}
	
	
	/**
	 * Remove border around the range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean removeBorderAroundRange(OleAutomation rangeAutomation){
		 
		int[] borderAroundMethodIds = rangeAutomation.getIDsOfNames(new String[]{"BorderAround", "LineStyle"});
		Variant methodParams[] = new Variant[1];
		
		int xlLineStyleNone = -4142; // no line
		methodParams[0] = new Variant(xlLineStyleNone);  
	
		int[] paramIds = Arrays.copyOfRange(borderAroundMethodIds, 1, borderAroundMethodIds.length);
		Variant result = rangeAutomation.invoke(borderAroundMethodIds[0], methodParams, paramIds);
		
		for (Variant v : methodParams) {
			v.dispose();
		}	
		
		if(result==null){
			return false;
		}
		
		result.dispose();
		return true;
	}
	
	
	/**
	 * Delete the given range of cells 
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return true if the operation succeeded, false otherwise
	 */
	public static boolean deleteRange(OleAutomation rangeAutomation){
		
		int[] deleteMethodIds = rangeAutomation.getIDsOfNames(new String[]{"Delete"});		
		Variant result = rangeAutomation.invoke(deleteMethodIds[0]);
		
		if(result == null){
			return false;
		}
		
		result.dispose();
		return true;
	}
	
	
	/**
	 * Check if the first range contains (completely) the second range based on the given addresses 
	 * @param rangeAddress1 the address of the first range (i.e. the potential container) 
	 * @param rangeAddress2 the address of the second range (i.e. the range that might be contained by the first one)
	 * @return true if the first range contains the second range, false otherwise. 
	 */
	public static boolean checkForContainment(String rangeAddress1 , String rangeAddress2){
		
		String r1Cells[] = rangeAddress1.split(":");
		String r2Cells[] = rangeAddress2.split(":");
		
		String r1TopLeft = null, r1DownRight = null, r2TopLeft = null, r2DownRight = null;	
		r1TopLeft =  r1Cells[0];
		if(r1Cells.length == 1){
			r1DownRight = r1Cells[0];
		}else{
			r1DownRight = r1Cells[1];
		}
		
		r2TopLeft = r2Cells[0];	
		if(r2Cells.length == 1){
			r2DownRight = r2Cells[0];
		}else{
			r2DownRight = r2Cells[1];
		}
			
		int topColComp = compareCellsByColumn(r1TopLeft, r2TopLeft);
		int topRowComp = compareCellsByRow(r1TopLeft, r2TopLeft);
			
		int downColComp = compareCellsByColumn(r1DownRight, r2DownRight);
		int downRowComp = compareCellsByRow(r1DownRight, r2DownRight);
				
		boolean downRightCellContained =  topColComp<=0 && topRowComp<=0;
		boolean topLeftCellContained = downColComp>=0 && downRowComp>=0; 
		
		if(downRightCellContained && topLeftCellContained)
			return true;
		
		return false;
	}
	
	
	/**
	 * Check if the first range contains at least a part of the second range based on the given addresses. 
	 * In other words, check if the ranges share cells. 
	 * @param rangeAddress1 the address of the first range 
	 * @param rangeAddress2 the address of the second range 
	 * @return true if the first range shares cells with the second range, false otherwise. 
	 */
	public static boolean checkForPartialContainment(String rangeAddress1 , String rangeAddress2){
		
		String r1Cells[] = rangeAddress1.split(":");
		String r2Cells[] = rangeAddress2.split(":");
		
		String r1TopLeft = null, r1DownRight = null, r2TopLeft = null, r2DownRight = null;	
		r1TopLeft =  r1Cells[0];	
		if(r1Cells.length == 1){
			r1DownRight = r1Cells[0];
		}else{
			r1DownRight = r1Cells[1];
		}
		
		r2TopLeft = r2Cells[0];	
		if(r2Cells.length == 1){
			r2DownRight = r2Cells[0];
		}else{
			r2DownRight = r2Cells[1];
		}
					
		int topDownColComp = compareCellsByColumn(r1TopLeft, r2DownRight);
		int topDownRowComp = compareCellsByRow(r1TopLeft, r2DownRight);
		
		int downTopColComp = compareCellsByColumn(r1DownRight, r2TopLeft);
		int downTopRowComp = compareCellsByRow(r1DownRight, r2TopLeft);
		
		boolean rowsNotItersecting =  downTopRowComp<0 || topDownRowComp>0; 
		boolean colsNotIntersecting =  downTopColComp<0 || topDownColComp>0; 
		
		if(!(rowsNotItersecting || colsNotIntersecting))
			return true;
		
		return false; 
	}
	
	/**
	 * Compare two cells based on their column address 
	 * @param cell1Address a string that represents the address of the first cell
	 * @param cell2Address a string that represents the address of the second cell
	 * @return it returns 0 if the cells have the same column address, 
	 * a negative number if the second cell has higher (Alphabetically) column address, 
	 * a positive number if the first cell has a higher (Alphabetically) column address
	 */
	public static int compareCellsByColumn(String cell1Address, String cell2Address){
		
		String col1 =  cell1Address.replaceAll("[0-9\\$]+","");
		String col2 =  cell2Address.replaceAll("[0-9\\$]+","");
		
		if(col1.length()>col2.length())
			return 1;
		
		if(col1.length()<col2.length())
			return -1;
			
		return col1.compareTo(col2);
	}
	
	/**
	 * Compare two cells based on their row number 
	 * @param cell1Address a string that represents the address of the first cell
	 * @param cell2Address a string that represents the address of the second cell
	 * @return it returns 0 if the cells have the same row, 
	 * a negative number if the second cell has higher row number, 
	 * a positive number if the first cell has a higher row number
	 */
	public static int compareCellsByRow(String cell1Address, String cell2Address){
		
		int row1 =  Integer.valueOf(cell1Address.replaceAll("[^0-9]+",""));
		int row2 =  Integer.valueOf(cell2Address.replaceAll("[^0-9]+",""));
				
		return row1 - row2;
	}
}
