/**
 * 
 */
package de.tudresden.annotator.annotations.utils;

import java.util.LinkedHashMap;

import de.tudresden.annotator.annotations.format.AnnotationClass;
import de.tudresden.annotator.annotations.format.AnnotationShape;
import de.tudresden.annotator.oleutils.ColorFormatUtils;

/**
 * @author Elvis Koci
 */
public class ClassGenerator {
	
	private static final LinkedHashMap<String, AnnotationClass> annotationClasses;
	
	static{
		annotationClasses = new LinkedHashMap<String, AnnotationClass>();
		AnnotationClass[] classes = createAnnotationClasses();
		for (AnnotationClass annotationClass : classes) {
			annotationClasses.put(annotationClass.getLabel(), annotationClass);
		}
	}
		
	private static AnnotationClass[] createAnnotationClasses(){
		
		AnnotationClass[] classes =  new AnnotationClass[8];
		
		long white =  ColorFormatUtils.getRGBColorAsLong(255, 255, 255);
		long bordo = ColorFormatUtils.getRGBColorAsLong(192, 0, 0);
		long blue_accent5 = ColorFormatUtils.getRGBColorAsLong(68, 114, 196);
		long yellow = ColorFormatUtils.getRGBColorAsLong(255, 255, 49);
		long green_accent6 = ColorFormatUtils.getRGBColorAsLong(112, 173, 71);
		long orange_accent2 = ColorFormatUtils.getRGBColorAsLong(237, 125, 49);
		long blue_accent1 = ColorFormatUtils.getRGBColorAsLong(91, 155, 213);
		long greyDark = ColorFormatUtils.getRGBColorAsLong(118, 113, 113);
		long purple = ColorFormatUtils.getRGBColorAsLong(112, 48, 160);
		long dark_gold = ColorFormatUtils.getRGBColorAsLong(128, 96, 0);

		
		// table can contains all the other classes
		classes[0] = createShapeAnnotationClass("Table", false, 1, 0, 1, true, blue_accent5, 2, true, greyDark, true, false, false, null); 
		
		// Annotations of the following classes can be outside of a table or inside. Example: tables can share Notes
		classes[5] = createTextBoxAnnotationClass("MetaTitle", orange_accent2, 0.80, true, white, true, false, null); 
		classes[6] = createTextBoxAnnotationClass("Notes", yellow, 0.75, true, white, true, false, null);
		classes[4] = createTextBoxAnnotationClass("Derived", bordo, 0.80,  true, white, true, false, null);
		classes[7] = createTextBoxAnnotationClass("Other", dark_gold, 0.75, true, white, true, false, null);
		
		// Annotations of the following classes can only be inside a table
		classes[2] = createTextBoxAnnotationClass("Header", blue_accent1, 0.70,  true, white, true, true, classes[0]);
		classes[1] = createTextBoxAnnotationClass("Data", green_accent6, 0.80,  true, white, true, true, classes[0]);	
		classes[3] = createTextBoxAnnotationClass("GroupHead", purple, 0.80, true, white, true, true, classes[0]);
	
		return classes;
	}
	
	
	private static AnnotationClass createShapeAnnotationClass(String label, boolean isUseText, int shapeType, long fillColor, int fillPattern,
													boolean useLine, long lineColor, int lineWeight, boolean useShadow, long shadowColor, 
												    boolean isContainer, boolean isContainable, boolean isDependent, AnnotationClass container){
		
		AnnotationClass ac = new AnnotationClass(label, AnnotationShape.SHAPE, false);
		
		if(isContainer){
			ac.setHasFill(false);
		}else{
			ac.setHasFill(true);
		}
		ac.setColor(fillColor);
		ac.setFillPattern(fillPattern);
		
		
		ac.setUseShadow(useShadow);
		ac.setShadowColor(shadowColor);
		
		ac.setUseLine(useLine);
		ac.setLineColor(lineColor);
		ac.setLineWeight(lineWeight);
		
		ac.setShapeType(shapeType);
		
		ac.setUseText(isUseText);
		
		ac.setIsContainer(isContainer);
		ac.setCanBeContained(isContainable);
		ac.setIsDependent(isDependent);
		ac.setContainer(container);
			
		return ac; 
	}

	private  static AnnotationClass createTextBoxAnnotationClass(String label, long backcolor, double transperency, boolean useText, long textColor,
														    boolean isContainable, boolean isDependent, AnnotationClass container){
		
		AnnotationClass ac = new AnnotationClass(label, AnnotationShape.TEXTBOX, backcolor);
		
		ac.setHasFill(true);
		ac.setColor(backcolor);
		ac.setFillTransparency(transperency);
		
		ac.setUseShadow(false);
		ac.setUseLine(false);
		
		ac.setUseText(useText);
		ac.setText(label.toUpperCase());
		ac.setTextColor(textColor);
		
		ac.setCanBeContained(isContainable);
		ac.setIsDependent(isDependent);
		ac.setContainer(container);
		
		return ac; 
	}


	/**
	 * @return the annotationclasses
	 */
	public static LinkedHashMap<String, AnnotationClass> getAnnotationClasses() {
		return annotationClasses;
	}
	
}
