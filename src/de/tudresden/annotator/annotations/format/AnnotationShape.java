/**
 * 
 */
package de.tudresden.annotator.annotations.format;

/**
 * @author Elvis Koci
 */
public enum AnnotationShape {
	
	BORDERAROUND (1),
	RANGEFILL (2),
	TEXTBOX (3),
	SHAPE (4),
	COMPLEXSHAPE (5);
	
	private final int code;

	private AnnotationShape(int itemCode){
		this.code = itemCode;
	}

	/**
	 * @return the code
	 */
	public int getCode() {
		return code;
	}
	
	/**
	 * 
	 * @param code
	 * @return
	 */
	public static AnnotationShape getByCode(int code){
		for (AnnotationShape tool : values()) {
			if (tool.getCode() == code) {
				return tool;
			}
		}		
		return null;
	}
}
