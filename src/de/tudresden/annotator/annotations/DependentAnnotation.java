/**
 * 
 */
package de.tudresden.annotator.annotations;

/**
 * @author Elvis Koci
 * @param <P> The class of the object that acts as the Parent Annotation 
 */
public abstract class DependentAnnotation< P extends Annotation<?>> extends Annotation<RangeAnnotation>{

	/*
	 * The parent Annotation of this annotation object
	 * This annotation object is dependent from the parent  
	 */
	private P parent;

	
	/**
	 * @return the parent
	 */
	public P getParent() {
		return parent;
	}

	/**
	 * @param parent the parent to set
	 */
	public void setParent(P parent) {
		this.parent = parent;
	}
	
}
