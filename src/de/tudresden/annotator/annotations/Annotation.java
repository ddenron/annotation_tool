/**
 * 
 */
package de.tudresden.annotator.annotations;

import java.util.Collection;
import java.util.HashMap;
import java.util.LinkedHashMap;

/**
 * This class represents an abstract annotation in a spreadsheet  
 * @author Elvis Koci
 * @param <T> The class of the objects that act as children (dependent) Annotations  
 */
public abstract class Annotation < T extends Annotation<?>>{
	
	/*
	 * A hash map that organizes (buckets) annotations by class label 
	 */
	protected HashMap <String, LinkedHashMap<String, T>> annotationsByClass ; 
	
	/*
	 * A linkedhashmap that stores all the annotations that are contained by this annotation
	 * In other words, all annotations in the linkedhashmap depend on this annotation object.  
	 */
	protected LinkedHashMap <String, T> allAnnotations;
	
	
	public Annotation(){
		this.annotationsByClass =  new HashMap<String, LinkedHashMap<String, T>>();
		this.allAnnotations =  new LinkedHashMap<String, T>();
	}
	
	/**
	 * Add a new annotation
	 * @param key a string that is used as an id (key) for the annotation object
	 * @param annotation the annotation object to add
	 */
	public void addAnnotation(String key, T annotation){
		this.allAnnotations.put(key, annotation);
	}
	
	
	/**
	 * Remove the annotation that has the given key
	 * @param key an object that is used as an id (key) for the annotation object
	 */
	public void removeAnnotation(String key){
		this.allAnnotations.remove(key);
	}
	
	/**
	 * Get the annotation by id (key)
	 * @param key a string that is used as an id (key) for the annotation object
	 * @return the annotation object
	 */
	public T getAnnotation(String key){
		return this.allAnnotations.get(key);
	}
	
	
	/**
	 * Get all annotations 
	 * @return a collection of annotation objects
	 */
	public Collection<T> getAllAnnotations(){
		return this.allAnnotations.values();
	}
	
	
	/**
	 * Get all the annotations that have the given class label
	 * @param label a string that represents the label of a Annotation class
	 * @return a collection of annotation objects
	 */
	public Collection<T> getAnnotationsByClass(String classLabel){
		HashMap<String, T> map = annotationsByClass.get(classLabel);
		if(map==null)
			return null;
		
		return map.values();
	}
	
	
	/**
	 * Add an annotation to the set containing annotations of the same class 
	 * @param label a string that represents the label of a Annotation class
	 * @param anntationId a string that represents the id (key) of the annotation object
	 * @param annotation an object that represents the annotation to add
	 */
	public void addAnnotationToBucket(String classLabel, String key, T annotation){
		LinkedHashMap<String, T>  map = annotationsByClass.get(classLabel);
		if(map == null){
			map = new LinkedHashMap<String, T>();
		}
		
		map.put(key, annotation);
		
		this.annotationsByClass.put(classLabel, map);
	}

	
	/**
	 * Remove an annotation from the set containing annotations of the same class 
	 * @param classLabel a string that represents the label of a Annotation class
	 * @param key a string that is used as an id (key) for the annotation object
	 */
	public void removeAnnotationFromBucket(String classLabel, String key){
		
		HashMap<String, T>  map = annotationsByClass.get(classLabel);
		if(map == null)
			return;
		
		map.remove(key);		
	}

	
	/**
	 * Remove all annotation having the given class label
	 * @param classLabel a string that represents the label of a Annotation class
	 */
	public void removeAllAnnotationsOfClass(String classLabel){
		
		LinkedHashMap<String, T>  map = annotationsByClass.get(classLabel);
		
		if(map == null)
			return;
		
		map.clear();		
		annotationsByClass.put(classLabel, map);
	}
	
	
	/**
	 * Remove all annotations 
	 */
	public void removeAllAnnotations(){
		this.allAnnotations.clear();
		this.annotationsByClass.clear();
	}
	
	
	/**
	 * Check if the given annotation object is equal to this one
	 * @param annotation the annotation object to compare this object to
	 * @return true if the objects are equal, false otherwise
	 */
	public abstract boolean equals(Annotation<T> annotation);

	
	/**
	 * Get the hashcode of this object
	 * @return the hashcode of the object
	 */
	public abstract int hashCode();
	
}
