/**
 * 
 */
package de.tudresden.annotator.annotations;

/**
 * @author Elvis Koci
 *
 */
public enum NotApplicableStatus {

	NONE (1, "None", ""), // Special value reserved for sheets or workbooks that are applicable (i.e., NotApplicable = false) 
	FORMTEMPLATE (2, "Form/Template", "Sheets are modeled so to be re-used over and over for similar tasks."),
	REPORTBALANCE (3, "Report/Balance", "Reporting the current status of money related transactions and processes."),
	CHART (4, "Chart/Graph", "When the contents are data visualizations, such as bar-chart, pie-chart, etc."),
	LIST (7, "List", "When the sheet/s contain lists of items."),
	NOHEADER (6, "No-Header", "When there are well structured Data in rows and columns, but no Header to describe them."),
	NOTENGLISH (7, "Not-English", "When the file contains foreign words, not in English language."),
	OTHER (8, "Other", "Use this when none of the above options is a match."),
	MULTINA (8, "Multi_Various_NA", ""); // Used only for workbooks, when there are multiple sheets with different not applicable statuses 
	
	private final int statusId;
	private final String statusName;
	private final String description;
	
	private NotApplicableStatus(int id, String name, String dscr){
		this.statusId = id;
		this.statusName = name;
		this.description = dscr;
	}
	
	/**
	 * @return the status id
	 */
	public int getStatusId() {
		return this.statusId;
	}
	
	/**
	 * @return the status name
	 */
	public String getStatusName() {
		return this.statusName;
	}
	
	
	/**
	 * @return the status description
	 */
	public String getStatusDescription() {
		return this.description;
	}
	
	/**
	 * 
	 * @param id
	 * @return the NotApplicableStatus that matches the given id
	 */
	public static NotApplicableStatus getById(int id){	
		for (NotApplicableStatus status : values()) {
			if (status.getStatusId() == id) {
				return status;
			}
		}	
		return null;
	}
	
	
	/**
	 * 
	 * @param name
	 * @return the NotApplicableStatus that matches the given name
	 */
	public static NotApplicableStatus getByName(String name){
		
		NotApplicableStatus naStatus = null;
		try{
			naStatus = NotApplicableStatus.valueOf(name.toUpperCase());		
		}catch(IllegalArgumentException e){}
		
		if(naStatus == null){
			for (NotApplicableStatus nas : values()) {
				if (nas.getStatusName().compareToIgnoreCase(name)==0) {
					naStatus = nas;
				}
			}	
		}
		return naStatus;
	}
}
