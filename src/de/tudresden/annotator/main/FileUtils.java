/**
 * 
 */
package de.tudresden.annotator.main;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Collection;

import org.eclipse.swt.SWT;
import org.eclipse.swt.SWTError;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.MessageBox;

import de.tudresden.annotator.annotations.NotApplicableStatus;
import de.tudresden.annotator.annotations.RangeAnnotation;
import de.tudresden.annotator.annotations.WorkbookAnnotation;
import de.tudresden.annotator.annotations.utils.AnnotationHandler;
import de.tudresden.annotator.annotations.utils.AnnotationStatusSheet;
import de.tudresden.annotator.annotations.utils.RangeAnnotationsSheet;
import de.tudresden.annotator.annotations.utils.UsageLogSheet;
import de.tudresden.annotator.oleutils.ApplicationUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;

/**
 * @author Elvis Koci
 */
public class FileUtils {
	
	public static final String CurrentProgressFileName = "annotated"; 
	
	public static final String completedFilesFolderName = "completed";
	public static final String notApplicableFilesFolderName = "not_applicable";
	public static final String ambiguousFilesFolderName = "ambiguous";
	public static final String inProgressFilesFolderName = "in_progress";
	public static final String otherFilesFolderName = "other";
	
	
	/**
	 * Open an excel file for annotation
	 */
	 public static void fileOpen(){
		
		// Select the excel file
		FileDialog dialog = Launcher.getInstance().createFileDialog(SWT.OPEN);
		String filePath = dialog.open();
		fileOpen(filePath);
	}
	
	 
	/**
	 * Open an excel file for annotation
	 * @param filePath the absolute path of the file to open
	 */
	 public static void fileOpen(String filePath){
		 	 
		// if no file was selected, return
		if (filePath == null) return;
		
		Launcher mw = Launcher.getInstance();
		
		// dispose OleControlSite if it is not null. 
		mw.disposeControlSite();
				
	    if (mw.isControlSiteNull()) {
			int index = filePath.lastIndexOf('.');
			if (index != -1) {
				String fileExtension = filePath.substring(index + 1); 
				if (fileExtension.equalsIgnoreCase("xls") || 
						fileExtension.equalsIgnoreCase("xlsx")) { //  || fileExtension.equalsIgnoreCase("xlsm")	
					
					try {		    	
						File excelFile = new File(filePath);
				        
				        // embed the excel file and set up the user interface
				        mw.embedExcelFile(excelFile);
				        
				    } catch (SWTError e) {
				        e.printStackTrace();
				        System.out.println("Unable to open ActiveX Control");
				        return;
				    }	    	  
				   
				}else{
					MessageBox msgbox = mw.createMessageBox(SWT.ICON_ERROR);
					msgbox.setMessage("The selected file format is not recognized: ."+fileExtension);
					msgbox.open();
				}
			}
	    }
	}
	 
	 
	/**
	 * Save all the annotation progress 
	 * @param embeddedWorkbook an OleAutomation that provides access to the functionalities of the embedded workbook
	 * @param filePath the path where to save the file
	 * @param beforeFileClose true if progress is saved before closing the file or exiting the application, 
	 * false if file will remain open after save.
	 * @return true if progress was successfully saved, false otherwise. 
	 */
	public static boolean saveProgress(OleAutomation embeddedWorkbook, String filePath, boolean beforeFileClose){
		
		// deactivate alerts before save
		OleAutomation application = WorkbookUtils.getApplicationAutomation(embeddedWorkbook);
		Launcher.getInstance().deactivateControlSite();
		ApplicationUtils.setDisplayAlerts(application, false);
		
		WorkbookAnnotation wa = AnnotationHandler.getWorkbookAnnotation();
	
		if(!wa.getAllAnnotations().isEmpty() || wa.hasActiveStatus() || wa.hasActiveWorksheetStatus()){
			
			// save the status of all worksheet annotations and the workbook annotation 
			AnnotationStatusSheet.saveAnnotationStatuses(embeddedWorkbook);		
			
			
			boolean isSaveTime = UsageLogSheet.saveTime(embeddedWorkbook, "Save");
			if (!isSaveTime){
				int messageStyle = SWT.ICON_ERROR;
				MessageBox message = Launcher.getInstance().createMessageBox(messageStyle);
				message.setMessage("ERROR: there was an issue with the systime! Please inform the admistrator!");
				message.open();
				return false;
			}
			
			// delete all shape annotations
			AnnotationHandler.deleteAllShapeAnnotations(embeddedWorkbook);
				
			// unprotect the workbook structure
			WorkbookUtils.unprotectWorkbook(embeddedWorkbook);
			// unprotect all the sheets
		    WorkbookUtils.unprotectAllWorksheets(embeddedWorkbook);
						
			// protect and hide the range_annotations sheet before save
			RangeAnnotationsSheet.protect(embeddedWorkbook);
			RangeAnnotationsSheet.setVisibility(embeddedWorkbook, false);
			
			// protect and hide the annotation_status sheet before save
			AnnotationStatusSheet.protect(embeddedWorkbook);
			AnnotationStatusSheet.setVisibility(embeddedWorkbook, false);
			
			// protect and hide the usage_log sheet before save
			UsageLogSheet.protect(embeddedWorkbook);
			UsageLogSheet.setVisibility(embeddedWorkbook, false);
						
		}else{
			
			// unprotect the workbook structure
			WorkbookUtils.unprotectWorkbook(embeddedWorkbook);
			// delete all sheets that store metadata about the annotations, 
			// if there are no annotation
			RangeAnnotationsSheet.delete(embeddedWorkbook);
			AnnotationStatusSheet.delete(embeddedWorkbook);
			UsageLogSheet.delete(embeddedWorkbook);
			
			// unprotect all the sheets
			WorkbookUtils.unprotectAllWorksheets(embeddedWorkbook);
		}
							
		// save the file
		boolean isSuccess = WorkbookUtils.saveWorkbookAs(embeddedWorkbook, filePath, null);
		
		// activate alerts after save
		ApplicationUtils.setDisplayAlerts(application, true);

		WorkbookUtils.closeEmbeddedWorkbook(embeddedWorkbook, false);
		Launcher.getInstance().setEmbeddedWorkbook(null);
				
		String newPath =  moveFileToStatusDirectory();
			
		if(!beforeFileClose){			
			
			fileOpen(newPath);
					
			OleAutomation reopenedWorkbook = Launcher.getInstance().getEmbeddedWorkbook();	
			
			// save time the file was re-opened 
			boolean isOpenTimeSaved = UsageLogSheet.saveOpenTime(reopenedWorkbook);
			if(!isOpenTimeSaved){
        		int messageStyle = SWT.ICON_ERROR;
				MessageBox message = Launcher.getInstance().createMessageBox(messageStyle);
				message.setMessage("ERROR: there was an issue with the systime! Please inform the admistrator!");
				message.open();
        	}
			
			
			// turn off screen updating before re-drawing range annotations. this speeds up the process. 
			OleAutomation reopenedApplication = WorkbookUtils.getApplicationAutomation(reopenedWorkbook);		
			ApplicationUtils.setScreenUpdating(reopenedApplication, false);
			
			// draw again the range annotations  
			Collection<RangeAnnotation> collection= AnnotationHandler.getWorkbookAnnotation().getAllAnnotations();
			RangeAnnotation[] rangeAnnotations = collection.toArray(new RangeAnnotation[collection.size()]);
			if(rangeAnnotations!=null){			
				
				// update workbook annotation and re-draw all the range annotations  	
				AnnotationHandler.drawManyRangeAnnotations(reopenedWorkbook, rangeAnnotations, false);
				// TODO: reconsider AnnotationHandler.drawManyRangeAnnotationsOptimized();
			}
			
			// make range_annotations sheet again visible
			RangeAnnotationsSheet.setVisibility(reopenedWorkbook, true);
			
			// turn on screen updating after all range annotations are re-drawn
			ApplicationUtils.setScreenUpdating(reopenedApplication, true);
		}
		return isSuccess;
	}


	/**
	 * Move the opened (embedded) excel file to the directory that corresponds to 
	 * its current annotation status. For example, if the file was marked as "Completed",
	 * it will be moved to the folder where all the completed files are grouped (placed).
	 * 
	 */
	public static String moveFileToStatusDirectory(){
		
		String fileName = Launcher.getInstance().getFileName();
		String fileDirPath = Launcher.getInstance().getDirectoryPath();
		
		File file = new File(fileDirPath+"\\"+fileName);
		File directory = new File(fileDirPath); 
				
		String originalDir = fileDirPath; 
		if(directory.getName().compareTo(completedFilesFolderName)==0 || 
		   directory.getName().compareTo(notApplicableFilesFolderName)==0 || 
		   directory.getName().compareTo(inProgressFilesFolderName)==0 ||
		   directory.getName().compareTo(ambiguousFilesFolderName)==0){
			
			originalDir = directory.getParentFile().getAbsolutePath();
			
		}else if(directory.getParentFile().getName().compareTo(notApplicableFilesFolderName)==0){
			originalDir = directory.getParentFile().getParentFile().getAbsolutePath();
		}	

		
		File newLocation = file;	
		if(AnnotationHandler.getWorkbookAnnotation().isCompleted()){
			
			File completed = new File (originalDir+"\\"+completedFilesFolderName);
			if(!completed.exists())
				completed.mkdir();
			
			newLocation = new File(completed.getAbsolutePath()+"\\"+fileName); 
				
			if(!file.getParent().equals(completed))
				moveFile(file, newLocation);
					
		}else if(AnnotationHandler.getWorkbookAnnotation().isNotApplicable()){
			
			NotApplicableStatus status = AnnotationHandler.getWorkbookAnnotation().getNotApplicableStatus();
			String subFolderName = status.getStatusName().replaceAll("/", "_");
			File notApplicable = new File (originalDir+"\\"+notApplicableFilesFolderName+"\\"+subFolderName.toLowerCase());
			if(!notApplicable.exists())
				notApplicable.mkdirs();
			
			newLocation = new File(notApplicable.getAbsolutePath()+"\\"+fileName); 
			
			if(!file.getParent().equals(notApplicable))
				moveFile(file, newLocation);
			
		}else if(AnnotationHandler.getWorkbookAnnotation().isAmbiguous()){
			
			File ambiguous = new File (originalDir+"\\"+ambiguousFilesFolderName);
			if(!ambiguous.exists())
				ambiguous.mkdir();
			
			newLocation = new File(ambiguous.getAbsolutePath()+"\\"+fileName); 
			
			if(!file.getParent().equals(ambiguous))
				moveFile(file, newLocation);
			
		}else{
			
			File inProgress = new File (originalDir+"\\"+inProgressFilesFolderName);
			if(!inProgress.exists())
				inProgress.mkdir();
			
			newLocation = new File(inProgress.getAbsolutePath()+"\\"+fileName); 
			
			if(!file.getParent().equals(inProgress))
				moveFile(file, newLocation);			
		}
		
		return newLocation.getAbsolutePath();
	}
	
	
	/**
	 * 
	 * @param file
	 * @param directory
	 */
	private static void moveFile(File file, File directory){
		
		try {
			Files.move(file.toPath(), directory.toPath());
		} catch (IOException e) {
			
			MessageBox message = Launcher.getInstance().createMessageBox(SWT.ICON_ERROR);
			message.setText("ERROR");
			message.setMessage("ERROR: Could not move file \""+file.getName()+"\" "
					+ "to the directory \""+directory.getName()+"\". \n\n"
							+ e.toString());
			message.open();
		}
		
	}
	
	
	/**
	 * 
	 * @param directory
	 * @param fileName
	 * @param status
	 * @return
	 */
	public static boolean markFileAsAnnotated(String directory, String fileName, int status){
		
		File file = new File(directory+"\\"+CurrentProgressFileName);
		
		try {
			if (!file.exists()) {
				file.createNewFile();
			}
			
			FileWriter fw = new FileWriter(file.getAbsoluteFile());
			BufferedWriter bw = new BufferedWriter(fw);
			
			String content = fileName+"\t"+status+"\n";
			bw.write(content);
			bw.close();
			
		} catch (IOException e) {
			e.printStackTrace();
			return false;
		}	
		return true;
	}
}
