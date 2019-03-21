/**
 * 
 */
package de.tudresden.annotator.main;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.widgets.Menu;
import org.eclipse.swt.widgets.MenuItem;

import de.tudresden.annotator.annotations.NotApplicableStatus;
import de.tudresden.annotator.annotations.WorkbookAnnotation;
import de.tudresden.annotator.annotations.WorksheetAnnotation;
import de.tudresden.annotator.annotations.utils.AnnotationHandler;
import de.tudresden.annotator.annotations.utils.RangeAnnotationsSheet;
import de.tudresden.annotator.oleutils.WindowUtils;

/**
 * @author Elvis Koci
 */
public class BarMenuUtils {
	
	protected static void adjustBarMenuForSheet(String sheetName){
					
		BarMenu  menuBar = Launcher.getInstance().getMenuBar();
		MenuItem[] menuItems = menuBar.getMenuItems();
		
		MenuItem annotationsMenu = null;
		MenuItem windowMenu = null;
		MenuItem goToMenu = null;
		for (MenuItem menuItem : menuItems) {
			if(menuItem.getID()==2000000){
				annotationsMenu = menuItem;
			}
			
			if(menuItem.getID()==3000000){
				windowMenu = menuItem;
			}
			
			if(menuItem.getID()==5000000){
				goToMenu = menuItem;
			}
		}
		
		MenuItem[] annotationsMenuItems = annotationsMenu.getMenu().getItems();
		
		// the active worksheet annotation
		WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
		WorksheetAnnotation  sheetAnnotation = workbookAnnotation.getWorksheetAnnotations().get(sheetName);
		
		
		// if the active sheet is the RangeAnnotationsSheet, then disable all annotation menus
		// if there does not exists a sheet annotation, do the same.
		if(sheetAnnotation==null || sheetName.compareTo(RangeAnnotationsSheet.getName())==0){
			for (MenuItem menuItem : annotationsMenuItems) {
				if(menuItem.getID()!=2030000){ // File as
					menuItem.setEnabled(false);
					disableAllSubMenus(menuItem.getMenu());
				}else{
					menuItem.setEnabled(true);
				}
			}
			
			if (goToMenu!=null){
				goToMenu.setEnabled(true);
				disableAllSubMenus(goToMenu.getMenu());
			}
			
		}else{
			
			// adjust according to the status of the sheet: Completed, NotApplicable, or non of these 		
			if(sheetAnnotation.isCompleted()){
			
				for (MenuItem menuItem : annotationsMenuItems) {
					
					if(menuItem.getID() == 2010000 || menuItem.getID() == 2050000 ||
					   menuItem.getID() == 2080000 || menuItem.getID() == 2090000 ||
					   menuItem.getID() == 2060000 ){ 
							// &Range as, &Delete, Undo, and Redo
							menuItem.setEnabled(false);
							disableAllSubMenus(menuItem.getMenu());
					}else{ 
							menuItem.setEnabled(true);
							if(menuItem.getID()==2020000){ // &Sheet as 
							
								MenuItem[] submenus = menuItem.getMenu().getItems();
								for (int i = 0; i < submenus.length; i++) {
									if(submenus[i].getID()==2020100){ // Not Applicable
										submenus[i].setEnabled(false);
										submenus[i].setSelection(false);
									}
									
									if(submenus[i].getID()==2020200){ // Completed
										submenus[i].setEnabled(true);
										submenus[i].setSelection(true);
									}
								}
							}
							
							if(menuItem.getID()!=2020000 && menuItem.getID()!=2030000){
								enableAllSubMenus(menuItem.getMenu());
							}
					}
				}
				
				for (MenuItem menuItem : windowMenu.getMenu().getItems()) {	
					if(menuItem.getID()==3010000){
						menuItem.setEnabled(false);
						break;
					}
				}
				
				if (goToMenu!=null){
					goToMenu.setEnabled(true);
					disableAllSubMenus(goToMenu.getMenu());
				}
							
			}else{
				
				if(sheetAnnotation.isNotApplicable()){				
					for (MenuItem menuItem : annotationsMenuItems) { 
						if(menuItem.getID()==2010000 || menuItem.getID()==2040000 || 
						   menuItem.getID()==2050000 || menuItem.getID()==2060000 || 
						   menuItem.getID()==2070000 || menuItem.getID() == 2080000 || 
						   menuItem.getID() == 2090000){  
						   // &Range as, &Hide, &Delete, and Show Annotations 
								menuItem.setEnabled(false);
								disableAllSubMenus(menuItem.getMenu());
						}else{
								menuItem.setEnabled(true);
								if(menuItem.getID()==2020000){ // &Sheet as 
									MenuItem[] submenus = menuItem.getMenu().getItems();
									for (int i = 0; i < submenus.length; i++) {
										if(submenus[i].getID()==2020100){ // Not Applicable
											submenus[i].setEnabled(true);
											submenus[i].setSelection(true);
											
											MenuItem[] naItems = submenus[i].getMenu().getItems();
											NotApplicableStatus activeStatus = sheetAnnotation.getNotApplicableStatus();
											for (int j = 0; j < naItems.length; j++){
												if(naItems[j].getText().compareToIgnoreCase(activeStatus.getStatusName())==0){					
													naItems[j].setEnabled(true);
													naItems[j].setSelection(true);
												}else{
													naItems[j].setEnabled(false);
												}
											}
										}
										
										if(submenus[i].getID()==2020200){ // Completed 
											submenus[i].setEnabled(false);
											submenus[i].setSelection(false);
										}
									}
								}
								
								if(menuItem.getID()!=2020000 && menuItem.getID()!=2030000){
									enableAllSubMenus(menuItem.getMenu());
								}
						}
					}
					
					for (MenuItem menuItem : windowMenu.getMenu().getItems()) {	
						if(menuItem.getID()==3010000){
							menuItem.setEnabled(false);
							break;
						}
					}
					
					if (goToMenu!=null){
						goToMenu.setEnabled(true);
						disableAllSubMenus(goToMenu.getMenu());
					}
					
				}else{
					
					boolean hasAnnotations = sheetAnnotation.getAllAnnotations().size() > 0;		
					if(!hasAnnotations){
						for (MenuItem menuItem : annotationsMenuItems) { 
							if(menuItem.getID() == 2040000 || menuItem.getID() == 2050000 || 
							   menuItem.getID() == 2070000 || menuItem.getID() == 2080000 ){
							   // &Hide, &Delete, &Show Annotations, Undo Annotation
							   menuItem.setEnabled(false);
							   disableAllSubMenus(menuItem.getMenu());
							}else if(menuItem.getID() == 2090000){ // Redo last annotation
								if(AnnotationHandler.getLastFromRedoList()==null){
										menuItem.setEnabled(false);
								}else{
										menuItem.setEnabled(true);
								}
								
							}else{
								
								menuItem.setEnabled(true);								
								enableAllSubMenus(menuItem.getMenu());
								unselectAllSubMenus(menuItem.getMenu());
							}
						}
						
						for (MenuItem menuItem : windowMenu.getMenu().getItems()) {
							if(menuItem.getID()==3010000){
								updateShowFormulas(menuItem);
								break;
							}
						}
								
					}else{
						for (MenuItem menuItem : annotationsMenuItems) { 
							if( menuItem.getID() == 2080000){	// Undo last annotations				
								if(AnnotationHandler.getLastFromUndoList()==null){
									menuItem.setEnabled(false);
								}else{
									menuItem.setEnabled(true);
								}
								
							}else if(menuItem.getID() == 2090000){ // Redo last annotation
								
								if(AnnotationHandler.getLastFromRedoList()==null){
									menuItem.setEnabled(false);
								}else{
									menuItem.setEnabled(true);
								}
							
							}else{
								
								menuItem.setEnabled(true);
								
								enableAllSubMenus(menuItem.getMenu());
								unselectAllSubMenus(menuItem.getMenu());
							}
						}
						
						for (MenuItem menuItem : windowMenu.getMenu().getItems()) {
							
							if(menuItem.getID()==3010000){
								updateShowFormulas(menuItem);
								break;
							}
						}
					}
					
					if (goToMenu!=null){
						goToMenu.setEnabled(true);
						enableAllSubMenus(goToMenu.getMenu());
					}
				}
			}
		}
	}
	
	
	protected static void adjustBarMenuForWorkbook(){
		
		WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
		
		BarMenu  menuBar = Launcher.getInstance().getMenuBar();
		MenuItem[] menuItems = menuBar.getMenuItems();
		
		MenuItem annotationsMenu = null;
		MenuItem goToMenu = null;
		for (MenuItem menuItem : menuItems) {
			if(menuItem.getID()==2000000){ // annotations Menu
				annotationsMenu = menuItem;
			}
			
			if(menuItem.getID()==5000000){ // goTo Menu
				goToMenu = menuItem;
			}
		}
		
		if(workbookAnnotation.isCompleted()){

			MenuItem[] annotationsMenuItems = annotationsMenu.getMenu().getItems();
			for (MenuItem menuItem : annotationsMenuItems) { 
				if(menuItem.getID() == 2010000 || menuItem.getID() == 2020000 ||
				   menuItem.getID() == 2040000 || menuItem.getID() == 2050000 ||
				   menuItem.getID() == 2060000 || menuItem.getID() == 2070000 ||
				   menuItem.getID() == 2080000 || menuItem.getID() == 2090000){ 
					// &Range as, &Sheet as, &Show Annotation, 
					// &Hide, &Delete, Show Formulas, Undo, and Redo
						menuItem.setEnabled(false);
						disableAllSubMenus(menuItem.getMenu());
				}else{
					menuItem.setEnabled(true);
					if(menuItem.getID()==2030000){ // &File as 
						MenuItem[] submenus = menuItem.getMenu().getItems();
						for (int i = 0; i < submenus.length; i++) {
							if(submenus[i].getID()==2030200){ // Completed 
								submenus[i].setEnabled(true);
								submenus[i].setSelection(true);
							}else{
								submenus[i].setEnabled(false);
								submenus[i].setSelection(false);
							}
						}
					}
				} 
			}		
			
			if (goToMenu!=null){
				goToMenu.setEnabled(true);
				disableAllSubMenus(goToMenu.getMenu());
			}
			
		}else{
			
			if(workbookAnnotation.isNotApplicable() || workbookAnnotation.isAmbiguous()){
				MenuItem[] annotationsMenuItems = annotationsMenu.getMenu().getItems();
				for (MenuItem menuItem : annotationsMenuItems) {
					if(menuItem.getID() == 2010000 || menuItem.getID() == 2020000 ||
					   menuItem.getID() == 2040000 || menuItem.getID() == 2050000 ||
					   menuItem.getID() == 2060000 || menuItem.getID() == 2070000 ||
					   menuItem.getID() == 2080000 || menuItem.getID() == 2090000){ 
						// &Range as, &Sheet as, &Show Annotation, 
						// &Hide, &Delete, Show Formulas, Undo, and Redo
						
							menuItem.setEnabled(false);
							disableAllSubMenus(menuItem.getMenu());
					}else{
						if(menuItem.getID() == 2030000 ){ // File as
							menuItem.setEnabled(true);
							MenuItem[] submenus = menuItem.getMenu().getItems();
							for (int i = 0; i < submenus.length; i++){
								if(workbookAnnotation.isNotApplicable()){
									if(submenus[i].getID()==2030100){ // Not Applicable
										submenus[i].setEnabled(true);
										submenus[i].setSelection(true);
										
//										MenuItem[] naItems = submenus[i].getMenu().getItems();
//										NotApplicableStatus activeStatus = workbookAnnotation.getNotApplicableStatus();
//										for (int j = 0; j < naItems.length; j++){
//											if(naItems[j].getText().compareToIgnoreCase(activeStatus.getStatusName())==0){					
//												naItems[j].setEnabled(true);
//												naItems[j].setSelection(true);
//											}else{
//												naItems[j].setEnabled(false);
//											}
//										}
										
									}else{
										submenus[i].setEnabled(false);
										submenus[i].setSelection(false);
									}
								}else if(workbookAnnotation.isAmbiguous()){
										
									if(submenus[i].getID()==2030300){ // Ambiguous
										submenus[i].setEnabled(true);
										submenus[i].setSelection(true);
									}else{
										submenus[i].setEnabled(false);
										submenus[i].setSelection(false);
									}
								}
							}
						}
					} 
				}	
				
				if (goToMenu!=null){
					goToMenu.setEnabled(true);
					disableAllSubMenus(goToMenu.getMenu());
				}
				
			}else{

				MenuItem[] annotationsMenuItems = annotationsMenu.getMenu().getItems();
				for (MenuItem menuItem : annotationsMenuItems) {
					if(menuItem.getID() == 2030000){ // &File as
						menuItem.setEnabled(true);
						enableAllSubMenus(menuItem.getMenu());
						unselectAllSubMenus(menuItem.getMenu());
					}
				}							
				String activeSheetName = Launcher.getInstance().getActiveWorksheetName();
				adjustBarMenuForSheet(activeSheetName);
			}			
		}		
	}
	
	protected static void adjustBarMenuForOpennedFile(){
		
		BarMenu  menuBar = Launcher.getInstance().getMenuBar();
		MenuItem[] menuItems = menuBar.getMenuItems();
		
		MenuItem fileMenu = null;
		for (MenuItem menuItem : menuItems) {
			if(menuItem.getID()==1000000){ // file menu
				fileMenu = menuItem;
			}
			
			if(menuItem.getID()==2000000){ // annotations menu
				menuItem.setEnabled(true);
				unselectAllSubMenus(menuItem.getMenu());
			}
			
			if(menuItem.getID()==3000000){ // window menu
				menuItem.setEnabled(true);
				unselectAllSubMenus(menuItem.getMenu());
				enableAllSubMenus(menuItem.getMenu());
			}
			
			if(menuItem.getID()==4000000){ // preferences menu
				menuItem.setEnabled(true);
				disableAllSubMenus(menuItem.getMenu());
			}
			
			if(menuItem.getID()==5000000){ // annotations menu
				menuItem.setEnabled(true);
				enableAllSubMenus(menuItem.getMenu());
			}
		}
		
		MenuItem[] fileMenuItems = fileMenu.getMenu().getItems();
		for (MenuItem menuItem : fileMenuItems) {
			if(!(menuItem.getID() == 1020000 || menuItem.getID() == 1030000)){ // Open Prev and Open Next
				menuItem.setEnabled(true);
			}	
		}		
		adjustBarMenuForWorkbook();
	}
	
	
	
	protected static void adjustBarMenuForFileClose(){
		BarMenu  menuBar = Launcher.getInstance().getMenuBar();
		MenuItem[] menuItems = menuBar.getMenuItems();
		
		MenuItem fileMenu = null;
		for (MenuItem menuItem : menuItems) {
			
			if(menuItem.getID()==1000000){ // file menu
				fileMenu = menuItem;
				fileMenu.setEnabled(true);
			}else{
				menuItem.setEnabled(true);
				disableAllSubMenus(menuItem.getMenu());
				unselectAllSubMenus(menuItem.getMenu());
			}		
		}
		
		MenuItem[] fileMenuItems = fileMenu.getMenu().getItems();
		for (MenuItem menuItem : fileMenuItems) {
			if(menuItem.getID() == 1010000 || menuItem.getID() == 1070000){ // Open and Exit
				menuItem.setEnabled(true);
			}else{
				menuItem.setEnabled(false);
				disableAllSubMenus(menuItem.getMenu());
				unselectAllSubMenus(menuItem.getMenu());
			}
		}
	}
	
	
	protected static void setEnabledForRangeAsMenuItem(boolean enabled){
		
		BarMenu  menuBar = Launcher.getInstance().getMenuBar();
		MenuItem[] menuItems = menuBar.getMenuItems();
		
		MenuItem annotationsMenuItem = null;
		for (MenuItem menuItem : menuItems) {
			if(menuItem.getID()==2000000){ // annotations menu
				annotationsMenuItem = menuItem;
				break;
			}
		}
		
		MenuItem[] annotationsMenuItems = annotationsMenuItem.getMenu().getItems();
		for (MenuItem menuItem : annotationsMenuItems) {
			if(menuItem.getID()==2010000){ // annotations menu
				menuItem.setEnabled(enabled);
				break;
			}
		}
	}
	
	
	protected static void updateShowFormulas(MenuItem showFormulasMenuItem){
		OleAutomation window = Launcher.getInstance().getEmbeddedWindow();
		boolean areFormulasDisplayed = WindowUtils.getDisplayFormulas(window);		
		showFormulasMenuItem.setSelection(areFormulasDisplayed);
	}
	
	
	private static void enableAllSubMenus(Menu cascadeMenu){
		if(cascadeMenu!=null){
			MenuItem[]  submenus = cascadeMenu.getItems();
			for (int i = 0; i < submenus.length; i++) {
				enableAllSubMenus(submenus[i].getMenu());
				submenus[i].setEnabled(true);
			}
		}
	}
	
	private static void disableAllSubMenus(Menu cascadeMenu){
		if(cascadeMenu!=null){
			MenuItem[]  submenus = cascadeMenu.getItems();
			for (int i = 0; i < submenus.length; i++) {
				disableAllSubMenus(submenus[i].getMenu());
				submenus[i].setEnabled(false);
			}
		}
	}
	
	private static void unselectAllSubMenus(Menu cascadeMenu){
		if(cascadeMenu!=null){
			MenuItem[]  submenus = cascadeMenu.getItems();
			for (int i = 0; i < submenus.length; i++) {
				unselectAllSubMenus(submenus[i].getMenu());
				submenus[i].setSelection(false);
			}
		}
	}
}
