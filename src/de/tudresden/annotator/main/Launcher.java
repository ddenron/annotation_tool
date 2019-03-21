/**
 * 
 */
package de.tudresden.annotator.main;


import java.io.File;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.eclipse.swt.SWT;
import org.eclipse.swt.SWTException;
import org.eclipse.swt.custom.SashForm;
import org.eclipse.swt.graphics.Color;
import org.eclipse.swt.graphics.Image;
import org.eclipse.swt.graphics.Point;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.ole.win32.OLE;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleControlSite;
import org.eclipse.swt.ole.win32.OleFrame;
import org.eclipse.swt.ole.win32.OleListener;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.ToolTip;

import de.tudresden.annotator.annotations.WorkbookAnnotation;
import de.tudresden.annotator.annotations.WorksheetAnnotation;
import de.tudresden.annotator.annotations.utils.AnnotationHandler;
import de.tudresden.annotator.annotations.utils.RangeAnnotationsSheet;
import de.tudresden.annotator.oleutils.ApplicationUtils;
import de.tudresden.annotator.oleutils.CommandBarUtils;
import de.tudresden.annotator.oleutils.RangeUtils;
import de.tudresden.annotator.oleutils.WindowUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;
import de.tudresden.annotator.oleutils.WorksheetUtils;

/**
 * 
 * @author Elvis Koci
 */
public class Launcher {
	
	// GUID Event Sink
	private final String IID_AppEvents = "{00024413-0000-0000-C000-000000000046}";
	// Event IDs
	private final int SheetSelectionChange = 0x00000616;
	private final int SheetActivate        = 0x00000619;
	private final int SheetBeforeRightClick  = 0x00000618;
	public final static int WindowActivate         = 0x00000614;
    public final static int WindowDeactivate       = 0x00000615;
		
	private final Display display = new Display();
	private final Shell shell = new Shell(display);
	private ToolTip tooltip = new ToolTip(shell, SWT.NONE);
	
	private OleFrame oleFrame;
	private OleControlSite controlSite;
	
	private BarMenu menuBar;
	
	private OleAutomation embeddedWorkbook;
	private String embeddedWorkbookName;
	
	private OleAutomation embeddedWindow;
	
	private String directoryPath;
	private String fileName;
	
	private String activeWorksheetName;
	private int activeWorksheetIndex;
	
	private String currentSelection[];
	
	private SashForm mainSash;
	private SashForm rightSash;
	private Composite folderExplorerPanel;
	private Composite rightPanel;
	private Composite annotationsPanel;
	private Composite excelPanel;
	
	protected long lastShellFocusTime = -1;
	protected boolean justClickedCompositeSite = false;
	
	
	private static Launcher instance = null;
	private Launcher(){}
	
	public static Launcher getInstance() {
		if(instance == null) {
			instance = new Launcher();
		}
		return instance;  
	}
	
	private static final Logger logger = LogManager.getLogger(Launcher.class.getName());
	
	/**
	 * Create the window that will serve as the main Graphical User Interface (GUI) for user interaction
	 * @param shell
	 */
	private void buildGUIWindow(){

		
		// shell properties
		this.shell.setText("Annotator");
	    this.shell.setLayout(new FillLayout());
	    this.shell.setSize(1500, 650);
	    // add listener for the close event ( user clicks the close button (X) )
	    this.shell.addListener(SWT.Close, GUIListeners.createCloseApplicationEventListener());
	    
		// the colors to use for the GUI
		// Color white = new Color (Display.getCurrent(), 255, 255, 255);
		Color lightGreyShade = new Color (Display.getCurrent(), 247, 247, 247);
		
	    // split shell in two horizontal panels 
	    mainSash = new SashForm(shell, SWT.HORIZONTAL);
		mainSash.setLayout(new FillLayout());
	
		// the left panel will contain the folder explorer. That is a tree structure of files and folders.
		folderExplorerPanel = new Composite(mainSash, SWT.BORDER );
		folderExplorerPanel.setLayout(new FillLayout());
		folderExplorerPanel.setVisible(true);
		//leftPanel.setEnabled(false);
		folderExplorerPanel.setBackground(lightGreyShade);

				
		// the right panel will be subdivided into two more panels
		rightPanel = new Composite(mainSash, SWT.BORDER);
		rightPanel.setLayout(new FillLayout());
		
		mainSash.setWeights(new int[] {0, 100});
			
		// sub split the right panel 
	    rightSash = new SashForm(rightPanel, SWT.VERTICAL);
		rightSash.setLayout(new FillLayout());
		
		// create the panel that will embed the excel application
		excelPanel =  new Composite(rightSash, SWT.BORDER);
		FillLayout excelPanelLayout = new FillLayout();
		excelPanelLayout.marginHeight = 4;
		excelPanelLayout.marginWidth = 4;
		excelPanel.setLayout(excelPanelLayout);
		
		
		// create the panel that will display the applied annotations for the current file
		annotationsPanel =  new Composite(rightSash, SWT.BORDER );
		annotationsPanel.setLayout(new FillLayout());
		annotationsPanel.setVisible(false);
		annotationsPanel.setBackground(lightGreyShade);
		
		rightSash.setWeights(new int[] {100,0});
		
		//create oleframe
		setOleFrame(new OleFrame(excelPanel, SWT.NONE));
		getOleFrame().setBackground(lightGreyShade);

		// add a bar menu
	    BarMenu  oleFrameMenuBar = new BarMenu(getOleFrame().getShell());
	    getOleFrame().setFileMenus(oleFrameMenuBar.getMenuItems());
	    
	    this.setMenuBar(oleFrameMenuBar);    
	}
	
	
	/**
	 * Embed the specified excel file to the application window
	 * @param excelFile the excel file to embed
	 */
	protected void embedExcelFile(File excelFile){	
		
		try {
			// setControlSite(new OleControlSite(getOleFrame(), SWT.NONE, excelFile));
			setControlSite(new OleControlSite(getOleFrame(), SWT.NONE, "Excel.Sheet", excelFile));
		} catch (IllegalArgumentException iaEx) {
			
			logger.fatal("IllegalArgumentException on embed file/creation of control site", iaEx);
			
			int style = SWT.ICON_ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("ERROR: Control site could not be created. Illegal argument exception was thrown.");
			message.open();
			Launcher.getInstance().disposeShell();

		} catch (SWTException swtEx) {
		
			logger.fatal("SWTException on embed file/creation of control site", swtEx);
			
			int style = SWT.ICON_ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("ERROR: could not embedd the excel workbook. SWT Exception was thrown");
			message.open();
			Launcher.getInstance().disposeShell();
			
		} catch (Exception ex) {
			
			logger.fatal("Exception on embed file/creation of control site", ex);
			
			int style = SWT.ICON_ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("Something went wrong!!! Ensure that you have a version of Microsoft Office Excel"
					+ " installed in your machine. Also, check that the file is not corrupted or wrong format.");
			message.open();
			Launcher.getInstance().disposeShell();
		}
		
		// activate and display excel workbook
		getControlSite().doVerb(OLE.OLEIVERB_INPLACEACTIVATE);
			
		// get excel application as OLE automation object
	    OleAutomation application = ApplicationUtils.getApplicationAutomation(getControlSite());
	    if(application==null){
	    	int style = SWT.ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("Something went wrong!!! Please take the following actions.\n\n"
					+ "1. Check if an instance of this application is already running.\n\n"
					+ "2. Ensure that the excel file you are trying to open it is not used by another application.\n\n"
					+ "3. If there is another excel file oppened outiside of this application, ensure that "
					+ "there are no pending windows or message boxes asking for the user input.\n\n"
					+ "4. Open task manager and check if there is any excel process running in the background. "
					+ "If there is such process, end it.");		
			message.open();
			Launcher.getInstance().disposeControlSite();
			Launcher.getInstance().disposeShell();
			return;
	    }
	        
	    // get active workbook, the one that is embedded in this application
	    OleAutomation workbook = ApplicationUtils.getActiveWorkbookAutomation(application);
	    setEmbeddedWorkbook(workbook);
	    
	    // get the name of workbook for future reference. 
	    // the name of the workbook might be different from the excel file name. 
	    String workbookName = WorkbookUtils.getWorkbookName(workbook);
	    setEmbeddedWorkbookName(workbookName);
	    
	    // get the active excel window 
	    OleAutomation window = ApplicationUtils.getActiveWindowAutomation(application);
	    setEmbeddedWindow(window);
	    
	    // get the active sheet automation, i.e. the one that is on top of the other worksheet
	    OleAutomation worksheet = WorkbookUtils.getActiveWorksheetAutomation(workbook);
	    
	    // get and store the name and index for the sheet that is "active"
	    String sheetName = WorksheetUtils.getWorksheetName(worksheet);
	    setActiveWorksheetName(sheetName);
	    int sheetIndex = WorksheetUtils.getWorksheetIndex(worksheet);
	    setActiveWorksheetIndex(sheetIndex);
	    worksheet.dispose();
	    
	    // save the directory and fileName for future reference
		setDirectoryPath(excelFile.getParent());
		setFileName(excelFile.getName());
		
		// update display window
	    Color green2 = new Color (Display.getCurrent(), 154, 200, 122);
	    this.excelPanel.setBackground(green2);
	    getShell().setText("Annotator - "+excelFile.getName());
	    
	    // deactivateControlSite();
	    
		// prepare the display for the annotation process
		setUpApplicationDisplay(application);
		setUpWorkbook(workbook);
	}
	
	
	/**
	 * set up (prepare) the excel application display for the annotation process
	 * @param application an OleAutomation object that provides access the functionalities of the Excel Application
	 */
	protected void setUpApplicationDisplay(OleAutomation application){
		    
	    // add event listeners
	    OleListener sheetSelectionListener = GUIListeners.createSheetSelectionEventListener();
        getControlSite().addEventListener(application, IID_AppEvents, SheetSelectionChange, sheetSelectionListener);
        
        OleListener sheetActivationListener = GUIListeners.createSheetActivationEventListener();
        getControlSite().addEventListener(application, IID_AppEvents, SheetActivate, sheetActivationListener);
        
        OleListener sheetBeforeRightClickListener = GUIListeners.createSheetBeforeRightClickEventListener();
        getControlSite().addEventListener(application, IID_AppEvents, SheetBeforeRightClick, sheetBeforeRightClickListener);	
        
        //OleListener windowActivateListener = GUIListeners.windowActivateEventListener();
        //getControlSite().addEventListener(application, IID_AppEvents, WindowActivate, windowActivateListener);
        
        //OleListener windowDeactivateListener = GUIListeners.windowDeactivateEventListener();
        //getControlSite().addEventListener(application, IID_AppEvents, WindowDeactivate, windowDeactivateListener);
 
        
		// minimize ribbon.	
	    ApplicationUtils.hideRibbon(application);	
	    
	    CommandBarUtils.setEnabledForCommandBar(application, "Ply", false);
	        
	    // show status bar
	    // ApplicationUtils.setVisibilityStatusBar(application, true);
	    
	    // hide menu on right click of user at a worksheet tab
	    CommandBarUtils.setEnabledForCommandBar(application, "Ply", false);
	    
	    // hide menu on right click of user on a cell
	    CommandBarUtils.setEnabledForCommandBar(application, "Cell", false);
	    
	    // hide menu on right click of user on rows
	    CommandBarUtils.setEnabledForCommandBar(application, "Row", false);
	    
	    // hide menu on right click of user on columns
	    CommandBarUtils.setEnabledForCommandBar(application, "Column", false);
	}
	
	
	/**
	 * set up the workbook for the annotation process
	 * @param workbook an OleAutomation object that provides access the functionalities of the embedded workbook
	 */
	protected void setUpWorkbook(OleAutomation workbook){		
		
		// show the annotation_data_sheet if it exists
	    OleAutomation  rangeAnnotationDataSheet = 
			WorkbookUtils.getWorksheetAutomationByName(workbook, RangeAnnotationsSheet.getName()); 
	    
	    if(rangeAnnotationDataSheet!=null){
	    	WorksheetUtils.setWorksheetVisibility(rangeAnnotationDataSheet, true);
	    }
	    
		// protect the structure of the workbook if it is not yet protected
		boolean isProtected = WorkbookUtils.protectWorkbook(workbook, true, false);		
		if(!isProtected){
			int style = SWT.ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("ERROR: Could not protect the workbook. Operation failed!");
			message.open();
			
			quitApplication();
		}
	    
		// protect all the worksheet in the embedded workbook 
		boolean areProtected = WorkbookUtils.protectAllWorksheets(workbook);
		if(!areProtected){
			int style = SWT.ERROR;
			MessageBox message = Launcher.getInstance().createMessageBox(style);
			message.setMessage("ERROR: Could not protect one or more sheets. Operation failed!");
			message.open();
			
			quitApplication();
		}		
	}
	
	
	/**
	 * prepare the sheet display for annotation 		
	 */
	protected void updateActiveSheetDisplay(){
		
		WorkbookAnnotation wa = AnnotationHandler.getWorkbookAnnotation();		
		WorksheetAnnotation sa= wa.getWorksheetAnnotations().get(this.activeWorksheetName);
		
		// set up display for activated sheet
		if(wa!=null && sa!=null){	
			
			// Move horizontal scroll to left most column
			WindowUtils.setScrollColumn(this.embeddedWindow, 1);
			// Move vertical scroll to the top row
			WindowUtils.setScrollRow(this.embeddedWindow, 1);
			
			// display tooltip if one of the following cases holds
			this.tooltip.setVisible(false);
			
			if(wa.isCompleted()){
				
				recreateToolTipWithStyle(SWT.ICON_INFORMATION);
				this.tooltip.setText("Information");
				this.tooltip.setMessage("This workbook was marked as \"Completed\".");
				placeToolTipOnExcelPanel();
				this.tooltip.setVisible(true);			
				
			}else if(wa.isNotApplicable()){
				
				recreateToolTipWithStyle(SWT.ICON_INFORMATION);
				this.tooltip.setText("Information");
				this.tooltip.setMessage("This workbook was marked as \"Not Applicable\".");
				placeToolTipOnExcelPanel();
				this.tooltip.setVisible(true);	
				
			}else if(sa.isCompleted()){
				
				recreateToolTipWithStyle(SWT.ICON_INFORMATION);
				this.tooltip.setText("Information");
				this.tooltip.setMessage("This sheet was marked as \"Completed\".");
				placeToolTipOnExcelPanel();
				this.tooltip.setVisible(true);
			
			}else if(sa.isNotApplicable()){
			
				recreateToolTipWithStyle(SWT.ICON_INFORMATION);
				this.tooltip.setText("Information");
				this.tooltip.setMessage("This sheet was marked as \"Not Applicable\".");
				placeToolTipOnExcelPanel();
				this.tooltip.setVisible(true);
			
			}else if(!sa.isCompleted() && !sa.isNotApplicable()){
				
				OleAutomation sheetAutomation = WorkbookUtils.getWorksheetAutomationByName(
						embeddedWorkbook, activeWorksheetName);
				
				WorksheetUtils.unprotectWorksheet(sheetAutomation);
				OleAutomation usedCells = WorksheetUtils.getUsedRange(sheetAutomation);
				sheetAutomation.dispose();
				
				if(usedCells!=null){
					OleAutomation formulaCells = RangeUtils.getSpecialCells(usedCells, -4123);
					if(formulaCells!=null){
						if(RangeUtils.count(formulaCells)>0){
							recreateToolTipWithStyle(SWT.ICON_WARNING);
							this.tooltip.setText("Warning");
							this.tooltip.setMessage("This sheet contains formulas!");
							placeToolTipOnExcelPanel();
							// this.tooltip.setAutoHide(false);
							this.tooltip.setVisible(true);						
						}
					}				
					usedCells.dispose();
				}
				
				OleAutomation sheetAutoAfterFormulaCheck = WorkbookUtils.getWorksheetAutomationByName(
						embeddedWorkbook, activeWorksheetName);
				WorksheetUtils.protectWorksheet(sheetAutoAfterFormulaCheck);
				sheetAutoAfterFormulaCheck.dispose();
			}		
		}
	}
	
	/**
	 * Set background color to the excel panel
	 * @param color an object that represents the color to set
	 */
	protected void setColorToExcelPanel(Color color){
		excelPanel.setBackground(color);
		
	}
	
	/**
	 * 
	 * @param enabled
	 */
	protected void setExcelPanelEnabled(boolean enabled){
		excelPanel.setEnabled(enabled);
	}
	
	/**
	 * @return the display
	 */
	private Display getDisplay() {
		return display;
	}

	/**
	 * 
	 * @param millis
	 * @param runnable
	 */
	protected void displayTimerExec(int millis, Runnable runnable) {
		this.display.timerExec(millis, runnable);		
	}
	
	/**
	 * 
	 * @param runnable
	 */
	protected void displaySyncExec(Runnable runnable) {
		this.display.syncExec(runnable);
	}
	
	/**
	 * @return the shell
	 */
	private Shell getShell() {
		return shell;
	}
		
	/**
	 * Dispose shell
	 */
	protected void disposeShell() {
		if (this.shell != null){
			this.shell.dispose();
		}
	}
	
	/**
	 * give the keyboard focus to the shell
	 */
	protected void setFocusToShell(){
		this.shell.setFocus();

		Color lightRed= new Color(Display.getCurrent(), 243, 121, 121);
		excelPanel.setBackground(lightRed);
	}
	
	/**
	 * force the keyboard focus to the shell
	 */
	protected void forceFocusToShell(){
		this.shell.forceFocus();

		Color lightRed= new Color(Display.getCurrent(), 243, 121, 121);
		excelPanel.setBackground(lightRed);
	}
	
	/**
	 * Set capture for shell
	 * @param capture true to capture all mouse events, false to release
	 */
	protected void setShellCapture(boolean capture){
		this.shell.setCapture(capture);		
	}
	
	/**
	 * Check if the shell has the focus
	 * @return true if the shell has the focus, false otherwise
	 */
	protected boolean isShellFocusControl(){
		return this.shell.isFocusControl();
	}
	
	/**
	 * @return the tooltip
	 */
	protected ToolTip getTooltip() {
		return tooltip;
	}
	
	/**
	 * Replace the existing tooltip with one that has the specified style
	 * 
	 * @param style one of the SWT values relevant to tooltip style
	 */
	protected void recreateToolTipWithStyle(int style){	
		this.tooltip = new ToolTip(shell, style);
	}
	
	/**
	 * Place tooltip on top of the embedded excel spreadsheet
	 */
	public void placeToolTipOnExcelPanel(){
		
		Point shellLoc = shell.getLocation();
		Point csLoc = controlSite.getLocation();
		
		int x = shellLoc.x + csLoc.x + 20;
		int y = shellLoc.y + csLoc.y + 20;
		
		this.tooltip.setLocation(x, y);
	}
		
	/**
	 * Get OleFrame
	 * @return
	 */
	private OleFrame getOleFrame() {
		return oleFrame;
	}		
	
	/**
	 * Set OleFrame
	 * @param oleFrame
	 */
	private void setOleFrame(OleFrame oleFrame) {
		this.oleFrame = oleFrame;
	}
		
	
	/**
	 * Get OleControlSite
	 * @return
	 */
	private OleControlSite getControlSite() {
		return controlSite;
	}
	
	/**
	 * Set OleControlSite
	 * @param controlSite
	 */
	private void setControlSite(OleControlSite controlSite) {
		this.controlSite = controlSite;
	}
	
	/**
	 * give the keyboard focus to the controlsite
	 */
	protected void setFocusToControlSite(){
		if(controlSite!=null){
			this.controlSite.setFocus();
		}
		
		Color green2 = new Color (Display.getCurrent(), 154, 200, 122);
		excelPanel.setBackground(green2);
	}
	
	/**
	 * force the keyboard focus to the controlsite
	 */
	protected void forceFocusToControlSite(){
		if(controlSite!=null)
			this.controlSite.forceFocus();
		
		Color green2 = new Color (Display.getCurrent(), 154, 200, 122);
		excelPanel.setBackground(green2);
	}
	
	/**
	 * Check if the control site has the focus
	 * @return true if the control site has the focus, false otherwise
	 */
	protected boolean isControlSiteFocusControl(){
		return this.controlSite.isFocusControl();
	}
	
	
	/**
	 * Check if the control site is enabled
	 * @return true if the control site is enabled, false otherwise
	 */
	protected boolean isControlSiteEnabled(){
		return this.controlSite.isEnabled();
	}
	
	
	/**
	 * Requests that the control site perform an action
	 * @param verb the operation that is requested. This is one of the OLE.OLEIVERB_ values
	 * @return an HRESULT value indicating the success of the operation request; OLE.S_OK indicates success
	 */
	protected int doVerbControlSite(int verb){
		return this.controlSite.doVerb(verb);
	}
	
	
	/**
	 * Deactivate control site
	 */
	protected void deactivateControlSite(){
		this.controlSite.deactivateInPlaceClient();
	}
	
	
	/**
	 * Dispose control site 
	 */
	protected void disposeControlSite() {
		if (controlSite != null){
			controlSite.dispose();
		}
		controlSite = null;
	}
	
	/**
	 * Check if control site is null
	 * @return true if it is null, false otherwise
	 */
	protected boolean isControlSiteNull(){
		return controlSite == null;
	}
	
	/**
	 * Check if control site is dirty
	 * @return true if it is dirty, false otherwise
	 */
	protected boolean isControlSiteDirty(){
		return controlSite.isDirty();
	}
	
	/**
	 * Set visibility for control site
	 * @param visible true to set control site visible, false to hide it
	 */
	protected void setVisibilityForControlSite(boolean visible){	
		controlSite.setVisible(visible);
	}
	
	/**
	 * Set capture for control site
	 * @param capture true to capture all mouse events, false to release
	 */
	protected void setControlSiteCapture(boolean capture){
		this.controlSite.setCapture(capture);		
	}

	
	/**
	 * Set if control site is enabled
	 * @param true to enable control site, false otherwise
	 */
	protected void setEnabledControlSite(boolean isEnable){
		this.controlSite.setEnabled(isEnable);		
	}
	
	
	/**
	 * @return the embeddedWorkbook
	 */
	protected OleAutomation getEmbeddedWorkbook() {
		return embeddedWorkbook;
	}

	/**
	 * @param embeddedWorkbook the embeddedWorkbook to set
	 */
	protected void setEmbeddedWorkbook(OleAutomation embeddedWorkbook) {
		this.embeddedWorkbook = embeddedWorkbook;
	}
	
	/**
	 * @return the activeWorkbookName
	 */
	protected String getEmbeddedWorkbookName() {
		return embeddedWorkbookName;
	}
	
	/**
	 * @param activeWorkbookName the activeWorkbookName to set
	 */
	protected void setEmbeddedWorkbookName(String activeWorkbookName) {
		this.embeddedWorkbookName = activeWorkbookName;
	}
		
	/**
	 * @return the embeddedWindow
	 */
	protected OleAutomation getEmbeddedWindow() {
		return embeddedWindow;
	}

	/**
	 * @param embeddedWindow the embeddedWindow to set
	 */
	protected void setEmbeddedWindow(OleAutomation embeddedWindow) {
		this.embeddedWindow = embeddedWindow;
	}

	/**
	 * @return the filePath
	 */
	protected String getDirectoryPath() {
		return directoryPath;
	}

	/**
	 * @param filePath the filePath to set
	 */
	protected void setDirectoryPath(String filePath) {
		this.directoryPath = filePath;
	}

	/**
	 * @return the fileName
	 */
	protected String getFileName() {
		return fileName;
	}

	/**
	 * @param fileName the fileName to set
	 */
	protected void setFileName(String fileName) {
		this.fileName = fileName;
	}

	/**
	 * @return the activeWorksheetName
	 */
	protected String getActiveWorksheetName() {
		return activeWorksheetName;
	}

	/**
	 * @param activeWorksheetName the activeWorksheetName to set
	 */
	protected void setActiveWorksheetName(String activeWorksheetName) {
		this.activeWorksheetName = activeWorksheetName;
	}

	/**
	 * @return the activeWorksheetIndex
	 */
	protected int getActiveWorksheetIndex() {
		return activeWorksheetIndex;
	}

	/**
	 * @param activeWorksheetIndex the activeWorksheetIndex to set
	 */
	protected void setActiveWorksheetIndex(int activeWorksheetIndex) {
		this.activeWorksheetIndex = activeWorksheetIndex;
	}
	
	/**
	 * @return the currentSelection
	 */
	protected String[] getCurrentSelection() {
		return currentSelection;
	}

	/**
	 * @param currentSelection the currentSelection to set
	 */
	protected void setCurrentSelection(String[] currentSelection) {
		this.currentSelection = currentSelection;
	}

	/**
	 * @return the menuBar
	 */
	protected BarMenu getMenuBar() {
		return menuBar;
	}

	/**
	 * @param menuBar the menuBar to set
	 */
	protected void setMenuBar(BarMenu menuBar) {
		this.menuBar = menuBar;
	}

	/**
	 * Create message box using the "main" window (this class) shell 
	 * @param style one of the relevant SWT constants or their combination
	 * @return a MessageBox object
	 */
	public MessageBox createMessageBox(int style){
		return new MessageBox(this.shell, style);
	}
	
	/**
	 * Create a file dialog using the main shell
	 * @param style one of the relevant SWT constants or their combination
	 * @return FileDialog object
	 */
	public FileDialog createFileDialog(int style){
		return  new FileDialog(this.shell, SWT.OPEN);
	}
	
	
	/**
	 * Create an input dialog using the main shell
	 * @return
	 */
	public StringInputDialog createInputRangeDialog(){		
		return new StringInputDialog(this.shell);
	}
	
	
	/**
	 * Create an input dialog using the main shell
	 * @return
	 */
	public TextDialog createTextDialog(){		
		return new TextDialog(this.shell);
	}
	
	
	/**
	 * Create an image using the main display as device
	 * @param fileName the name of the image file 
	 * @return an object that represents an SWT image
	 */
	public Image createImage(String fileName){
		return new Image(this.display, fileName);
	}
	
	/**
	 * Quit this application
	 * This method is meant to be used for abnormal exit
	 * when exceptions occur. 
	 */
	public void  quitApplication(){
		WorkbookUtils.closeEmbeddedWorkbook(this.embeddedWorkbook, false);
		disposeControlSite();
		disposeShell();
		// getDisplay().dispose();
		// System.exit(1);
	}
		
	/**
	 * @param args
	 */
	public static void main(String[] args) {

		Launcher main = Launcher.getInstance(); 
	
	    main.buildGUIWindow();
	    
  		main.getShell().open();
  			
  		try{
	  	    while (!main.getShell().isDisposed ()) {
	  	        if (!main.getDisplay().readAndDispatch ()) main.getDisplay().sleep();
	  	    }
	  	    
		    main.getDisplay().dispose();
  		}catch (Exception ex){
  			logger.fatal("Generic exception thrown in main", ex);
  		}
	}		
}
