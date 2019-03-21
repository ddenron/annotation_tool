/**
 * 
 */
package de.tudresden.annotator.main;

import org.eclipse.swt.SWT;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Dialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Event;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Listener;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Text;


/**
 * @author Elvis Koci
 *
 */
public class StringInputDialog extends Dialog {
	
  String value;

  /**
   * @param parent
   */
  public StringInputDialog(Shell parent) {
    super(parent);
  }
  
  
  /**
   * @param parent
   * @param style
   */
  public StringInputDialog(Shell parent, int style) {
    super(parent, style);
  }
  
  
  /**
   * Makes the dialog visible.
   * 
   * @return
   */
  public String open(String title, String dscText) {
    Shell parent = getParent();
    final Shell shell =
      new Shell(parent, SWT.CLOSE | SWT.TITLE | SWT.BORDER | SWT.APPLICATION_MODAL);
    shell.setText(title);

    shell.setLayout(new GridLayout(3, false));

    Label label = new Label(shell, SWT.NULL);
    label.setText(dscText);
    

    final Text text = new Text(shell, SWT.SINGLE | SWT.BORDER );
    GridData inputFieldLayout = new GridData(80, 25);
    inputFieldLayout.grabExcessHorizontalSpace = true;
    inputFieldLayout.horizontalAlignment = GridData.FILL;
    text.setLayoutData(inputFieldLayout);

    final Button buttonOK = new Button(shell, SWT.PUSH);
    buttonOK.setText(" GO ");
    GridData buttonLayout = new GridData(50, 30);
    buttonOK.setLayoutData(buttonLayout);
    
    shell.setDefaultButton(buttonOK);
    
    
    text.addListener(SWT.Modify, new Listener() {
      public void handleEvent(Event event) {
        try {
          value = text.getText();
          buttonOK.setEnabled(true);
        } catch (Exception e) {
          buttonOK.setEnabled(false);
        }
      }
    });

    
    buttonOK.addListener(SWT.Selection, new Listener() {
      public void handleEvent(Event event) {
    	if (value==null || value==""){
    		value = "!&@!*";
    	}
    	
    	if (shell != null){
  			shell.dispose();
  		}
      }
    });
    
    
    shell.addListener(SWT.Close, new Listener() {
        public void handleEvent(Event event) {
        	value = null;
            if (shell != null){
       			shell.dispose();
       		}
        }
      });
    
    
    shell.addListener(SWT.Traverse, new Listener() {
      public void handleEvent(Event event) {
        if(event.detail == SWT.TRAVERSE_ESCAPE)
          event.doit = false;
      }
    });

    text.setText("");
    shell.pack();
    shell.open();

    Display display = parent.getDisplay();
    while (!shell.isDisposed()) {
      if (!display.readAndDispatch())
        display.sleep();
    }

    return value;
  }
  
}
