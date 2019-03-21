/**
 * 
 */
package de.tudresden.annotator.main;

import org.eclipse.swt.SWT;
import org.eclipse.swt.graphics.Rectangle;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Dialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Event;
import org.eclipse.swt.widgets.Listener;
import org.eclipse.swt.widgets.Monitor;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Text;


/**
 * @author Elvis Koci
 *
 */
public class TextDialog extends Dialog {
	
  /**
   * @param parent
   */
  public TextDialog(Shell parent) {
    super(parent);
  }
  
  
  /**
   * @param parent
   * @param style
   */
  public TextDialog(Shell parent, int style) {
    super(parent, style);
  }
  
 
  
  /**
   * Makes the dialog visible.
   * 
   * @return
   */
  public String open(String title, String value) {
    Shell parent = getParent();
    final Shell shell =
      new Shell(parent, SWT.CLOSE | SWT.TITLE | SWT.BORDER | SWT.APPLICATION_MODAL);
    shell.setText(title);
    
    Monitor primary = parent.getDisplay().getPrimaryMonitor();
    Rectangle bounds = primary.getBounds();
    int x = (bounds.x + bounds.width) / 2;
    int y = (bounds.y + bounds.height) / 2;
    //System.out.println("x:"+x+", y:"+y);
    shell.setLocation(x, y);
    
    shell.setLayout(new GridLayout(1, false));
    
    
    Text text = new Text(shell, SWT.MULTI | SWT.READ_ONLY | SWT.BORDER | SWT.V_SCROLL | SWT.H_SCROLL);
    text.setText(value);
    GridData textLayout = new GridData(310, 90);
    text.setLayoutData(textLayout);
    
    
    shell.addListener(SWT.Close, new Listener() {
        public void handleEvent(Event event) {
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
