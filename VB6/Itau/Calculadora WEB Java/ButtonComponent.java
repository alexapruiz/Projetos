/**
* File:  		ButtonComponent.java
* Description:  A Button component that feature's three a dimensional
*				similar to swing components.
*/

import java.lang.*;
import java.util.*;
import java.awt.*;
import java.awt.event.*;

public class ButtonComponent extends Component {

    private String label;                 //the button's label
    private boolean pressed = false;	  //used to determine button presses
    private ActionListener action;		  
    private boolean mouseOver = false;    //used to determine mouse overs 
    private Border3D border;			  //to draw a 3D border around button
  
public ButtonComponent(String label)
{
    this.label = label;
    border = new Border3D( this );    
    enableEvents(AWTEvent.MOUSE_EVENT_MASK);
}
  
    
/**
* Returns the preferred size of the button. This method is called automatically
* when the component is painted.
*
*/
public Dimension getPreferredSize() 
{
    Font f = getFont();
    if(f != null) {
       FontMetrics fm = getFontMetrics(getFont());
       return new Dimension(fm.stringWidth(label) + 10, fm.getHeight() + 5);
    } 
    else {
       return new Dimension(25, 25);
    }
}
  
/**
* Returns the minimum size of the button. 
*/
public Dimension getMinimumSize() { return new Dimension(20, 20); }
   
/**
* Detects mouse events on the component.
*/
public void processMouseEvent(MouseEvent e) 
{  
    switch(e.getID()) {
    	
      case MouseEvent.MOUSE_PRESSED:
        
        //Create a new actionevent and pass to the action listener -
        //( the Button Handler )
          
        if ( action != null ) {            
             ActionEvent event = new ActionEvent( this, e.getID(), label );
             action.actionPerformed( event );
        }
        
        pressed = true;                    
        
        //Invoke the repaint method which draws the button to appear pressed    
        repaint(); 
        break;
        
      case MouseEvent.MOUSE_RELEASED:
        //When repaint is invoked, component color is returned to normal   
        if(pressed == true) {
           pressed = false;
           repaint();
         }
         break;
         
      case MouseEvent.MOUSE_ENTERED:
         //repaint method lighten's component's color to give a hover effect 
         mouseOver = true;
         repaint();
         break;
         
      case MouseEvent.MOUSE_EXITED:
         
         //cancel hover effect
         mouseOver = false;
         if(pressed == true) pressed = false;               
            
         repaint();
         break;
       }
       
    super.processMouseEvent(e);
}

//Add's action listener   
public void addActionListener( ActionListener a )
{
	action = a;    //this is the ButtonHandler
}
 
//returns the buttons label  
public String getLabel() { return label; }

/*
* repaints the background
*/
public void paint(Graphics g) {
    int width = getSize().width - 1;
    int height = getSize().height - 1;
      
    //set background darker to give a pressed effect
    if(pressed) {
       g.setColor(getBackground().darker().darker());
    } 
    //set background lighter to give hover effect
    else if ( mouseOver ) {
       g.setColor( getBackground().brighter() ); 
    
    //set background to normal    
    } else {
       g.setColor(getBackground());
    }
    
    //fill background                          
    g.fillRect(0, 0, width, height );  
    
    //draws 3D border
    border.draw3DBorder(g, 0, 0, width, height );
    
    //center and draw the button's label         		    
    Font f = getFont();
    if(f != null) {
       FontMetrics fm = getFontMetrics(getFont());
       g.setColor(getForeground());
       g.drawString(label,
                    width/2 - fm.stringWidth(label)/2,
                    height/2 + fm.getMaxDescent());
    }
}

}
