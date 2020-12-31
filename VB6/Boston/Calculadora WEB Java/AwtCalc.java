/**
*File:  		AwtCalc.java
*Version:		1.1
*Description:   A simple calculator that uses AWT component's.  
*              
**/

import java.awt.*;
import java.awt.event.*;
import java.awt.Graphics;

public class AwtCalc extends Panel
{
    //Labels for the number panel of the calculator
    private String[] numPanelText = { " 1 ", " 2 ", " 3 ",
    								  " 4 ", " 5 ", " 6 ",
    								  " 7 ", " 8 ", " 9 ",
    								  Operator.CLEAR  , " 0 ", Operator.DOT };
    //Labels for the operator panel of the calculator								  				  
    private String[] operPanelText = {  Operator.ADD, Operator.SUBTRACT, 
    									Operator.MULTIPLY, Operator.DIVIDE, 
    									Operator.POW, Operator.SQRT, 
    									Operator.NEGATE, Operator.EQUALS };
    									
    	
    private Panel numButtonPanel;     //used to hold the number buttons
    private Panel operButtonPanel;	  //used to hold the operator buttons
    private Panel3D displayPanel;	  //used for the calculator's display 
    private ButtonHandler handler;	  //action listener for the buttons							 
    private CalcDisplay display;      //displays the output     
    private Font buttonfont;		  
    
public AwtCalc() 
{    
    //Initialize   
    buttonfont = new Font( "Courier", Font.PLAIN, 13 );
    setLayout( new BorderLayout() );
    setBackground( new Color( 212, 208, 200 ) );
    Panel3D mainPanel = new Panel3D( Border3D.EXCLUDE_TOP_BORDER );
        
        	
    numButtonPanel = new Panel( new GridLayout(4,3, 1, 1) );
    operButtonPanel = new Panel( new GridLayout(4, 2, 1, 1) );
    displayPanel = new Panel3D( Border3D.EXCLUDE_BOTTOM_BORDER );
    display = new CalcDisplay( 192,26);
    handler = new ButtonHandler( display );
        
    displayPanel.add( display );
        
    mainPanel.add( createNumberPanel() );
    mainPanel.add( createOperPanel() );
        
    add( displayPanel, BorderLayout.NORTH );     
    add( mainPanel, BorderLayout.CENTER );                     
}
 
/*
*  Method: 		 createNumberPanel
*  Description:  contructs and returns the calculator's number panel
*/
   
private Panel createNumberPanel()
{
   if ( display != null ) {
      	
      ButtonComponent btn = null;
      	
   for ( int i = 0; i < numPanelText.length; i++ )
   {
      btn = new ButtonComponent( numPanelText[i] );   
      btn.addActionListener( handler );
      btn.setFont( buttonfont );
      numButtonPanel.add( btn );
   } 	   	
   }
	
return numButtonPanel;
}
 
/**
* Method: 		createOperPanel
* Description:	contructs and returns the calculator's number panel
**/   
private Panel createOperPanel()
{
    ButtonComponent btn = null;
    	
    for ( int i = 0; i < operPanelText.length; i++ )
    {
    	btn = new ButtonComponent( operPanelText[i] );
    	btn.setFont( buttonfont );   
    	btn.addActionListener( handler );  
    	operButtonPanel.add( btn );	
    }
    	
return operButtonPanel;
}
    
}


