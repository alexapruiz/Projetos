/**
* File:			AwtCalcApp.java
* Description:  Allows AwtCalc to be run as a stand-alone application.
**/

import java.awt.*;
import java.awt.event.*;

class AwtCalcApp {
	
public static void main( String args[] )
{
   Frame fr = new Frame();
   fr.setTitle( "Awt Calculator" );
   fr.setSize( 220, 175 );
   fr.setResizable( false );
   fr.add( new AwtCalc(), BorderLayout.CENTER );
   fr.setVisible( true );
   
   fr.addWindowListener( new WindowAdapter() {
   	
   	public void windowClosing( WindowEvent e ) {
   		
   		System.exit( 1 );
   		}
   	}
   	);	
}

}