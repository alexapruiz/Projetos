/*
*  File:   		AwtCalcApplet.java
*  Description: Allow's AwtCalc to be run as an applet.
*/

import java.applet.Applet;
import java.awt.*;

public class AwtCalcApplet extends Applet {
	
public void init()
{

setLayout( new BorderLayout() );
add( new AwtCalc(), BorderLayout.CENTER );
	
}	


}