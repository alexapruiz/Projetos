/*
*  File: 		 CalcDisplay.java
*  Description:  This class creates a display screan for AwtCalc 
*
*/

import java.awt.*;

class CalcDisplay extends Canvas {
	
   private String text;
   private Rectangle area;
   private Border3D border;
   
CalcDisplay( int w, int h)
{
	setSize( w , h  );
	
	//Determine the center of the component and create Rectangle object
	//which will be used to draw the white portion and the borders of the
	//display
	
	int x = ( size().width - w ) / 2;
	int y = ( size().height - h ) / 2 + 6;
	
	border = new Border3D( this );	
	area = new Rectangle( x, y, w, h );	  
			
   	setFont( new Font( "monospace", Font.PLAIN, 14 ) ); 	
   	setForeground( Color.black );
   	text = "0";  	
   
}


public Dimension getPreferredSize()
{
    return new Dimension( size().width, size().height );	
}

//set's text ( numbers ) to the display
void append( String s )
{
	//clear screen first before updating text
	clear();
	text = s;	
	repaint();
}

//clear's screen
void clear()
{
    text = ""; 
    repaint();	
}

//get's the text on the display
String getText() { return text; }

//paints all contents of the component
public void paint( Graphics g )
{
    drawDisplayScreen( g );
    drawOutput( g );				
}

//draw's the display
protected void drawDisplayScreen( Graphics g )
{
	g.setColor( Color.white );
	g.fillRect( area.x, area.y, area.width, area.height );
	border.draw3DBorder( g, area.x, area.y, area.width, area.height );
}

//draw's the text ( numbers ) on the screen
protected void drawOutput( Graphics g )
{
	g.setColor( Color.black );		
	g.drawString( text, area.x+3, area.y+15 );
}

}