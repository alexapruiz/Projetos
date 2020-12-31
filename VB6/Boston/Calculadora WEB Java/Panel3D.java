/*
* File:   		Panel3D.java
* Description:  Creates a panel with a three dimensional border
*/
import java.awt.*;

class Panel3D extends Panel {
	                
private Border3D border;     //used for drawing the border

//create a full border by default
public Panel3D() { this( Border3D.FULL_BORDER ); }

public Panel3D( int type )
{
  	border = new Border3D( this, type );
}
	    
public void paint( Graphics g )
{
	border.draw3DBorder( g );		
}

}