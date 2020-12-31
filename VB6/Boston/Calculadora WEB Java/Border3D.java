/**
*File:  Border3D.java
*
*Description:  A utility class useful for creating components with
*              three dimensional borders.
**/


import java.awt.*;

public class Border3D {
    
    /*****************************************************************
    * Border Types:
    *
    * Static variables that allows control of how the the borders is 
    * displayed.
    *
    * EXCLUDE_TOP_BORDER:      Use to create components that lack a     
    *						   top border.
    *
    * EXCLUDE_BOTTOM_BORDER:   Use to create components that lack a
    *						   bottom border.
    *
    * FULL_BORDER:			   Default type.  Creates a component with
    *						   a full border.
    ******************************************************************/
    
    static final int EXCLUDE_TOP_BORDER = 1,       
                     EXCLUDE_BOTTOM_BORDER = 2,
                     FULL_BORDER = 3;
    
    //used to store border types             
    private int type;
    
    //used to store the component 
    private Component comp;
          	

public Border3D( Component comp ) { this( comp, FULL_BORDER ); }

public Border3D( Component comp, int type )
{
   this.comp = comp;
   this.type = type;	
}

/*
* Method:      draw3DBorder
* 
* Description: Draws a 3D border around components perimeter.  
*
*/

public void draw3DBorder( Graphics g ) 
{
   draw3DBorder( g, comp.size() );	
}

public void draw3DBorder( Graphics g, Dimension area ) {
   draw3DBorder( g, 0, 0, area.width, area.height );
}

public void draw3DBorder( Graphics g, int x, int y, int width, int height ) {

   //Draws top part of border		                             
   if ( type == EXCLUDE_BOTTOM_BORDER || type == FULL_BORDER ) 
   {
        g.setColor( comp.getBackground().darker() );
      	g.drawLine( x, y, width, y );
      		 
      	g.setColor( Color.white );
      	g.drawLine( x+1, y+1, width-1, y+1 );    		      	           	     
   }
   
   //Draws bottom portion of border     
   if ( type == EXCLUDE_TOP_BORDER || type == FULL_BORDER ) 
   {            	
      	g.setColor( comp.getBackground().brighter() );     	          	
      	g.drawLine( x, height, width, height);     	      	    	
      	g.setColor( comp.getBackground().darker() );          	    	
      	g.drawLine( x+1, height-1, width-1, height -1);
   }
   
   //draws the sides of the border   
   g.setColor( comp.getBackground().brighter() );
   g.drawLine( x+1, y+1, x+1, height - 1 );
   g.drawLine( width, y, width, height );
      
   g.setColor( comp.getBackground().darker() );
   g.drawLine( x, y, x, height ); 
   g.drawLine( width-1, y+1, width-1, height - 1);
        
}
	
}