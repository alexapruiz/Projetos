/*
* File:  		ButtonHandler.java
* Description:  Handle's what happen's when the button is pressed.
*/

import java.awt.event.*;
import java.awt.*;

class ButtonHandler implements ActionListener {


private CalcDisplay display;	//holds instance of calculator's display
private String lastOp;			//holds the last operator that was pressed
private String strVal;			//holds the string value of the number 
private double total;			//accumulator
private double number;			//used to store new number's

//flag used to determine if op is pressed for the first time
private boolean firsttime;		

//flag used to determine if an operator has been pressed
private boolean operatorPressed; 

ButtonHandler( CalcDisplay d )
{
   display = d;	
   firsttime = true;
   strVal = "";
}

// this method is called automatically when a button is pressed
public void actionPerformed( ActionEvent e )
{
	//store the instance of the button that was pressed
	ButtonComponent button = (ButtonComponent) e.getSource();
	
	//get the button's label
	String s = button.getLabel().trim();
	
	//determine if the button was a number or operator button
	if ( Character.isDigit( s.charAt(0) ) )		
		handleNumber( s );					
	else
	    calculate( s );	    	   		
}

//determine's whether the operator is unary or binary
//and does calculation's according to operator type
void calculate( String op )
{
  operatorPressed = true;
	
  //if it's the firsttime the button has been pressed after:
  //
  //		- the program first starts
  //		- the equal has been pressed
  //		- or the calculator has been cleared
  //
  // set the first number displayed on the display to total
  
  if ( firsttime && !Operator.isUnary( op ) ){
   total = getNumberOnDisplay();  	
   firsttime = false;
   } 
    
    
    if ( Operator.isUnary( op ) ) {   	
    	handleUnaryOp( op );        
    }
	else if ( lastOp != null ) {		
		handleBinaryOp( lastOp );		
	}

   //store the calculator's last op -- important for binary operators
   if ( !Operator.isUnary( op ) )
         lastOp = op;
           
}

//this method handles unary operators
void handleUnaryOp( String op )
{
	 if ( op.equals( Operator.NEGATE ) )
       {  
         //negate the number on the display screen 	 
         number = negate( getNumberOnDisplay()+"");    	    	    
         display.append(  number + "" );
         return;        	
       }
       else if ( op.equals( Operator.DOT ) )
       {
       	 handleDecPoint();
       	 return;
       }
       else if ( op.equals( Operator.SQRT ) )
       {  
         //calculate the square root of the number on the display    	 			     
	     number = Math.sqrt( getNumberOnDisplay() );
	     display.append( number+"" );
	     return;	     	   
       }      
       else if ( op.equals( Operator.EQUALS ) )
       {
       	 //if a binary operator was pressed before the equals
       	 //complete the operation first
       	 if ( lastOp != null && !Operator.isUnary( lastOp ) ) 
       	      handleBinaryOp( lastOp );
       	 
       	 lastOp = null;
         firsttime = true;
         return;
       }
       else 
       clear();   
}

// handles operators that require two operands
void handleBinaryOp( String op )
{
	if ( op.equals( Operator.ADD ) ) 
		 total += number;	
	else if ( op.equals( Operator.SUBTRACT ) )
		 total -= number;
	else if ( op.equals( Operator.MULTIPLY ) )
		 total *= number;
	else if ( op.equals( Operator.DIVIDE ) )
		 total /= number;
	else if ( op.equals( Operator.POW ) )
	     total = Math.pow( total, number );
					   
    display.append( total+"" );		    	
}

//This method is called each time a number is pressed.
//A string object is used to concatenate each number pressed in succession
//which then is converted into a double data type.

void handleNumber( String s )
{
	//concatenate to strVal if an operator was pressed before the current
	//button
	if ( !operatorPressed )
	strVal+=s;
	
	//if an operator was pressed, clear strVal and store the first number
	//pressed
	else {
		operatorPressed = false;
		strVal = s;		
	}
	
	//convert strVal to double 
	number = new Double( strVal ).doubleValue();	
	display.append( strVal );
	
}

//this method is called when the decimal point button has been pressed
void handleDecPoint()
{
	operatorPressed = false;
	
	//put a decimal point at the end of strVal only if there is no
	//decimal point already
	if ( strVal.indexOf( "." ) < 0 ) {
	     strVal+=Operator.DOT;	
	}       
	display.append( strVal );
}

//Method used to negate a value
double negate( String s )
{
	operatorPressed = false;
	
	//if number is a whole number, get rid of the '0' at the end of the
	//number to allow more number's to be added to the right of the
	//decimal point.
	
	if ( number == ( int ) number )
	     s = s.substring( 0, s.indexOf( "." ) );
	
	//add negative sign to the number if it doesn't exist	
	if ( s.indexOf( "-" ) < 0 )
	     strVal = "-"+s;
	
	//If a negative sign exist's, remove it
	else
	     strVal = s.substring( 1 );	  
	     
return new Double( strVal ).doubleValue();  
}

//get's the number that is currently on the display screen and
//convert it to double
double getNumberOnDisplay() {
   return new Double( display.getText() ).doubleValue();	
}

//clear's screen and reset's all variables
void clear()
{
	firsttime = true;
	lastOp = null;
	strVal = "";
	total = 0;
	number = 0;
	display.clear();
	display.append( "0" );
}

}