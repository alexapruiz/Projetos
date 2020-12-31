/**
*  File:  		 Operator.java
*  Description:  Utility class used for the calculator's operators
*
**/

class Operator {


final static String ADD = "+",
    				SUBTRACT = "-",
    				MULTIPLY = "x",
    				DIVIDE = "/",
    				POW = "pow",
    				SQRT = "sqrt",
    				CLEAR = "C",
    				EQUALS = "=",
    				NEGATE = "+/-",
    				DOT = ".";

//prevent's class from being instantiated  				    	
private Operator() {}

//determine's if the operator is unary ( and operator that need's only one
//operand ).	
static boolean isUnary( String s ) 
{
	return    s.equals( EQUALS )
	       || s.equals( CLEAR )
	       || s.equals( SQRT ) 
	       || s.equals( NEGATE )
	       || s.equals( DOT );	       		
}

}