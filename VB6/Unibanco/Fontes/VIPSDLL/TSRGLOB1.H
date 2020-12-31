/*****************************************************************/
/*    filename:   tsrglob1.h                                     */
/*    owner:      dpet                                           */
/*    last modif: 92/09/15                                       */
/*    content:    All important def. for tsrexe.c and tsrinit.c. */
/*****************************************************************/

void tsr(void);
void exec_res(void);

/*********************/
/* Global variables. */
/*********************/
#ifdef _DECLARE_
   /***********************************************/
   /* Global variable for popup to set and test.  */
   /***********************************************/
   int unloading;	                  /* TSR unload flag. */
   char signature[] = "MAGICSIG";	/* the TSR's signature. */
   int magic_int;                   /* Interrupt vector. */

#else
   extern int unloading;
   extern char signature[];
   extern int magic_int;

   /**************************************************/
   /* Register value before and after a call to TSR. */
   /**************************************************/
   extern union REGS regIN, regOUT;   
   extern struct SREGS sReg;
#endif
