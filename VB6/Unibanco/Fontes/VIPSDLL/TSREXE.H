/**************************************************************/
/*    filename:   tsrexe.h                                    */
/*    owner:      dpet                                        */
/*    last modif: 92/09/15                                    */
/*    content:    Function prototype and Variable definition. */
/**************************************************************/

   /***************************************************/
   /* Function prototype for tsrexe.c and tsrutil?.c. */
   /***************************************************/
   void BoardInfo(void);         /* In tsrutil2.c */
   void Initialize(void);        /* In tsrutil1.c                    */
                                 /* Prototype is also in tsrglob1.c. */
   void GetAttribut(void);       /* Get References and levels in FLASH EEPROM */
   void InitIntController(void); /* Init interruption controller. */
   void GetVersion(void);        /* Get version of the MAGIC TSR. */
   void CfgDone(void);           /* Obtain information about XILINX load. */
   void GOutPort(void);          /* change bits in the GOUT_PORT register. */
   void UpdateFinished(void);    /* Update "finished". */
   void WaitFinished(void);      /* Polled "finished" until a bit is 0. */
   void DmaWriteReg(void);       
   void DmaReadReg(void);
   void DmaProgReg(void);
   void WaitDmaFinished(void);   /* Polled "dmafinished" until a bit is 0. */
   void GetDmaFinished(void);    /* Get dmafinished pointer */
   void GetGrabFinished(void);   /* Get grabfinished pointer */
   void GetVsyncFinished(void);  /* Get vsyncfinished pointer */
   void GetGrabDone(void);       /* Get grabdone pointer */
   void GetDmaDone(void);        /* Get dmadone pointer */
   unsigned long AscHexToLong(char *str);  /* conversion function. */
   void SubSysCntl(void);        /* Write in the SUBSYS_CNTL S3 registers */
   void GetDispCntl(void);       /* Get the address of the display control
                                    registers */
   void GetVsyncDone(void);      /* Get vsyncdone pointer */
   void GetDispVsyncDone(void);  /* Get dispvsyncdone pointer */
   void GetHsyncDone(void);      /* Get hsyncdone pointer */

   /*****************************************************/
   /* Prototype for the interrupt handler of the MAGIC. */
   /*****************************************************/
   void interrupt MagicInt(void);


/*********************************************************************/
             /* Variable definition for tsrexe.c */
/*********************************************************************/

#ifdef _DECLARE_
   /************************************/
   /* Structure for board information. */
   /************************************/
   struct MagicTsrBoard MagicBd[BOARD_MAX];  
   struct MagicAdjAtt   MagicBdAtt[BOARD_MAX];

   /*********************************/
   /* Structure for MAGIC register. */
   /*********************************/
   struct MagicTsrReg MagicReg;  

   /*****************************************/
   /* Structure for MAGIC display register. */
   /*****************************************/
   struct MagicDispInfo MagicDisp[4];  

   /**************************************************/
   /* Register value before and after a call to TSR. */
   /**************************************************/
   union REGS regIN, regOUT;   
   struct SREGS sReg;

   /***********************/
   /* Total board number. */
   /***********************/
   unsigned TotBdNb;

   /*****************************************/
   /* Indice use in the parameter's choice. */
   /*****************************************/
   unsigned short indice;
   char **Argv;

	/**********************************/
   /* Use for the interrupt polling. */
   /**********************************/
   unsigned long  dmafinished[4] = {0,0,0,0};
   unsigned long  grabfinished[4] = {0,0,0,0};
   unsigned long  vsyncfinished[4] = {0,0,0,0};
   unsigned long  grabdone[4] = {0,0,0,0};
   unsigned long  dmadone[4] = {0,0,0,0};
   unsigned long  vsyncdone[4] = {0,0,0,0};
   unsigned long  dispvsyncdone[4] = {0,0,0,0};
   unsigned long  hsyncdone[4]={0,0,0,0};

   /*****************************************************************/
   /* Function pointer to old handler of the interruption hooked by */
   /* MagicInt handler.                                             */
   /*****************************************************************/
   void interrupt (*oldhandlerMagicInt)(void);

#else
   /************************************/
   /* Structure for board information. */
   /************************************/
   extern struct MagicTsrBoard MagicBd[BOARD_MAX];  
   extern struct MagicAdjAtt   MagicBdAtt[BOARD_MAX];

   /*********************************/
   /* Structure for MAGIC register. */
   /*********************************/
   extern struct MagicTsrReg MagicReg;  

   /*****************************************/
   /* Structure for MAGIC display register. */
   /*****************************************/
   extern struct MagicDispInfo MagicDisp[4];  

   /**************************************************/
   /* Register value before and after a call to TSR. */
   /**************************************************/
   extern union REGS regIN, regOUT;   
   extern struct SREGS sReg;

   /***********************/
   /* Total board number. */
   /***********************/
   extern unsigned TotBdNb;   

   /*****************************************/
   /* Indice use in the parameter's choice. */
   /*****************************************/
   extern unsigned short indice;
   extern char **Argv;

	/**********************************/
   /* Use for the interrupt polling. */
   /**********************************/
   extern unsigned long  dmafinished[4];
   extern unsigned long  grabfinished[4];
   extern unsigned long  vsyncfinished[4];
   extern unsigned long  grabdone[4];
   extern unsigned long  dmadone[4];
   extern unsigned long  vsyncdone[4];
   extern unsigned long  dispvsyncdone[4];
   extern unsigned long  hsyncdone[4];

   /*****************************************************************/
   /* Function pointer to old handler of the interruption hooked by */
   /* MagicInt handler.                                             */
   /*****************************************************************/
   extern void interrupt (*oldhandlerMagicInt)(void);
#endif

