/******************************************************************/
/*    filename:   tsrdef1.h                                       */
/*    owner:      dpet                                            */
/*    last modif: 92/09/15                                        */
/*    content:    All important def. for tsrexe.c and tsrutil?.c. */
/******************************************************************/

   /*********************/
   /* Standard Include. */
   /*********************/
   #include <dos.h>
   #include <alloc.h>
   #include <conio.h>
   #include <stdio.h>
   #include <stdlib.h>
   #include <string.h>


   /*****************************************/
   /* Include for HRA structure definition. */
   /*****************************************/
   #include "tsrreg.h"


   /*************************/
   /* Miscellaneous define. */
   /*************************/
   #define BYTE unsigned char
   #define WORD unsigned int
   #define READ                0
   #define WRITE               1
   #define UNABLE              0
   #define DISABLE             0
   #define ENABLE              1
   #define FIND_ACTIVE_VGA     0xA

   /**************************************************************/
   /* Macro definition for writing to slot specific IO register. */
   /*   x   : slot number.                                       */
   /*   y   : address of IO.                                     */
   /*   z   : value.                                             */
   /**************************************************************/
   #define WriteIoSpecific(x,y,z) outportb( ((unsigned int)x<<12) + y, z)
   #define ReadIoSpecific(x,y)    inportb( ((unsigned int)x<<12) + y)
   #define ReadIoSpecificWord(x,y) inport( ((unsigned int)x<<12) + y)

   
   /*********************************/
   /* Magic Tsr Register structure. */
   /*********************************/
   struct MagicTsrReg
      {
      HRA  far *Access[4];
      BYTE ActiveSlot;
      };


   /************************************************/
   /* Standard configuration data block structure. */
   /************************************************/
   struct EisaCfgInfo
      {
      WORD IdChar2A   : 2;
      WORD IdChar1    : 5;
      WORD            : 1;
      WORD IdChar3    : 5;
      WORD IdChar2B   : 3;

      BYTE buf1[113];

      WORD MemCfg1      : 7;
      WORD MemLastEntry : 1;
      WORD MemDataSize1 : 8;
      BYTE MemStartAdr1[3];
      BYTE MemSize1[2];
      BYTE MemCfg2;
      BYTE MemDataSize2;
      BYTE MemStartAdr2[3];
      BYTE MemSize2[2];
      BYTE MemCfg3;
      BYTE MemDataSize3;
      BYTE MemStartAdr3[3];
      BYTE MemSize3[2];

      BYTE buf2[42];
     
      WORD IntUse  : 4;
      WORD         : 1;   
      WORD IntTrig : 1;
      WORD         : 1;   
      WORD MoreInt : 1;
      WORD         : 8;
      
      BYTE buf3[12];
      
      WORD DmaChannelNumber : 3;
      WORD                  : 5;
      WORD                  : 2;
      WORD DmaAdrMode       : 2;
      WORD DmaTiming        : 2;
      WORD                  : 2;
      };

      


