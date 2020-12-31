/******************************************************************/
/*    filename:   tsrfunc.h                                       */
/*    owner:      dpet                                            */
/*    last modif: 92/09/21                                        */
/*    content:    All important def. for tsrexe.c and tsrutil?.c  */
/*                and needed by the LowLevel SW.                  */
/******************************************************************/

#ifndef __TSRFUNC_H__
#define __TSRFUNC_H__


   /****************************************************/
   /* Function number that we can find in register AH  */
   /* for execute when we have a Magic interruption.   */
   /****************************************************/
   #define UNLOAD_TSR           0x00
   #define BOARD_INFO           0x01
   #define GET_VERSION          0x02
   #define CFG_DONE             0x05
   #define GOUT_PORT            0x07
   #define DMA_WRITE_REG        0x0D
   #define DMA_READ_REG         0x0E
   #define DMA_PROG_REG         0x0F
   #define GET_DMAFINISHED      0x11
   #define GET_GRABFINISHED     0X12
   #define GET_VSYNCFINISHED    0X13
   #define GET_GRABDONE         0X14
   #define GET_DMADONE          0X15
   #define SET_SUBSYS_CNTL      0x16
   #define GET_DISP_CNTL        0x17
   #define GET_VSYNCDONE        0x18
   #define GET_DISPVSYNCDONE    0x19
   #define GET_HSYNCDONE        0x20
   #define GET_ATTRIBUT         0x21


   /*************************/
   /* Miscellaneous define. */
   /*************************/
   #define COMPRESSED_ID    0x00009836
   #define BOARD_MAX        4         /* Maximum board number. */
   #define MAGIC_VERSION    32        /* Version of the MAGIC TSR. */
   #define MAGIC_ID         0xCAFE    /* ID for the MAGIC TSR. */
   #define DMA_READ_A       0xA
   #define DMA_READ_B       0xB
   #define DMA_READ_C       0xC
   #define DMA_WRITE_A      0xA
   #define DMA_WRITE_B      0xB
   #define DMA_WRITE_C      0xC
   #define ML_UNABLE                               0
   #define ML_ENABLE                               1
   #define KEYING_ENABLE                           0
   #define FRAME_BUFFER_ACCESS                     1

   /****************************/
   /* MagicTsrBoard structure. */
   /****************************/
   struct MagicTsrBoard
      {
      unsigned long MemAdr1;
      unsigned long HighMapAdr;
      unsigned long HighMapSize;
      unsigned short IntUse  : 4;
      unsigned short        : 1;
      unsigned short IntTrig : 1;
      unsigned short        : 1;
      unsigned short MoreInt : 1;

      unsigned short DmaChannelNumber : 3;
      unsigned short DmaAdrMode       : 2;
      unsigned short DmaTiming        : 2;
      unsigned short                  : 1;

      unsigned short Slot             : 8;
      unsigned short                  : 8;
      unsigned long LinearAddress;        // initialize to 0.
      unsigned short DmaRead;
      unsigned short RefreshRate;
      unsigned short BoardRev;
      unsigned short VgaEnable;
      unsigned short S3SegRegMap;
      unsigned long VgaLinearAddress;
      };

   typedef struct MagicTsrBoard MAGICTSRBOARD, *PMAGICTSRBOARD;


   struct MagicAdjAtt
     {
      unsigned short StructRev;
      unsigned short RedLevel;
      unsigned short RedRef;
      unsigned short GreenLevel;
      unsigned short GreenRef;
      unsigned short BlueLevel;
      unsigned short BlueRef;
      unsigned short NtscBright;
      unsigned short NtscContrast;
      unsigned short NtscSaturation;
      unsigned short NtscHue;
      unsigned short PalBright;
      unsigned short PalContrast;
      unsigned short PalSaturation;
      unsigned short PalHue;
      unsigned short YcNtscBright;
      unsigned short YcNtscContrast;
      unsigned short YcNtscSaturation;
      unsigned short YcNtscHue;
      unsigned short YcPalBright;
      unsigned short YcPalContrast;
      unsigned short YcPalSaturation;
      unsigned short YcPalHue;

      };

   typedef struct MagicAdjAtt MAGICADJATT, *PMAGICADJATT;


   /************************************************/
   /* Structure containing display information     */
   /************************************************/
   struct MagicDispInfo
      {
      unsigned long SrtStartl;
      unsigned long SrtStartm;
      unsigned long SrtStarth;
      unsigned long DispVsyncInt;
      unsigned long DispChanged;
      };

   typedef struct MagicDispInfo MAGICDISPINFO, *PMAGICDISPINFO;



/****************************************************************************/

   /****************************************************************/
   /* void BoardInfo( void )                                       */
   /*   IN   parameters  AH = 0x01.                                */
   /*   OUT  parameters  AH = error code : 0 ->success.            */
   /*                    BH = total board number found.            */
   /*                    ES = segment for MagicTsrBoard structure. */
   /*                    DI = offset for MagicTsrBoard structure.  */
   /*                                                              */
   /****************************************************************/

   /****************************************************************/
   /* void GetVersion( void )                                      */
   /*   IN   parameters  AH = 0x02                                 */
   /*   OUT  parameters  AH = error code                           */
   /*                         : 0 ->success.                       */
   /*                    BX = Version.                             */
   /*                    CX = MAGIC ID ->0xCAFE.                   */
   /*                                                              */
   /****************************************************************/

   /****************************************************************/
   /* void CfgDone( void )                                         */
   /*   IN   parameters  AH = 0x05                                 */
   /*   OUT  parameters  AH = error code                           */
   /*                         : 0 ->success.                       */
   /*                    CL = : 0 ->One or more FPGAs are not      */
   /*                               configured.                    */
   /*                         : 1 ->All FPGAs are oprational.      */
   /*                                                              */
   /****************************************************************/

   /****************************************************************/
   /* void GOutPort( void )                                        */
   /*   IN   parameters  AH = 0x07.                                */
   /*                    BL = KEYING_ENABLE = 0                    */
   /*                         FRAME_BUFFER_ACCESS = 1.             */
   /*                    BH = UNABLE=0 ENABLE=1.                   */
   /*   OUT  parameters  AH = error code : 0 ->success.            */
   /*                                                              */
   /****************************************************************/

   /****************************************************************/
   /* void DmaWriteReg( void )                                     */
   /*   IN   parameters  AH = 0x0D                                 */
   /*   OUT  parameters  AH = error code                           */
   /*                         : 0 ->success.                       */
   /*                                                              */
   /****************************************************************/

   /****************************************************************/
   /* void DmaReadReg( void )                                      */
   /*   IN   parameters  AH = 0x0E                                 */
   /*   OUT  parameters  AH = error code                           */
   /*                         : 0 ->success.                       */
   /*                                                              */
   /****************************************************************/

   /****************************************************************/
   /* void DmaProgRag( void )                                      */
   /*   IN   parameters  AH = 0x0F                                 */
   /*   OUT  parameters  AH = error code                           */
   /*                         : 0 ->success.                       */
   /*                                                              */
   /****************************************************************/

   /****************************************************************/
   /* void GetDmaFinished( void )                                  */
   /*   IN   parameters  AH = 0x11                                 */
   /*                    AL = board number                         */
   /*   OUT  parameters  AH = error code                           */
   /*                         : 0 ->success.                       */
   /*                    BX = segment of dmafinished.              */
   /*                    CX = offset of dmafinished.               */
   /*                    AL = Read value.                          */
   /*                                                              */
   /****************************************************************/

   /****************************************************************/
   /* void GetGrabFinished( void )                                 */
   /*   IN   parameters  AH = 0x12                                 */
   /*                    AL = board number                         */
   /*   OUT  parameters  AH = error code                           */
   /*                         : 0 ->success.                       */
   /*                    BX = segment of grabfinished.             */
   /*                    CX = offset of grabfinished.              */
   /*                    AL = Read value.                          */
   /*                                                              */
   /****************************************************************/

   /****************************************************************/
   /* void GetVsyncFinished( void )                                */
   /*   IN   parameters  AH = 0x13                                 */
   /*                    AL = board number                         */
   /*   OUT  parameters  AH = error code                           */
   /*                         : 0 ->success.                       */
   /*                    BX = segment of vsyncfinished.            */
   /*                    CX = offset of vsyncfinished.             */
   /*                    AL = Read value.                          */
   /*                                                              */
   /****************************************************************/
   /****************************************************************/
   /* void GetGrabDone( void )                                     */
   /*   IN   parameters  AH = 0x14                                 */
   /*                    AL = board number                         */
   /*   OUT  parameters  AH = error code                           */
   /*                         : 0 ->success.                       */
   /*                    BX = segment of grabdone.                 */
   /*                    CX = offset of grabdone.                  */
   /*                    AL = Read value.                          */
   /*                                                              */
   /****************************************************************/

   /****************************************************************/
   /* void GetDmaDone( void )                                      */
   /*   IN   parameters  AH = 0x15                                 */
   /*                    AL = board number                         */
   /*   OUT  parameters  AH = error code                           */
   /*                         : 0 ->success.                       */
   /*                    BX = segment of dmadone.                  */
   /*                    CX = offset of dmadone.                   */
   /*                    AL = Read value.                          */
   /*                                                              */
   /****************************************************************/

   /****************************************************************/
   /* void SubSysCntl( void )                                      */
   /*   IN   parameters  AH = 0x16                                 */
   /*                    BX = SUBSYS_CNTL register value           */
   /*   OUT  parameters  AH = error code                           */
   /*                         : 0 ->success.                       */
   /*                                                              */
   /****************************************************************/

   /****************************************************************/
   /* void GetDispCntl( void )                                     */
   /*   IN   parameters  AH = 0x17                                 */
   /*                    AL = board number                         */
   /*   OUT  parameters  AH = error code                           */
   /*                         : 0 ->success.                       */
   /*                    ES = segment of MagicDisp structure       */
   /*                    DI = offset of MagicDisp structure        */
   /*                                                              */
   /****************************************************************/

   /****************************************************************/
   /* void GetVsyncDone( void )                                    */
   /*   IN   parameters  AH = 0x18                                 */
   /*                    AL = board number                         */
   /*   OUT  parameters  AH = error code                           */
   /*                         : 0 ->success.                       */
   /*                    BX = segment of vsyncdone.                */
   /*                    CX = offset of vsyncdone.                 */
   /*                    AL = Read value.                          */
   /*                                                              */
   /****************************************************************/

   /****************************************************************/
   /* void GetDispVsyncDone( void )                                */
   /*   IN   parameters  AH = 0x19                                 */
   /*                    AL = board number                         */
   /*   OUT  parameters  AH = error code                           */
   /*                         : 0 ->success.                       */
   /*                    BX = segment of dispvsyncdone.            */
   /*                    CX = offset of dispvsyncdone.             */
   /*                    AL = Read value.                          */
   /*                                                              */
   /****************************************************************/

   /****************************************************************/
   /* void GetHsyncDone( void )                                    */
   /*   IN   parameters  AH = 0x20                                 */
   /*                    AL = board number                         */
   /*   OUT  parameters  AH = error code                           */
   /*                         : 0 ->success.                       */
   /*                    BX = segment of hsyncdone.                */
   /*                    CX = offset of hsyncdone.                 */
   /*                    AL = Read value.                          */
   /*                                                              */
   /****************************************************************/


   /****************************************************************/
   /* void GetAttribut( void )                                     */
   /*   IN   parameters  AH = 0x21.                                */
   /*   OUT  parameters  AH = error code : 0 ->success.            */
   /*                    ES = segment for MagicAdjAtt structure.   */
   /*                    DI = offset for MagicAdjAtt structure.    */
   /*                                                              */
   /****************************************************************/


#endif




