/*                                                                                                                                /*
    filename:   iofile.h
    Synopsis:   This header file is included only in iofile.c
                This file is used only by mlcamload and mlcamsave functions
    Date:       24-09-92  BFOU  created

COPYRIGHT (c) 1992 Matrox Electronic Systems Ltd.
All Rights Reserved.
*/

#ifndef __IOFILE__
#define __IOFILE__

// Product name used in the file
#define  PRODUCT_NAME   "MAGIC"

// Magic code include in configuration file
#define  CONF_MAG_CODE  "50CF"

// Section name keyword for the configuration file
#define  KW_CONF_HEAD   "[CONFIG_FILE]"
#define  KW_DIGIT_REG   "[REG_DIGIT]"
#define  KW_MAGIC_REG   "[REG_MAGIC]"

// Line length used for file access
// Note: Du to software behavior( \0 add at the end of string),
//       add 2 at the desired line length
#define  CFLLENGTH      120+2

// General purposes definition
#define  NO             0
#define  YES            1

// Define use in the extraction of a line
#define  GET_KEYWD      1              // use to get or not a line
#define  NGET_KYWD      2              // containning a keyword

#define  RM_COM         1              // use to remove or not
#define  NRM_COM        2              // a comments line

#define  CT_STR         1              // use to remove or not the space
#define  NT_STR         2              // that may following a line


// Error code - bitfield used to flag any special event/error
#ifndef MSC_NT
#define  NO_ERROR       0x00000000L
#endif /* MSC_NT */

#define  ERROR_FOUND    0x00000001L    // An error occurred somewhere


#define  ERR_RDFILE     0x00000010L    // Found error while reading file
#define  ERR_WRFILE     0x00000020L    // Same while writing one
#define  ERR_ALCMEM     0x00000040L    // Cannot allocate memory


#define  ERR_BFIELD     0x00000100L    // Bad field value
#define  ERR_FDKMIS     0x00000200L    // Field or keyword missing
#define  ERR_BKEYWD     0x00000400L    // Bad keyword or keyword not found
#define  ERR_BCFPRD     0x00000800L    // Bad configuration file for
                                       // the current product
#define  MES_UGFILE     0x10000000L    // Upgrade an existing config file
#define  MES_NAPARI     0x20000000L    // Not all parameter initialised while
                                       // loading config file

// --- digt_reg
#define NBR_DIG_REG        104
#define DIG_REG_OFFSET     0

#define NBR_MAG_REG        39
#define MAG_REG_OFFSET     (NBR_DIG_REG+0)      // + a gap

#define NBR_REGISTER       (NBR_DIG_REG + NBR_MAG_REG)

typedef struct _reg_interface
{
   char PTYPE* name;
   long PTYPE* value;
} register_interface;


#endif  /* __IOFILE__ */
