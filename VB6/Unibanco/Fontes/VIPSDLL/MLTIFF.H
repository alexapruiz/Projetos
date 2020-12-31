/**************************************************************************/
/*

    Filename:  MLTIFF.H
    Owner   :  mEid
    Adapted : H. Hasswani for the MAGIC project
    Rev     :  $Revision:   1.0  $
    Content :  This file contains the define and structures for the
               TIFF encoding/decoding routines

    COPYRIGHT (c) 1989, 1990, 1991, 1992 Matrox Electronic Systems Ltd.
    All Rights Reserved

*/
/**************************************************************************/

/*--- TIFF general defines -----------------------------------------------*/

#define ML_TIFF_INTEL       0x4949  /* INTEL byte order */
#define ML_TIFF_VERSION     0x002a  /* TIFF flag */
#define ML_TIFF_IMAGEDESCRIPTION "TIFF image saved with MAGIC Library "
#define ML_TIFF_RESOLUTIONX_NUM 75L /* TIFF x resolution : 75 pixels */
#define ML_TIFF_RESOLUTIONX_DEN  1L /*                     per inch  */
#define ML_TIFF_RESOLUTIONY_NUM 75L /* TIFF y resolution : 75 pixels */
#define ML_TIFF_RESOLUTIONY_DEN  1L /*                     per inch  */

/*--- TIFF tag value -----------------------------------------------------*/

#define ML_TIFF_NB_DE    16    /* nb of Directory Entry used when encoding */

#define ML_TIFF_TAG_SUBFILETYPE           0x00fe
#define ML_TIFF_TAG_IMAGEWIDTH            0x0100
#define ML_TIFF_TAG_IMAGELENGTH           0x0101
#define ML_TIFF_TAG_BITPERSAMPLE          0x0102
#define ML_TIFF_TAG_COMPRESSION           0x0103
#define ML_TIFF_TAG_PHOTOMETRICINTERP     0x0106
#define ML_TIFF_TAG_IMAGEDESCRIPTION      0x010e
#define ML_TIFF_TAG_STRIPOFFSET           0x0111
#define ML_TIFF_TAG_SAMPLEPERPIXEL        0x0115
#define ML_TIFF_TAG_ROWSPERSTRIP          0x0116
#define ML_TIFF_TAG_STRIPBYTECOUNT        0x0117
#define ML_TIFF_TAG_XRESOLUTION           0x011a
#define ML_TIFF_TAG_YRESOLUTION           0x011b
#define ML_TIFF_TAG_PLANARCONFIGURATION   0x011c
#define ML_TIFF_TAG_RESOLUTIONUNIT        0x0128
#define ML_TIFF_TAG_SOFTWARE              0x0131

/*--- TIFF tag type ------------------------------------------------------*/

#define ML_TIFF_BYTE                      0x0001
#define ML_TIFF_ASCII                     0x0002
#define ML_TIFF_SHORT                     0x0003
#define ML_TIFF_LONG                      0x0004
#define ML_TIFF_RATIONAL                  0x0005

#define ML_TIFF_BYTE_SIZE                 0x0001
#define ML_TIFF_ASCII_SIZE                0x0001
#define ML_TIFF_SHORT_SIZE                0x0002
#define ML_TIFF_LONG_SIZE                 0x0004
#define ML_TIFF_RATIONAL_SIZE             0x0008

#define ML_SIGNED                        0x10000000L

#define ML_FILE_MODEL_SIZE                128
#define ML_TIFF_BUF_SIZE                 4096


typedef unsigned char ML_FILE_MODEL[ML_FILE_MODEL_SIZE];
typedef unsigned char ML_TIFF_BUFFER[ML_TIFF_BUF_SIZE];

/*--- TIFF Image File Header ---------------------------------------------*/

typedef struct ml_tiff_ifh
{
  short ByteOrder;
  short Version;
  long  FirstIfdOffset;

} ML_TIFF_IFH;

/*--- TIFF Directory Entry -----------------------------------------------*/

typedef struct ml_tiff_de
{
  short Tag;
  short Type;
  long  Length;
  long  Offset;

} ML_TIFF_DE;

/*--- TIFF Image File Directory ------------------------------------------*/

typedef struct ml_tiff_ifd
{
  short DeCount;
  ML_TIFF_DE de[ML_TIFF_NB_DE];
  long  NextIfdOffset;
} ML_TIFF_IFD;

/*--- Global structure for tiff file */

typedef struct ml_tiff
{
   long sizex;
   long sizey;
   long sizez;
   long nb_band;
   long nb_time_slice;
   long type;
   long mode;
   long compression;
   long samplesPerPixel;
   long tagSubfileType;
   long rowsPerStrip;
   long bitsPerSample;
   long photometricInterp;
   long planarConfig;
   long resolutionxNum;
   long resolutionxDen;
   long resolutionyNum;
   long resolutionyDen;
   unsigned char MFAR * imageDesc;
   ML_FILE_MODEL model;

} ML_TIFF;


