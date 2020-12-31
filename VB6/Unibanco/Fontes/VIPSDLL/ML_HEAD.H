/*
    filename:   ml_head.h
    Synopsis:   this header file contains defines passed to user's function
    Date:       15-09-92  MLO  created

COPYRIGHT (c) 1992 Matrox Electronic Systems Ltd.
All Rights Reserved.
*/

#ifndef __ML_HEAD_H__
#define __ML_HEAD_H__

#define ML_SUCCESS               0
#define ML_ERROR                -1
#define ML_DIGITIZER_NOT_LOADED -1

#define ML_NULL     0
#define ML_DEFAULT  0

// mlBoardSelect()
#define ML_BOARD1   0
#define ML_BOARD2   1
#define ML_BOARD3   2
#define ML_BOARD4   3

// define for mlAllocateDevice()
#define ML_ALLOC_BOARD1      1
#define ML_ALLOC_BOARD2      2
#define ML_ALLOC_BOARD3      4
#define ML_ALLOC_BOARD4      8

// mlGrabStartSync
#define ML_ACTIVE1  0x01
#define ML_ACTIVE2  0x02
#define ML_ACTIVE3  0x04
#define ML_ACTIVE4  0x08

#define ML_SYNC1  ML_BOARD1
#define ML_SYNC2  ML_BOARD2
#define ML_SYNC3  ML_BOARD3
#define ML_SYNC4  ML_BOARD4

// Lut functions
#define ML_LUT_RAMP        1           /* define ramp     */
#define ML_LUT_INVERSE     2           /* define inverse ramp */
#define ML_LUT_CONST       3           /* define constant */
#define ML_COLOR_LUT       1           /* true color mode */
#define ML_MONO_LUT        2           /* monochrome color mode */
#define ML_DISP_LUT        1           /* display lut */
#define ML_DIGITIZER_LUT   2           /* digitizer lut */
#define ML_PALETTE0        0x00
#define ML_PALETTE1        0x08
#define ML_PALETTE2        0x10
#define ML_PALETTE3        0x18

#define ML_MONO_LUT_SIZE   256
#define ML_COLOR_LUT_SIZE  256*3

// vga keying
#define ML_KEY_VGAONLY    1
#define ML_KEY_NOVGA      2
#define ML_KEY_ENABLE     3

// new values documented in magiclib rev 2.0
#define ML_KEY_NORMAL     2
#define ML_KEY_REVERSE    1


// obsolete values kept for compatibility with revisions 1.0 and 1.1
#define ML_KEY_VGAEQUAL   1
#define ML_KEY_VGANEQUAL  2

// memory organizations
#define ML_ORG_3x1Kx1Kx8      0x00
#define ML_ORG_1x1Kx1Kx24     0x80
#define ML_ORG_3x2Kx512x8     0x10
#define ML_ORG_1x2Kx512x24    0x90
#define ML_ORG_3x4Kx256x8     0x20
#define ML_ORG_1x4Kx256x24    0xA0
#define ML_ORG_3x8Kx128x8     0x30
#define ML_ORG_1x8Kx128x24    0xB0
#define ML_ORG_1x1Kx3Kx8      0x40
#define ML_ORG_1x2Kx1_5Kx8    0x50
#define ML_ORG_1x4Kx768x8     0x60
#define ML_ORG_1x8Kx384x8     0x70
#define ML_ORG_UNINITIALIZED  0x7f

// mask function
#define ML_MASK_HOSTREAD      0x01
#define ML_MASK_HOSTWRITE     0x02
#define ML_MASK_DMA           0x04
#define ML_MASK_GRAB          0x08
#define ML_MASK_DISPLAY       0x10

// read and write line defines
#define ML_LINE               0
#define ML_COLUMN             1

// defines for surface ID
#define ML_IMAGE_RED          0
#define ML_IMAGE_GREEN        1
#define ML_IMAGE_BLUE         2
#define ML_IMAGE              3

// tiff mode defines
#define ML_TIFF_1BAND         0
#define ML_TIFF_3BAND         1
#define ML_MODE_1BAND         2

// ram tiff defines
#define ML_TIFF_WIDTH         0
#define ML_TIFF_HEIGHT        1
#define ML_TIFF_MODE          2

// type of buffers
#define ML_TIFF_DATABUF       0
#define ML_TIFF_DMABUF        1

// defines for mlGrabIteration
#define ML_FOR                1
#define ML_FOREVER            2
#define ML_FIELD              0xC0
#define ML_NEXT_FIELD         0xE0

// defines for mlGrabMode
#define ML_SW_TRIG       0x00
#define ML_HW_TRIG       0x08
#define ML_DIS_VRESET    0x10
#define ML_FALLING_EDGE  0x20
#define ML_RISING_EDGE   0x01

// defines for mlCamSelectChannel
#define ML_CHANNEL_0    0
#define ML_CHANNEL0     0
#define ML_CHANNEL_1    1
#define ML_CHANNEL1     1
#define ML_CHANNEL_2    2
#define ML_CHANNEL2     2
#define ML_CHANNEL_SYNC 3
#define ML_CHANNEL_ALL  3
#define ML_CHANNEL_NONE 4
#define ML_CHANNEL_YC   5
#define ML_KEEP_SYNC   0x8000

// defines for mlCamGain
#define ML_GAIN0       0
#define ML_GAIN1       1
#define ML_GAIN2       2
#define ML_GAIN3       3

// defines for mlCamUserBitOut
#define ML_USER_BIT0   0
#define ML_USER_BIT1   1

// defines for mlGrabInterlace
#define ML_STORE_PROGRESSIVE 1
#define ML_STORE_INTERLACE   2

// defines for mlGrabUnderSampling
#define ML_SUBSAMP_1   1
#define ML_SUBSAMP1    1
#define ML_SUBSAMP_2   2
#define ML_SUBSAMP2    2
#define ML_SUBSAMP_4   4
#define ML_SUBSAMP4    4
#define ML_SUBSAMP_8   8
#define ML_SUBSAMP8    8

// mlGrabStop() mlGrabStopSync()
#define ML_IMMEDIATE  0x1
#define ML_NEXT_FRAME 0x2

// mlGrabStopSync
#define ML_MONOSHOT   0x3

// define for reference.
#define ML_RED_REF     0x100
#define ML_GREEN_REF   0x200
#define ML_BLUE_REF    0x400
#define ML_ALL_REF     0x000

#define ML_CHANNEL0_REF   ML_RED_REF
#define ML_CHANNEL1_REF   ML_GREEN_REF
#define ML_CHANNEL2_REF   ML_BLUE_REF
#define ML_ALLCHANNEL_REF ML_ALL_REF

// define for gain.
#define ML_RED_GAIN     0x100
#define ML_GREEN_GAIN   0x200
#define ML_BLUE_GAIN    0x400
#define ML_ALL_GAIN     0x800

#define ML_CHANNEL0_GAIN   ML_RED_GAIN
#define ML_CHANNEL1_GAIN   ML_GREEN_GAIN
#define ML_CHANNEL2_GAIN   ML_BLUE_GAIN
#define ML_ALLCHANNEL_GAIN ML_ALL_GAIN

// define for gain, reference and bt812 functions.
#define ML_SET_IMMEDIATE       0x1000

// defines for mlSysInquire
#define ML_INQ_MEMORG               1
#define ML_INQ_GRAB_ITERATION       2
#define ML_INQ_GRAB_CONDITION       3
#define ML_INQ_GRAB_SURFACE         4
#define ML_INQ_GRAB_XPOSITION       5
#define ML_INQ_GRAB_YPOSITION       6
#define ML_INQ_GRAB_XSIZE           7
#define ML_INQ_GRAB_YSIZE           8
#define ML_INQ_GRAB_XSUBSAMPLING    9
#define ML_INQ_GRAB_YSUBSAMPLING    10
#define ML_INQ_DRIVER_VERSION       11
#define ML_INQ_MAGICLIB_VERSION     12
#define ML_INQ_CURRENT_BOARD        13
#define ML_INQ_NUMBER_OF_BOARD      14
#define ML_INQ_HIGH_MAP_WINDOW      15
#define ML_INQ_MASK_HOSTREAD        16
#define ML_INQ_MASK_HOSTWRITE       17
#define ML_INQ_MASK_DMA             18
#define ML_INQ_MASK_GRAB            19
#define ML_INQ_MASK_DISPLAY         20
#define ML_INQ_DISP_MODE            21
#define ML_INQ_DISP_XSTART          22
#define ML_INQ_DISP_YSTART          23
#define ML_INQ_BOARD_REV            24
#define ML_INQ_HUE                  25
#define ML_INQ_SATURATION           26
#define ML_INQ_BRIGHTNESS           27
#define ML_INQ_CONTRAST             28
#define ML_INQ_REFWHITE_RED         29
#define ML_INQ_REFWHITE_GREEN       30
#define ML_INQ_REFWHITE_BLUE        31
#define ML_INQ_REFBLACK_RED         32
#define ML_INQ_REFBLACK_GREEN       33
#define ML_INQ_REFBLACK_BLUE        34
#define ML_INQ_GRAB_CHANNEL         35
#define ML_INQ_SYNC_CHANNEL         36
#define ML_INQ_BOARD_TYPE           37
#define ML_INQ_SYNC_PRESENT         38
#define ML_INQ_DISP_SURFACE         39
#define ML_INQ_GRAB_PIXDEPTH        40
#define ML_INQ_PLL_LOCKED           41
#define ML_INQ_S3SEGREGMAP          42
#define ML_INQ_VGA_ENABLED          43
#define ML_INQ_EXPSIG_MODE          44
#define ML_INQ_EXPSIG_T0            45
#define ML_INQ_EXPSIG_T1            46
#define ML_INQ_DISPXSIZE            47
#define ML_INQ_DISPYSIZE            48
#define ML_INQ_USER_OUTBIT0         49
#define ML_INQ_USER_OUTBIT1			50
#define ML_INQ_USER_INBIT				51
#define ML_INQ_USER_OUTBITSTATE		52
#define ML_INQ_S3SEGREGMAP_SEG      53
#define ML_INQ_BOARD_STATUS         54


// defines for ML_INQ_BOARD_TYPE
#define ML_FB_RED                   0x01
#define ML_FB_GREEN                 0x02
#define ML_FB_BLUE                  0x04
#define ML_BT812                    0x08
#define ML_24BIT_MODULE             0x10

// Patch for MIL
#define ML_COLOR                    ML_FB_GREEN

// defines for ML_INQ_SYNC_PRESENT
#define ML_SYNC_NOT_PRESENT         0x0
#define ML_SYNC_PRESENT             0x1
#define ML_UNKNOWN                  0x2

// defines for ML_INQ_BOARD_STATUS
#define ML_BOARD_IDLE               0x00
#define ML_BOARD_GRABING            0x01
#define ML_BOARD_DMAREAD            0x02
#define ML_BOARD_DMAWRITE           0x04

// Defines for ML_INQ_EXPSIG
#define ML_EXPSIG_PERIODIC				0x1
#define ML_EXPSIG_TRIGGER 				0x2

// Defines for ML_INQ_PLL_LOCKED
#define ML_PLL_LOCKED               0x1
#define ML_PLL_NOT_LOCKED           0x0

// used by interpreter
#define ML_TCAMERA_SIZE 500

// defines for GOUTPORT registers.
#define KEYING_ENABLE       0
#define FRAME_BUFFER_ACCESS 1

// Defines for User bits
#define ML_ONE					1
#define ML_ZERO				0
#define ML_ACTIVE				1
#define ML_TRISTATE			0
#define ML_USERBIT0			0
#define ML_USERBIT1			1

// defines for interrupts  flags
#define ML_DMADONE          0
#define ML_GRABDONE         1
#define ML_VSYNCDONE        2
#define ML_DISPVSYNCDONE    3
#define ML_HSYNCDONE        4

#endif /* __ML_HEAD_H__ */



