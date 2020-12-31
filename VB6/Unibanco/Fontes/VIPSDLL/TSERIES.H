/*
    filename :  tseries.h

COPYRIGHT (c) 1992 Matrox Electronic Systems Ltd.
All Rights Reserved

*/

#ifndef __TSERIES_H__
#define __TSERIES_H__


/*
 * define one (1) of the following operating system files
 *
 *       #define T_WINDOWS
 *           for windows 3.1 and pharlap|286 dos extenders
 *       #define MA_MS_DOS
 *           for real mode dos (without dos extenders)
 *       #define  WATCOM
 *           for watcom compiler
 *       #define  MSC_NT
 *           for Windows NT with Microsoft tools
 *       #define  BORLAND_32
 *           for Borland PowerPack
 */

//    #define T_WINDOWS
//    #define MA_MS_DOS
//    #define WATCOM
//    #define MSC_NT
//    #define BORLAND_32

#ifdef MSC_NT
#include <win_nt.h>
#include <windows.h>
#include <winioctl.h>
#include <tsrfunc.h>
#include <magicdrv.h>
#endif  /* MSC_NT     */

#ifdef T_WINDOWS
#include <win310.h>
#endif  /* T_WINDOWS  */

#ifdef MA_MS_DOS
#include <msdos.h>
#endif  /* MA_MS_DOS  */

#ifdef WATCOM
#include <watcom.h>
#endif  /* WATCOM     */

#ifdef BORLAND_32
#include <b32.h>
#endif  /* BORLAND_32 */

#ifndef T_WINDOWS
#ifndef MA_MS_DOS
#ifndef WATCOM
#ifndef MSC_NT
#ifndef BORLAND_32
#include <win310.h>
#endif  /* BORLAND_32    */
#endif  /* MSC_NT        */
#endif  /* WATCOM        */
#endif  /* MA_MS_DOS     */
#endif  /* T_WINDOWS     */
#endif  /* __TSERIES_H__ */
