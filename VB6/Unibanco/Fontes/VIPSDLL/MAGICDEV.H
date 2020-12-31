/*
    filename:   magicdev.h
    Synopsis:
    Date:       92.12.4

COPYRIGHT (c) 1992 Matrox Electronic Systems Ltd.
All Rights Reserved.
*/

#ifndef __MAGICDEV_H__
#define __MAGICDEV_H__

#include <tseries.h>
#include <tsrfunc.h>
#include <treg.h>
#include <tglob.h>
#include <misc.h>
#include <ml_head.h>
#include <proto.h>
#include <pr_prot.h>
#include <dma_prot.h>
#include <dma_util.h>
#include <mltiff.h>

#ifndef ALLOC_TGLOB
#ifdef __cplusplus
extern "C" {
#endif /* __cplusplus */
extern T_GLOB  PTYPE t_glob;
extern T_USER  PTYPE t_user;
extern ML_TIFF PTYPE ml_tiff;
extern T_USER  PTYPE tt_user[T_N_OF_BOARDS];
extern T_DEVICE PTYPE t_device;
#ifdef __cplusplus
}
#endif /* __cplusplus */

#endif /* !ALLOC_TGLOB */
#endif /* __MAGICDEV_H__ */

