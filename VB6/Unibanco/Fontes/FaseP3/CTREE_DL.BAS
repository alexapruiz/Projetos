Attribute VB_Name = "CTREE_DLL"
Option Explicit

'-------------------------------------------------------------------------------
' * This program is the CONFIDENTIAL and PROPRIETARY property
' * of FairCom(R) Corporation. Any unauthorized use, reproduction or
' * transfer of this computer program is strictly prohibited.
' *
' *      Copyright (c) 1984 - 1996 FairCom Corporation.
' * This is an unpublished work, and is subject to limited distribution and
' * restricted disclosure only. ALL RIGHTS RESERVED.
' *
' *       RESTRICTED RIGHTS LEGEND
' * Use, duplication, or disclosure by the Government is subject to
' * restrictions set forth in subparagraph (c)(1)(ii) of the Rights in
' *    Technical Data and Computer Software clause at DFARS 252.227-7013.
' * FairCom Corporation, 4006 West Broadway, Columbia, MO 65203.
' *
' * c-tree PLUS(tm)   Version 6.6
' *       Release A2
' *       January 30, 1996
' */
'-------------------------------------------------------------------------------


'INCLUDES PADRAO DO CTREE

'#include "ctstdr.h"    Estruturas do CTREE
'#include "ctoptn.h"
'#include "ctaerr.h"    Substituído pela sequencia abaixo de define
'                                        #define CTPERM
'                                        #include "ctstrc.h"
'                                        #include "ctgvar.h"
'#include "cterrc.h"    O codigo dos erros foi incluído e convertido para <Global Const> VB.
'                                        O texto dos erros foi íncluído como RESOURCE da DLL, use a função
'                                        ViewCtreeError() da DLL para apresentar um box de ERRMSG
'#include "ctdecl.h"    Foi incluído ao fim deste arquivo com compatibilização de
'                                        chamadas para Visual Basic

'---------------------------------------------------------------------------------------------------------
'INÍCIO DO cterrc.h

'#ifndef ctERRCH
'#define ctERRCH

'               /* USER ERROR CODES */

Global Const NO_ERROR = 0

Global Const KDUP_ERR = 2       '/* Key value already exists                  */

Global Const KMAT_ERR = 3       '/* Could not delete since pntr's don't match */
Global Const KDEL_ERR = 4       '/* Could not find key to delete              */
Global Const KBLD_ERR = 5       '/* Cannot call delete w/o verification with  */
                                '/* duplicate keys                            */
Global Const BJMP_ERR = 6       '/* c-tree(...) jump table error              */
Global Const TUSR_ERR = 7       '/* Terminate user                            */

Global Const SPAC_ERR = 10      '/* INTREE parameters require too much space  */
Global Const SPRM_ERR = 11      '/* Bad INTREE parameters                     */
Global Const FNOP_ERR = 12      '/* Could not open file: not there or locked  */
Global Const FUNK_ERR = 13      '/* Unknown file type                         */
Global Const FCRP_ERR = 14      '/* File corrupt at open                      */
Global Const FCMP_ERR = 15      '/* File has been compacted                   */
Global Const KCRAT_ERR = 16     '/* Could not create index file               */
Global Const DCRAT_ERR = 17     '/* Could not create data file                */
Global Const KOPN_ERR = 18      '/* Tried to create existing index file       */
Global Const DOPN_ERR = 19      '/* Tried to create existing data file        */
Global Const KMIN_ERR = 20      '/* Key length too large for node size        */
Global Const DREC_ERR = 21      '/* Record length too small                   */
Global Const FNUM_ERR = 22      '/* File number out of range                  */
Global Const KMEM_ERR = 23      '/* Illegal index member info                 */
Global Const FCLS_ERR = 24      '/* Could not close file                      */
Global Const KLNK_ERR = 25      '/* Bad link in deleted node list. REBUILD    */
Global Const FACS_ERR = 26      '/* File number not active                    */
Global Const LBOF_ERR = 27      '/* drn before beginning of data records      */
Global Const ZDRN_ERR = 28      '/* Zero drn in ADDKEY                        */
Global Const ZREC_ERR = 29      '/* Zero drn in data file routine             */
Global Const LEOF_ERR = 30      '/* drn exceeds logical end of file           */
Global Const DELFLG_ERR = 31    '/* Flag not set on record in delete chain    */
Global Const DDRN_ERR = 32      '/* Attempt to delete record twice in a row   */
Global Const DNUL_ERR = 33      '/* Attempt to use NULL ptr in read/write     */
Global Const PRDS_ERR = 34      '/* Predecessor repeat attempts exhausted     */
Global Const SEEK_ERR = 35      '/* Seek error:  check sysiocod value         */
Global Const READ_ERR = 36      '/* Read error:  check sysiocod error         */
Global Const WRITE_ERR = 37     '/* Write error: check sysiocod error         */
Global Const VRTO_ERR = 38      '/* Could not convert virtual open to actual  */
Global Const FULL_ERR = 39      '/* No more records availble                  */
Global Const KSIZ_ERR = 40      '/* Index node size too large                 */
Global Const UDLK_ERR = 41      '/* Could not unlock data record              */
Global Const DLOK_ERR = 42      '/* Could not obtain data record lock         */
Global Const FVER_ERR = 43      '/* Version incompatibility                   */
Global Const OSRL_ERR = 44      '/* Data file serial number overflow          */
Global Const KLEN_ERR = 45      '/* Key length exceeds MAXLEN parameter       */
Global Const FUSE_ERR = 46      '/* File number already in use                */
Global Const FINT_ERR = 47      '/* c-tree has not been initialized           */
Global Const FMOD_ERR = 48      '/* Operation incompatible with type of file  */
Global Const FSAV_ERR = 49      '/* Could not save file                       */
Global Const LNOD_ERR = 50      '/* Could not lock node                       */
Global Const UNOD_ERR = 51      '/* Could not unlock node                     */
Global Const KTYP_ERR = 52      '/* Variable length keys disabled OR invalid key type */
Global Const FTYP_ERR = 53      '/* File mode inconsistent with c-tree config */
Global Const REDF_ERR = 54      '/* Attempt to write a read only file         */
Global Const DLTF_ERR = 55      '/* File deletion failed                      */
Global Const DLTP_ERR = 56      '/* File must be opened exclusive for delete  */
Global Const DADV_ERR = 57      '/* Proper lock is not held (CHECKLOCK/READ)  */
Global Const KLOD_ERR = 58      '/* LOADKEY called with incorrect key number. */
                                '/* You cannot continue                       */
Global Const KLOR_ERR = 59      '/* LOADKEY called with key out of order      */
                                '/* You may skip this key & continue          */
Global Const KFRC_ERR = 60      '/* Percent out of range                      */
Global Const CTNL_ERR = 61      '/* NULL fcb detected during I/O              */
Global Const LERR_ERR = 62      '/* File must be opened exclusively           */
Global Const RSER_ERR = 63      '/* Start file / log file serial number error */
Global Const RLEN_ERR = 64      '/* File name too long in log file            */
Global Const RMEM_ERR = 65      '/* Not enough memory during tran processing  */
Global Const RCHK_ERR = 66      '/* Log file entry failed to find checkpoint  */
Global Const RENF_ERR = 67      '/* Could not rename file                     */
Global Const LALC_ERR = 68      '/* Could not allocate memory for control list*/
Global Const BNOD_ERR = 69      '/* Node does not belong to index             */
Global Const TEXS_ERR = 70      '/* Transaction already pending               */
Global Const TNON_ERR = 71      '/* No active transaction                     */
Global Const TSHD_ERR = 72      '/* No space for shadow buffer                */
Global Const TLOG_ERR = 73      '/* LOGFIL encountered during shadow only     */
Global Const TRAC_ERR = 74      '/* Recovery: two active tran for user        */
Global Const TROW_ERR = 75      '/* Recovery: bad tran owner                  */
Global Const TBAD_ERR = 76      '/* Recovery: bad tran type                   */
Global Const TRNM_ERR = 77      '/* Recovery: file name too long              */
Global Const TABN_ERR = 78      '/* Transaction abandoned: too many log extents*/
                                '/* or dynamic dump wait exhausted            */
Global Const FLOG_ERR = 79      '/* Could not log file opn/cre/cls/del        */
Global Const BKEY_ERR = 80      '/* NULL target or bad keyno                  */
Global Const ATRN_ERR = 81      '/* Transaction allocation error              */
Global Const UALC_ERR = 82      '/* User allocation error                     */
Global Const IALC_ERR = 83      '/* IASM allocation error                     */
Global Const MUSR_ERR = 84      '/* Maximum users exceeded                    */
Global Const LUPD_ERR = 85      '/* Reduce lock to read lock after update     */
Global Const DEAD_ERR = 86      '/* Dead lock detected                        */
Global Const QIET_ERR = 87      '/* System not quiet: files in use            */
Global Const LMEM_ERR = 88      '/* Linked list memory allocation error       */
Global Const TMEM_ERR = 89      '/* Memory allocation during tran processing  */
Global Const NQUE_ERR = 90      '/* Could not create queue                    */
Global Const QWRT_ERR = 91      '/* Queue write error                         */
Global Const QMRT_ERR = 92      '/* Queue memory error during write           */
Global Const QRED_ERR = 93      '/* Queue read error                          */
Global Const PNDG_ERR = 94      '/* Pending error: cannot save or commit tran */
Global Const STSK_ERR = 95      '/* Could not start task                      */
Global Const LOPN_ERR = 96      '/* Start-file/log open error                 */
Global Const SUSR_ERR = 97      '/* Bad user handle                           */
Global Const BTMD_ERR = 98      '/* Bad transaction mode                      */
Global Const TTYP_ERR = 99      '/* Transaction type / filmod conflict        */

Global Const ICUR_ERR = 100     '/* No current record for isam datno          */
Global Const INOT_ERR = 101     '/* Could not find isam keyno request         */
Global Const INOD_ERR = 102     '/* Could not open ISAM parameter file        */
Global Const IGIN_ERR = 103     '/* Could not read first 5 parameters in ISAM */
                                '/* parameter file                            */
Global Const IFIL_ERR = 104     '/* Too many files in ISAM parameter file     */
Global Const IUND_ERR = 105     '/* Could noy undo ISAM update. Rebuild Files */
Global Const IDRI_ERR = 106     '/* Could not read data file record in ISAM   */
                                '/* parameter file                            */
Global Const IDRK_ERR = 107     '/* Too many keys for data file in ISAM       */
                                '/* parameter file                            */
Global Const IMKY_ERR = 108     '/* Incorrect keyno for index member in       */
                                '/* parameter file                            */
Global Const IKRS_ERR = 109     '/* Too many key segments defined in ISAM     */
                                '/* parameter file                            */
Global Const ISRC_ERR = 110     '/* Could not read segment record in ISAM     */
                                '/*parameter file                             */
Global Const IKRI_ERR = 111     '/* Could not read index file record in ISAM  */
                                '/* parameter file                            */
Global Const IPND_ERR = 112     '/* LKISAM(ENABLE) found pending locks        */
Global Const INOL_ERR = 113     '/* No memory for user lock table             */
Global Const IRED_ERR = 114     '/* 1st byte of data record equals delete flag*/
                                '/*or bad variable length record mark         */
Global Const ISLN_ERR = 115     '/* Key segments do not match key length      */
Global Const IMOD_ERR = 116     '/* Bad mode parameter                        */
Global Const IMRI_ERR = 117     '/* Could not read index member record        */
Global Const SKEY_ERR = 118     '/* NXTSET called before FRSSET for keyno     */
Global Const SKTY_ERR = 119     '/* FRSSET called for index with wrong keytyp */

Global Const RRLN_ERR = 120     '/* Data record length exceeds rebuild max    */
Global Const RSPC_ERR = 121     '/* Not enough space for sort area            */
Global Const RMOD_ERR = 122     '/* Attempt to change fixed vs variable len   */
Global Const RVHD_ERR = 123     '/* Var length header has bad record mark     */
Global Const INIX_ERR = 124     '/* # of indices does not match (OPNIFIL)     */
Global Const IINT_ERR = 125     '/* c-tree already initialized                */

Global Const ABDR_ERR = 126     '/* Bad directory path get                    */
Global Const ARQS_ERR = 127     '/* Could not send request                    */
Global Const ARSP_ERR = 128     '/* Could not receive answer                  */
Global Const NINT_ERR = 129     '/* c-tree not initialized                    */
Global Const AFNM_ERR = 130     '/* Null file name pointer in OPNFIL          */
Global Const AFLN_ERR = 131     '/* File name length exceeds msg size         */
Global Const ASPC_ERR = 132     '/* No room for application message buffer    */
Global Const ASKY_ERR = 133     '/* Server is not active                      */
Global Const ASID_ERR = 134     '/* Could not get servers message id          */
Global Const AAID_ERR = 135     '/* Could not allocate application id         */
Global Const AMST_ERR = 136     '/* Could not get application msg status      */
Global Const AMQZ_ERR = 137     '/* Could not set message appl msg size       */
Global Const AMRD_ERR = 138     '/* Could not get rid of application msg      */
Global Const ABNM_ERR = 139     '/* Badly formed file name                    */
Global Const VMAX_ERR = 140     '/* Variable record length too long           */
Global Const AMSG_ERR = 141     '/* Required message size exceeds maximum     */

Global Const SMXL_ERR = 142     '/* Application MAXLEN > server's MAXLEN      */
Global Const SHND_ERR = 143     '/* Communications handler not installed      */
Global Const QMEM_ERR = 144     '/* Application could not id output queue     */
Global Const SCSF_ERR = 145     '/* Could not find COMM software              */
Global Const VDLK_ERR = 146     '/* Could not update free space info          */
Global Const VDLFLG_ERR = 147   '/* Space to be reused is not marked deleted  */
Global Const VLEN_ERR = 148     '/* WRTVREC cannot fit record at recbyt       */
Global Const VRLN_ERR = 149     '/* Varlen less than minimum in ADDVREC       */
Global Const SHUT_ERR = 150     '/* Server is shutting down                   */
Global Const STRN_ERR = 151     '/* Could not shut down. transactions pending */
Global Const LEXT_ERR = 152     '/* Could not extend logfile                  */
Global Const VBSZ_ERR = 153     '/* Buffer too small                          */
Global Const VRCL_ERR = 154     '/* Zero length record in REDVREC             */
Global Const SYST_ERR = 155     '/* Native system failure                     */
Global Const VFLG_ERR = 158     '/* REDVREC record not marked active          */
Global Const VPNT_ERR = 159     '/* Zero recbyt value                         */
Global Const ITIM_ERR = 160     '/* Multi-user interefernce: index information updated*/
                                '/* by the time user got to actual data record*/
Global Const SINA_ERR = 161     '/* User appears inactive                     */
Global Const SGON_ERR = 162     '/* Server has gone away                      */
Global Const SFRE_ERR = 163     '/* No more room in server lock table         */
Global Const SFIL_ERR = 164     '/* File number out of range                  */
Global Const SNFB_ERR = 165     '/* No file control block available           */
Global Const SNMC_ERR = 166     '/* No more ct file control blocks in server  */
Global Const SRQS_ERR = 167     '/* Could not read request                    */
Global Const SRSP_ERR = 168     '/* Could not send answer                     */
Global Const TCRE_ERR = 169     '/* Create file already opened (in recovery)  */

Global Const SFUN_ERR = 170     '/* Bad function number                       */
Global Const SMSG_ERR = 171     '/* Application msg size exceeds server size  */
Global Const SSPC_ERR = 172     '/* Could not allocate server msg buffer      */
Global Const SSKY_ERR = 173     '/* Could not identify server                 */
Global Const SSID_ERR = 174     '/* Could not get server message id           */
Global Const SAMS_ERR = 175     '/* Server could not allocate user msg area   */
Global Const SMST_ERR = 176     '/* Could not get server msg status           */
Global Const SMQZ_ERR = 177     '/* Could not set message server msg size     */
Global Const SINM_ERR = 178     '/* Unexpected file# assigned to [si] in rcv  */
Global Const SOUT_ERR = 179     '/* Server is at full user capacity           */


Global Const IKRU_ERR = 180     '/* Could not read symbolic key name          */
Global Const IKMU_ERR = 181     '/* Could not get mem for key symb name       */
Global Const IKSR_ERR = 182     '/* No room for sort key. increase MAXFIL     */
Global Const IDRU_ERR = 183     '/* Could not read file field number values   */
Global Const ISDP_ERR = 184     '/* Attempt to reallocate set space           */
Global Const ISAL_ERR = 185     '/* Not enough memory for multiple sets       */
Global Const ISNM_ERR = 186     '/* Set number out of range                   */
Global Const IRBF_ERR = 187     '/* Null buffer in rtread.c                   */
Global Const ITBF_ERR = 188     '/* Null target buffer in rtread.c            */
Global Const IJSK_ERR = 189     '/* Join_to skip                              */
Global Const IJER_ERR = 190     '/* Join_to error                             */
Global Const IJNL_ERR = 191     '/* Join_to null fill                         */
Global Const IDSK_ERR = 192     '/* Detail_for skip                           */
Global Const IDER_ERR = 193     '/* Detail_for error                          */
Global Const IDNL_ERR = 194     '/* Detail_for null fill                      */
Global Const IDMU_ERR = 195     '/* Could not get mem for dat symb name       */
Global Const ITML_ERR = 196     '/* Exceeded RETRY_LIMIT in RTREAD.C          */

Global Const IMEM_ERR = 197     '/* Could net get memory for ifil block       */
Global Const BIFL_ERR = 198     '/* Improper ifil block                       */
Global Const NSCH_ERR = 199     '/* Schema not defined for data file          */

Global Const RCRE_ERR = 400     '/* Resource already enabled                  */
Global Const RNON_ERR = 401     '/* Resources not enabled                     */
Global Const RXCL_ERR = 402     '/* File must be exclusive to enable res      */
Global Const RZRO_ERR = 403     '/* Empty resource id                         */
Global Const RBUF_ERR = 404     '/* Output buffer to small                    */
Global Const RDUP_ERR = 405     '/* Resource id already added                 */
Global Const RCSE_ERR = 406     '/* Bad resource search mode                  */
Global Const RRED_ERR = 407     '/* Attempt to get non-resource info          */
Global Const RNOT_ERR = 408     '/* Resource not found                        */
Global Const LKEP_ERR = 409     '/* Not in use: available                     */
Global Const USTP_ERR = 410     '/* User not active                           */
Global Const BSUP_ERR = 411     '/* Not a superfile                           */
Global Const LCIP_ERR = 412     '/* WRL to WXL commit promote pending(CIL)    */
Global Const SDIR_ERR = 413     '/* Superfile host not opened                 */
Global Const SNST_ERR = 414     '/* Cannot nest superfiles                    */
Global Const SADD_ERR = 415     '/* Illegal ADDKEY to superfile               */
Global Const SDEL_ERR = 416     '/* Illegal DELBLD to superfile               */
Global Const SPAG_ERR = 417     '/* Cache page size error                     */
Global Const SNAM_ERR = 418     '/* Max name inconsistency                    */
Global Const SRCV_ERR = 419     '/* Host superfile does not support recovery  */
Global Const TPND_ERR = 420     '/* Key update with pending transaction       */
Global Const BTFL_ERR = 421     '/* Filter not supported yet                  */
Global Const BTFN_ERR = 422     '/* Other functions not sup                   */
Global Const BTIC_ERR = 423     '/* Incomplete                                */
Global Const BTAD_ERR = 424     '/* Add list err                              */
Global Const BTIP_ERR = 425     '/* Batch in progress                         */
Global Const BTNO_ERR = 426     '/* No batch active                           */
Global Const BTST_ERR = 427     '/* Status info already returned              */
Global Const BTMT_ERR = 428     '/* No more info, batch cancelled             */
Global Const BTBZ_ERR = 429     '/* Bufsiz too small for record               */
Global Const BTRQ_ERR = 430     '/* Request is empty or inconsistent          */
Global Const LAGR_ERR = 431     '/* Aggregate/serialization lock denied       */
Global Const FLEN_ERR = 432     '/* Fixed length string requires len in DODA  */
Global Const SSCH_ERR = 433     '/* Segment def inconsistent with schema      */
Global Const DLEN_ERR = 434     '/* Very long def block not supported         */
Global Const FMEM_ERR = 435     '/* File def memory error                     */
Global Const DNUM_ERR = 436     '/* Bad def number                            */
Global Const DADR_ERR = 437     '/* defptr NULL during GETDEFBLK              */
Global Const DZRO_ERR = 438     '/* Requested def blk is empty                */
Global Const DCNV_ERR = 439     '/* No conversion routine for Definition Block*/
Global Const DDDM_ERR = 440     '/* Dynamic dump already in progress          */
Global Const DMEM_ERR = 441     '/* No memory for dynamic dump file buffer    */
Global Const DAVL_ERR = 442     '/* One or more files not available for dump  */
Global Const DSIZ_ERR = 443     '/* File length discrepancy                   */
Global Const DCRE_ERR = 444     '/* Could not create file during dump rcv     */
Global Const SDAT_ERR = 445     '/* Not enough data to assemble key value     */
Global Const BMOD_ERR = 446     '/* Bad key segment mode                      */
Global Const BOWN_ERR = 447     '/* Only the file's owner can perform op      */
Global Const DEFP_ERR = 448     '/* Permission to set file definition denied  */
Global Const DADM_ERR = 449     '/* ADMIN has opened file. Cannot delete file */
Global Const LUID_ERR = 450     '/* Invalid user id                           */
Global Const LPWD_ERR = 451     '/* Invalid password                          */
Global Const LSRV_ERR = 452     '/* Server could not process user/acct info   */
Global Const NSRV_ERR = 453     '/* No such server                            */
Global Const SGRP_ERR = 455     '/* User does not belong to group             */
Global Const SACS_ERR = 456     '/* Group access denied                       */
Global Const SPWD_ERR = 457     '/* File password invalid                     */
Global Const SWRT_ERR = 458     '/* Write permission not granted              */
Global Const SDLT_ERR = 459     '/* File delete permission denied             */
Global Const SRES_ERR = 460     '/* Resource not enabled                      */
Global Const SPER_ERR = 461     '/* Bad permission flag                       */
Global Const SHDR_ERR = 462     '/* No directory found in superfile           */
Global Const UQID_ERR = 463     '/* File id uniqueness error                  */
Global Const IISM_ERR = 464     '/* ISAM level logon not performed            */

Global Const IINI_ERR = 465     '/* Incremental Index: dnumidx < 1            */
Global Const IIDT_ERR = 466     '/* Incremental Index: dfilno not a ISAM file */
Global Const IINM_ERR = 467     '/* Incremental Index: aidxnam NULL for 1st   */
Global Const IITR_ERR = 468     '/* Incremental Index: active tran not allowed*/
Global Const NGIO_ERR = 469     '/* Negative I/O request                      */

Global Const SORT_ERR = 370     '/* Sort base: errors SORT_ERR + 101 thru 126 */
                                '/* see CTSORT.C or CTERRC.H for error listing*/

'/* Sort errors range from 471 thru 496 (e.g.: SORT_ERR + 101 thru 126).
'   These errors are returned from routines found in CTSORT.C.
'   Error codes and short descriptions returned by these routines are as
'   follows:
'
'471 = error deleting sortwork file
'472 = error creating unique name
'473 = error opening first dummy file
'474 = too few handles available min 3
'475 = error closing dummy file
'476 = error unlinking dummy file
'477 = error getting first data area
'478 = sinit phase not previously performed-srelease
'479 = sreturn phase already started
'480 = no records in data buffers
'481 = sint phase not previously performed-sreturn
'482 = not enough memory
'483 = no valid record pointers in merge buffers
'484 = error opening sortwork file
'485 = error creating sortwork.00x file
'486 = no records fit in output buffer
'487 = error reading sortwork file
'488 = bytes in buf <> merge buf size
'489 = error adjusting file pointer
'490 = error closing sortwork.00x
'491 = error closing sortwork file
'492 = error deleting sortwork file
'493 = error renaming sortwork.00x
'494 = error closing output file
'495 = error creating output file
'496 = insufficient disk space
'*/

Global Const NLOG_ERR = 498               '/* Old log file found during log create  */
Global Const FIDD_ERR = 499               '/* Mismatch between recv log & file id   */
Global Const SQLINIT_ERR = 500            '/* Server could not init SQL engine      */
Global Const SQLCONNECT_ERR = 501         '/* Could not init SQL for a user         */
Global Const SQL_REQUEST_ERROR = 502      '/* Could not access SQL master info      */
Global Const SQL_INVALID_CONTINUE = 503   '/* Could not continue SQL request        */
Global Const NSQL_ERR = 504               '/* Server does not support SQL           */
Global Const USQL_ERR = 505               '/* User profile does not enable SQL      */
Global Const SRFL_ERR = 506               '/* Could open save-restore file          */
Global Const SRIO_ERR = 507               '/* Could not process save-restore file   */
Global Const SRIN_ERR = 508               '/* Save restore inconsistency            */
Global Const DSRV_ERR = 509               '/* Duplicate server                      */
Global Const RFCK_ERR = 510               '/* Active chkpnt at start of roll-forward*/
Global Const ILOP_ERR = 511               '/* Index nodes form illegal loop: rebuild*/
Global Const DLOP_ERR = 512               '/* Data file loop detected               */
Global Const SBLF_ERR = 513               '/* FPUTFGET does not support CTSBLDX ()  */
Global Const CQUE_ERR = 514               '/* Queue has been closed                 */
Global Const OIFL_ERR = 515               '/* Cannot convert old IFIL structure     */
Global Const GNUL_ERR = 516               '/* ctNOGLOBALS not allocated             */
Global Const GNOT_ERR = 517               '/* 'regid' is not registered             */
Global Const GEXS_ERR = 518               '/* 'regid' is already registered         */

Global Const HTRN_ERR = 520               '/* Attempt to update index with          */
                                          '/* inconsistent tran#                    */
                                    
'-------------------------------------------------------------------------------------
'/*                             521 - 540 reserved for BANYAN env */
'#ifdef VINES
'#define BMAL_ERR               521 /* Could not allocate memory for the
'                               streettalk login message buffer */
'#define STID_ERR               522 /* Userid in INTISAM does not match
'                               current login id */
'#endif /* ifdef VINES */

'/* end of cterrc.h */
'#endif
'FIM DO cterrc.h
'-------------------------------------------------------------------------------------

'INCLUDES EXTRAS USADOS NA GERAÇÃO DA DLL
'#include "ctcmpl.h"               Só tinha abobrinha para geração da DLL
'#include "ctport.h"

'-------------------------------------------------------------------------------------
'FUNÇÕES NECESSÁRIAS PARA CTREE WINDOWS

'ctCONV pCTGVAR ctDECL GetCtreeGV( void );
Declare Function GetCtreeGV Lib "CTREESTD.DLL" () As Long
'ctCONV void ctDECL SetCtreeGV( pCTGVAR );
Declare Sub SetCtreeGV Lib "CTREESTD.DLL" (ByVal pCTGVAR As Long)
'ctCONV pCTGVAR ctDECL InitCtreeGV( void );
Declare Function InitCtreeGV Lib "CTREESTD.DLL" () As Long
'ctCONV int ctDECL ViewCtreeError( unsigned int );
Declare Function ViewCtreeError Lib "CTREESTD.DLL" (ByVal hWnd As Integer) As Integer


'#ifndef ctDECLH
'#define ctDECLH
'
'#include "ctifil.h"
'
'#ifdef PROTOTYPE
'
'#ifdef __cplusplus
'extern "C" {
'#endif
'
'extern ctCONV  COUNT ctDECL ADDKEY(COUNT keyno,pVOID target,LONG recbyt,COUNT typadd);
Declare Function ADDREC Lib "CTREESTD.DLL" (ByVal datno As Integer, recptr As Any) As Integer
'extern ctCONV  COUNT ctDECL ADDRES(COUNT datno,pVOID resptr,Integer varlen)
Declare Function ADDVREC Lib "CTREESTD.DLL" (ByVal datno As Integer, ByVal recptr As Any, ByVal varlen As Integer) As Integer
'extern ctCONV  COUNT ctDECL ALCBAT(COUNT numbat);
Declare Function ALCSET Lib "CTREESTD.DLL" (ByVal numset As Integer) As Integer
'extern ctCONV  COUNT ctDECL AVLFILNUM(COUNT numfils);
'extern ctCONV  COUNT ctDECL BATSET(COUNT filno,pVOID request,pVOID bufptr,Integer bufsiz,UCOUNT mode);
'extern ctCONV  COUNT ctDECL CHGBAT(COUNT batnum);
Declare Function CHGSET Lib "CTREESTD.DLL" (ByVal setnum As Integer) As Integer
'extern ctCONV  COUNT ctDECL CLIFIL(struct ifil ctMEM *ifilptr);
'extern ctCONV  COUNT ctDECL DELIFIL(struct ifil ctMEM *ifilptr);
'extern ctCONV  COUNT ctDECL CLNIDXX(pTEXT filnam,pTEXT fileword);
Declare Function CLISAM Lib "CTREESTD.DLL" () As Integer
'extern ctCONV  COUNT ctDECL CLRFIL(COUNT datno);
'extern ctCONV  COUNT ctDECL DELRFIL(COUNT datno);
'extern ctCONV  COUNT ctDECL CLSFIL(COUNT filno,COUNT filmod);
'extern ctCONV  COUNT ctDECL CREDAT(COUNT datno,pTEXT filnam,UCOUNT datlen,UCOUNT xtdsiz,COUNT filmod);
'extern ctCONV  COUNT ctDECL CREDATX(COUNT datno,pTEXT filnam,UCOUNT datlen,UCOUNT xtdsiz,COUNT filmod,LONG permmask,pTEXT groupid,pTEXT fileword);
'extern ctCONV  COUNT ctDECL CREIDX(COUNT keyno,pTEXT filnam,COUNT keylen,COUNT keytyp,COUNT keydup,COUNT nomemb,UCOUNT xtdsiz,COUNT filmod);
'extern ctCONV  COUNT ctDECL CREIDXX(COUNT keyno,pTEXT filnam,COUNT keylen,COUNT keytyp,COUNT keydup,COUNT nomemb,UCOUNT xtdsiz,COUNT filmod,LONG permmask,pTEXT groupid,pTEXT fileword);
'extern ctCONV  COUNT ctDECL CREIFIL(struct ifil ctMEM *ifilptr);
'extern ctCONV  COUNT ctDECL PRMIIDX(struct ifil ctMEM *ifilptr);
'extern ctCONV  COUNT ctDECL CREIFILX(struct ifil ctMEM *ifilptr,pTEXT dataextn,pTEXT indxextn,LONG permmask,pTEXT groupid,pTEXT fileword);
'extern ctCONV  COUNT ctDECL TMPIIDXX(struct ifil ctMEM *ifilptr,LONG permmask,pTEXT groupid,pTEXT fileword);
Declare Function CREISAM Lib "CTREESTD.DLL" (ByVal filnam As String) As Integer
'extern ctCONV  COUNT ctDECL CREISAMX(pTEXT filnam,COUNT userprof,pTEXT userid,pTEXT userword,pTEXT servname,LONG permmask,pTEXT groupid,pTEXT fileword);
'extern ctCONV  COUNT ctDECL CREMEM(COUNT keyno,COUNT keylen,COUNT keytyp,COUNT keydup,COUNT membno);
'extern ctCONV  COUNT ctDECL CTSBLDX(pTEXT filnam,pTEXT fileword);
'#ifdef ctNewSQL
'extern ctCONV  COUNT ctDECL CTSQL(pVOID ,pVOID ,Integer, COUNT ,LONG );
'#else
'extern ctCONV  COUNT ctDECL CTSQL(pVOID ,pVOID ,Integer );
'#endif /* ctNewSQL */
'extern ctCONV  COUNT ctDECL cttestfunc(VOID);
'extern ctCONV  COUNT ctDECL cttseg(COUNT segpos,COUNT mod,COUNT slen,pTEXT tarptr,pCOUNT aq,pConvMap mp);
'extern ctCONV  COUNT ctDECL ctuseg(COUNT spos,COUNT mod,COUNT slen,pTEXT tarptr,pCOUNT aq,pConvMap mp);
'extern ctCONV  LONG  ctDECL DATENT(COUNT datno);
Declare Function DATENT Lib "CTREESTD.DLL" (ByVal datno As Integer) As Long   '**** p/FF
'extern ctCONV  LONG  ctDECL DELBLD(COUNT keyno,pVOID target);
'extern ctCONV  COUNT ctDECL DELCHK(COUNT keyno,pVOID target,LONG recbyt);
'extern ctCONV  COUNT ctDECL DELFIL(COUNT filno);
Declare Function DELREC Lib "CTREESTD.DLL" (ByVal datno As Integer) As Integer
'extern ctCONV  COUNT ctDECL DELRES(COUNT datno,pVOID resptr);
Declare Function DELVREC Lib "CTREESTD.DLL" (ByVal datno As Integer) As Integer
'extern ctCONV  COUNT ctDECL ENARES(COUNT datno);
Declare Function EQLKEY Lib "CTREESTD.DLL" (ByVal keyno As Integer, ByVal target As Any) As Long
Declare Function EQLREC Lib "CTREESTD.DLL" (ByVal keyno As Integer, target As Any, recptr As Any) As Integer
'extern ctCONV  LONG  ctDECL ESTKEY(COUNT keyno,pVOID idxval1,pVOID idxval2);
'extern ctCONV  LONG  ctDECL FRCKEY(COUNT keyno,pVOID idxval,COUNT fractal);
'extern ctCONV  COUNT ctDECL FREBAT(VOID );
Declare Function FRESET Lib "CTREESTD.DLL" () As Integer
'extern ctCONV  COUNT ctDECL frmkey(COUNT keyno,pTEXT recptr,pTEXT txt,LONG pntr,Integer datlen);
'extern ctCONV  LONG  ctDECL FRSKEY(COUNT keyno,pVOID idxval);
Declare Function FRSREC Lib "CTREESTD.DLL" (ByVal filno As Integer, recptr As Any) As Integer
Declare Function FRSSET Lib "CTREESTD.DLL" (ByVal keyno As Integer, ByVal target As Any, recptr As Any, ByVal siglen As Integer) As Integer
'extern ctCONV  COUNT ctDECL GETALTSEQ(COUNT keyno,pCOUNT altseq);
'extern ctCONV  pTEXT ctDECL GETCURK(COUNT keyno,pVOID idxval,pInteger plen);
'extern ctCONV  pTEXT ctDECL GETCURKL(COUNT keyno,pVOID idxval);
Declare Function GETCURP Lib "CTREESTD.DLL" (ByVal datno As Integer) As Long
'extern ctCONV  Integer ctDECL GETDODA(COUNT datno,LONG buflen,pVOID bufptr,COUNT mode);
'extern ctCONV  LONG  ctDECL GETFIL(COUNT filno,COUNT mode);
'extern ctCONV  Integer ctDECL GETIFIL(COUNT datno,LONG buflen,pVOID bufptr);
'extern ctCONV  COUNT ctDECL GETMNAME(COUNT superFileNo, pTEXT nambuf, Integer buflen,COUNT memberFileNo);
'extern ctCONV  COUNT ctDECL GETNAM(COUNT filno,pTEXT nambuf,Integer buflen,COUNT mode);
'extern ctCONV  LONG  ctDECL GETRES(COUNT datno,pVOID resptr,pVOID bufptr,Integer bufsiz,COUNT resmode);
Declare Function GETVLEN Lib "CTREESTD.DLL" (ByVal datno As Integer) As Integer
'extern ctCONV  LONG  ctDECL GTEKEY(COUNT keyno,pVOID target,pVOID idxval);
Declare Function GTEREC Lib "CTREESTD.DLL" (ByVal keyno As Integer, target As Any, recptr As Any) As Integer
'extern ctCONV  LONG  ctDECL GTKEY(COUNT keyno,pVOID target,pVOID idxval);
'extern ctCONV  COUNT ctDECL GTREC(COUNT keyno,pVOID target,pVOID recptr);
'extern ctCONV  Integer ctDECL GTVLEN(COUNT datno,LONG recbyt);
'extern ctCONV  LONG  ctDECL IDXENT(COUNT keyno);
Declare Function INTISAM Lib "CTREESTD.DLL" (ByVal bufs As Integer, ByVal fils As Integer, ByVal sect As Integer) As Integer
'Ta Errado Precisa copiar        extern ctCONV  COUNT ctDECL INTISAMX(COUNT bufs,COUNT fils,COUNT sect);
'extern ctCONV  COUNT ctDECL INTREE(COUNT bufs,COUNT fils,COUNT sect);
'extern ctCONV  COUNT ctDECL INTREEX(COUNT bufs,COUNT fils,COUNT sect,COUNT dbufs,COUNT userprof,pTEXT userid,pTEXT userword,pTEXT servname);
'extern ctCONV  COUNT ctDECL IOPERFORMANCE(pVOID bufptr);
Declare Function LKISAM Lib "CTREESTD.DLL" (ByVal lokmod As Integer) As Integer
'extern ctCONV  COUNT ctDECL LOADKEY(COUNT keyno,pVOID target,LONG recbyt,COUNT typadd);
'extern ctCONV  COUNT ctDECL LOKREC(COUNT datno,COUNT lokmod,LONG recbyt);
'extern ctCONV  LONG  ctDECL LSTKEY(COUNT keyno,pVOID idxval);
Declare Function LSTREC Lib "CTREESTD.DLL" (ByVal filno As Integer, recptr As Any) As Integer
Declare Function LSTSET Lib "CTREESTD.DLL" (ByVal keyno As Integer, ByVal target As Any, recptr As Any, ByVal siglen As Integer) As Integer
'extern ctCONV  LONG  ctDECL LTEKEY(COUNT keyno,pVOID target,pVOID idxval);
Declare Function LTEREC Lib "CTREESTD.DLL" (ByVal keyno As Integer, ByVal target As String, recptr As String) As Integer
'extern ctCONV  LONG  ctDECL LTKEY(COUNT keyno,pVOID target,pVOID idxval);
'extern ctCONV  COUNT ctDECL LTREC(COUNT keyno,pVOID target,pVOID recptr);
'extern ctCONV  COUNT ctDECL MIDSET(COUNT keyno,pVOID recptr,COUNT siglen);
'extern ctCONV  LONG  ctDECL NEWREC(COUNT datno);
'extern ctCONV  LONG  ctDECL NEWVREC(COUNT datno,Integer varlen);
'extern ctCONV  pTEXT ctDECL NXTCTREE(VOID );
'extern ctCONV  LONG  ctDECL NXTKEY(COUNT keyno,pVOID idxval);
Declare Function NXTREC Lib "CTREESTD.DLL" (ByVal filno As Integer, recptr As Any) As Integer
Declare Function NXTSET Lib "CTREESTD.DLL" (ByVal keyno As Integer, recptr As Any) As Integer
'extern ctCONV  COUNT ctDECL OPNFIL(COUNT ufilno,pTEXT filnam,COUNT filmod);
'extern ctCONV  COUNT ctDECL OPNFILX(COUNT filno,pTEXT filnam,COUNT filmod,pTEXT fileword);
'extern ctCONV  COUNT ctDECL OPNIFIL(struct ifil ctMEM *ifilptr);
'extern ctCONV  COUNT ctDECL OPNIFILX(struct ifil ctMEM *ifilptr,pTEXT dataextn,pTEXT indxextn,pTEXT fileword);
Declare Function OPNISAM Lib "CTREESTD.DLL" (ByVal filnam As String) As Integer
'extern ctCONV  COUNT ctDECL OPNISAMX(pTEXT filnam,COUNT userprof,pTEXT userid,pTEXT userword,pTEXT servname,pTEXT fileword);
'extern ctCONV  COUNT ctDECL OPNRFIL(COUNT filno,pTEXT filnam,COUNT filmod);
'extern ctCONV  COUNT ctDECL OPNRFILX(COUNT filno,pTEXT filnam,COUNT filmod,pTEXT fileword);
'extern ctCONV  LONG  ctDECL ORDKEY(COUNT keyno,pVOID target,Integer offset,pVOID idxval);
'extern ctCONV  LONG  ctDECL PRVKEY(COUNT keyno,pVOID idxval);
Declare Function PRVREC Lib "CTREESTD.DLL" (ByVal filno As Integer, recptr As Any) As Integer
Declare Function PRVSET Lib "CTREESTD.DLL" (ByVal keyno As Integer, recptr As Any) As Integer
'extern ctCONV  COUNT ctDECL PUTDODA(COUNT datno,pDATOBJ doda,UCOUNT numfld);
'extern ctCONV  COUNT ctDECL PUTFIL(COUNT filno,COUNT filmod);
'extern ctCONV  COUNT ctDECL PUTIFIL(struct ifil ctMEM *ifilptr);
'extern ctCONV  COUNT ctDECL PUTIFILX(struct ifil ctMEM *ifilptr,pTEXT dataextn,pTEXT indxextn,pTEXT fileword);
'extern ctCONV  COUNT ctDECL TRANA(VOID );
'extern ctCONV  COUNT ctDECL TRANCLR(VOID );
'extern ctCONV  LONG  ctDECL TSTVREC(COUNT datno,Integer varlen);
'extern ctCONV  COUNT ctDECL RBLIFIL(struct ifil ctMEM *ifilptr);
'extern ctCONV  COUNT ctDECL RBLIFILX(struct ifil ctMEM *ifilptr,pTEXT dataextn,pTEXT indxextn,LONG permmask,pTEXT groupid,pTEXT fileword);
'extern ctCONV  COUNT ctDECL RDVREC(COUNT datno,LONG recbyt,pVOID recptr,Integer bufsiz);
'extern ctCONV  COUNT ctDECL REDREC(COUNT datno,LONG recbyt,pVOID recptr);
Declare Function REDVREC Lib "CTREESTD.DLL" (ByVal datno As Integer, ByVal recptr As String, ByVal bufsiz As Integer) As Integer
'extern ctCONV  COUNT ctDECL REGCTREE(pTEXT regid);
'extern ctCONV  COUNT ctDECL RETREC(COUNT datno,LONG recbyt);
'extern ctCONV  COUNT ctDECL RETVREC(COUNT datno,LONG recbyt);
'extern ctCONV  LONG  ctDECL RNGENT(COUNT keyno, pVOID key1, pVOID key2);
Declare Function RRDREC Lib "CTREESTD.DLL" (ByVal datno As Integer, recptr As Any) As Integer
Declare Function RWTREC Lib "CTREESTD.DLL" (ByVal datno As Integer, recptr As Any) As Integer
Declare Function RWTVREC Lib "CTREESTD.DLL" (ByVal datno As Integer, ByVal recptr As String, ByVal varlen As Integer) As Integer
'extern ctCONV  COUNT ctDECL SECURITY(COUNT filno,pVOID bufptr,Integer bufsiz,COUNT mode);
'extern ctCONV  LONG  ctDECL SERIALNUM(COUNT datno);
'extern ctCONV  COUNT ctDECL SETALTSEQ(COUNT keyno,pCOUNT altseq);
Declare Function SETCURI Lib "CTREESTD.DLL" (ByVal datno As Integer, ByVal recbyt As Long, ByVal recptr As String, ByVal datlen As Integer) As Integer
'extern ctCONV  COUNT ctDECL SETDEFBLK(COUNT datno,pDEFDEF defp);
'extern ctCONV  COUNT ctDECL SETVARBYTS(COUNT filno,pUTEXT pvbyte);
'#ifndef CTBOUND
'extern ctCONV  COUNT ctDECL SETSFLV(COUNT mode, UTEXT value);
'extern ctCONV  COUNT ctDECL SETWNAM(pTEXT pName);
'#endif
'#ifdef ctNewSQL
'extern ctCONV  COUNT ctDECL SQR(pVOID bufptr,Integer bufsiz,COUNT funcId,LONG cursorNumber);
'#else
'extern ctCONV  COUNT ctDECL SQR(pVOID bufptr,Integer bufsiz);
'#endif
'extern ctCONV  COUNT ctDECL STPSRV(pTEXT userword,pTEXT servname,COUNT delay);
'extern ctCONV  COUNT ctDECL STPUSR(VOID );
'extern ctCONV  COUNT ctDECL SWTCTREE(pTEXT regid);
Declare Function TFRMKEY Lib "CTREESTD.DLL" (ByVal keyno As Integer, target As Any) As Integer
'extern ctCONV  COUNT ctDECL TMPNAME(pTEXT bufptr,Integer bufsiz);
'extern ctCONV  LONG  ctDECL TRANBEG(COUNT mode);
'extern ctCONV  COUNT ctDECL TRANEND(COUNT mode);
'extern ctCONV  COUNT ctDECL TRANRST(COUNT savpnt);
'extern ctCONV  COUNT ctDECL TRANSAV(VOID );
'extern ctCONV  COUNT ctDECL SPCLSAV(VOID );
'extern ctCONV  COUNT ctDECL UNRCTREE(pTEXT regid);
Declare Function UPDCURI Lib "CTREESTD.DLL" (ByVal datno As Integer, ByVal mode As Integer) As Integer
'extern ctCONV  COUNT ctDECL UPDRES(COUNT datno,pVOID resptr,Integer varlen);
'extern ctCONV  pTEXT ctDECL WCHCTREE(VOID );
'extern ctCONV  COUNT ctDECL WRTREC(COUNT datno,LONG recbyt,pVOID recptr);
'extern ctCONV  COUNT ctDECL WRTVREC(COUNT datno,LONG recbyt,pVOID recptr,Integer varlen);
'extern ctCONV  COUNT ctDECL setfndval(pVOID bufptr,Integer bufsiz);
'#ifdef ctreeVARLST
'extern ctCONV  COUNT ctVDECL ctree();
'#else
'extern ctCONV COUNT ctVDECL ctree(COUNT fn,COUNT filno,pVOID ptr1,pLONG plong,pVOID ptr2,pInteger psize,COUNT mode);
'#endif
'
'#ifdef __cplusplus
'}
'#endif
'
'#endif /* ctDECLH */

'/* end of ctdecl.h */

