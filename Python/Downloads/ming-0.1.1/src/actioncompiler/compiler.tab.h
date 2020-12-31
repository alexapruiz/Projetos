typedef union {
  Buffer action;
  char *str;
  SWFActionFunction function;
  SWFGetUrl2Method getURLMethod;
} YYSTYPE;
#define	BREAK	257
#define	FOR	258
#define	CONTINUE	259
#define	IF	260
#define	ELSE	261
#define	DO	262
#define	WHILE	263
#define	THIS	264
#define	EVAL	265
#define	TIME	266
#define	RANDOM	267
#define	LENGTH	268
#define	INT	269
#define	CONCAT	270
#define	DUPLICATECLIP	271
#define	REMOVECLIP	272
#define	TRACE	273
#define	STARTDRAG	274
#define	STOPDRAG	275
#define	ORD	276
#define	CHR	277
#define	CALLFRAME	278
#define	GETURL	279
#define	GETURL1	280
#define	LOADMOVIE	281
#define	LOADVARIABLES	282
#define	POSTURL	283
#define	SUBSTR	284
#define	NEXTFRAME	285
#define	PREVFRAME	286
#define	PLAY	287
#define	STOP	288
#define	TOGGLEQUALITY	289
#define	STOPSOUNDS	290
#define	GOTOFRAME	291
#define	FRAMELOADED	292
#define	SETTARGET	293
#define	STRING	294
#define	NUMBER	295
#define	IDENTIFIER	296
#define	GETURL_METHOD	297
#define	EQ	298
#define	LE	299
#define	GE	300
#define	NE	301
#define	LAN	302
#define	LOR	303
#define	INC	304
#define	DEC	305
#define	IEQ	306
#define	DEQ	307
#define	MEQ	308
#define	SEQ	309
#define	STREQ	310
#define	STRNE	311
#define	STRCMP	312
#define	PARENT	313
#define	END	314
#define	UMINUS	315
#define	POSTFIX	316
#define	NEGATE	317


extern YYSTYPE yylval;
