//
//  Values are 32 bit values layed out as follows:
//
//   3 3 2 2 2 2 2 2 2 2 2 2 1 1 1 1 1 1 1 1 1 1
//   1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0
//  +---+-+-+-----------------------+-------------------------------+
//  |Sev|C|R|     Facility          |               Code            |
//  +---+-+-+-----------------------+-------------------------------+
//
//  where
//
//      Sev - is the severity code
//
//          00 - Success
//          01 - Informational
//          10 - Warning
//          11 - Error
//
//      C - is the Customer code flag
//
//      R - is a reserved bit
//
//      Facility - is the facility code
//
//      Code - is the facility's status code
//
//
// Define the facility codes
//


//
// Define the severity codes
//


//
// MessageId: EVMSG_INSTALLED
//
// MessageText:
//
//  O servico %1 foi instalado.
//
#define EVMSG_INSTALLED                  0x00000064L

//
// MessageId: EVMSG_REMOVED
//
// MessageText:
//
//  O servico %1 foi removido.
//
#define EVMSG_REMOVED                    0x00000065L

//
// MessageId: EVMSG_NOTREMOVED
//
// MessageText:
//
//  O servico %1 nao pode ser removido.
//
#define EVMSG_NOTREMOVED                 0x00000066L

//
// MessageId: EVMSG_CTRLHANDLERNOTINSTALLED
//
// MessageText:
//
//  O gerenciador de mensagens nao pode ser instalado.
//
#define EVMSG_CTRLHANDLERNOTINSTALLED    0x00000067L

//
// MessageId: EVMSG_FAILEDINIT
//
// MessageText:
//
//  O processo de inicializacao falhou.
//
#define EVMSG_FAILEDINIT                 0x00000068L

//
// MessageId: EVMSG_STARTED
//
// MessageText:
//
//  O servico %1 foi iniciado.
//
#define EVMSG_STARTED                    0x00000069L

//
// MessageId: EVMSG_BADREQUEST
//
// MessageText:
//
//  O servico recebeu uma chamada invalida.
//
#define EVMSG_BADREQUEST                 0x0000006AL

//
// MessageId: EVMSG_DEBUG
//
// MessageText:
//
//  Debug: %1
//
#define EVMSG_DEBUG                      0x0000006BL

//
// MessageId: EVMSG_STOPPED
//
// MessageText:
//
//  O servico foi interropido.
//
#define EVMSG_STOPPED                    0x0000006CL

//
// MessageId: EVMSG_SHUTDOWN
//
// MessageText:
//
//  O Servico %1 foi finalizado pelo usuario.
//
#define EVMSG_SHUTDOWN                   0x0000006DL

//
// MessageId: EVMSG_CONTINUE
//
// MessageText:
//
//  O Servico %1 foi reinicializado pelo usuario.
//
#define EVMSG_CONTINUE                   0x0000006EL

//
// MessageId: EVMSG_GENERIC_ERROR
//
// MessageText:
//
//  Descrição do Erro: %1.
//
#define EVMSG_GENERIC_ERROR              0x0000006FL

