#include <stdio.h>
#include <string.h>
#include <windows.h>

long WINAPI Encripta(long lngIn, char *strOut);
void wiro76 (char *strIn, char *strOut);


long WINAPI Encripta(long lngIn, char *strOut)
{
    char ent[9];
    char ret[9];
    char saida[17];
    short i;
    char chAux;

    //Garante range 0-99999999
    if (lngIn<0L) lngIn = 0L;
    if (lngIn>99999999L) lngIn = 99999999L;
    
    //Formata em string
    sprintf (ent, "%08ld", lngIn);

    //Converte para EBCDIC
    for (i=0; i<8; i++)
    {
        ent[i]+=(char)192;
    }

    //Chama a rotina de criptografia
    wiro76(ent, ret);

    //Converte saida para hexadecimal (dígito a dígito)
    for (i=0; i<8; i++)
    {
        chAux = ((ret[i] & 0xF0) >> 4);
        saida[i*2] = ((chAux)<10) ? ('0'+chAux) : ('A'+chAux-10);

        chAux = (ret[i] & 0x0F);
        saida[i*2+1] = ((chAux)<10) ? ('0'+chAux) : ('A'+chAux-10);
    }
    saida[16] = 0;

    //Copia o resultado final
    strcpy (strOut, saida);

    return 1;
}

