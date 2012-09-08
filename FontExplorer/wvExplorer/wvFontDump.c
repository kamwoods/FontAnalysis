#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include <stdlib.h>
#include <stdio.h>
#include <errno.h>
#include <string.h>
#include "iconv.h"
#include <signal.h>
//#include <sys/time.h>

#ifdef HAVE_UNISTD_H
#include <unistd.h>
#endif

#include <time.h>
#include "wv.h"

static iconv_t converter;


void on_alarm(int signum)
{
  printf("Error : Timed out\n");
  exit(-1);
}

int fontnumber;
char *fname;


int
main (int argc, char *argv[])
{
    FILE *input;
    int i,j,k;
    wvParseStruct ps;
    char *dir = NULL;
    wvVersion ver;

    //struct sigaction newhandler;
    //sigset_t blocked;
    //struct itimerval itimer;

    /* limit time */
/*
    newhandler.sa_handler = on_alarm;
    if (sigaction(SIGALRM, &newhandler, NULL) == -1)
      {
	printf("Error: couldn't set handler\n");
	exit(-1);
      }

    itimer.it_value.tv_sec = 10;
    itimer.it_value.tv_usec = 0;
    setitimer(ITIMER_REAL, &itimer, NULL);

    if (argc < 2)
      {
	fprintf(stderr, "usage %s filename\n", argv[0]);
	exit(-1);
      }
*/
	
	fname = "c:\\TestDocs\\Test.doc"; //argv[1];

    if ((converter = iconv_open("utf-8", "utf-16LE")) < 0)
      return -1;

    input = fopen (fname, "rb");
    if (!input)
      {
	  fprintf (stderr, "Failed to open %s: %s\n", fname, strerror (errno));
	  return -1;
      }
    fclose (input);
    wvInit ();

    printf("File: %s\n", fname);
	i = wvInitParser (&ps, fname);
    ps.filename = fname;
	ps.dir = "" ; //dir;

	ver = wvQuerySupported(&ps.fib, NULL);

	/* get font list */
    /* associated strings */

    if ((ver == WORD6)
        || (ver == WORD7))
      {
	wvGetFFN_STTBF6 (&ps.fonts, ps.fib.fcSttbfffn, ps.fib.lcbSttbfffn,
			 ps.tablefd);
	wvGetSTTBF6(&ps.anSttbfAssoc, ps.fib.fcSttbfAssoc, ps.fib.lcbSttbfbkmk, ps.tablefd);
      }
    else
      {
	wvGetFFN_STTBF (&ps.fonts, ps.fib.fcSttbfffn, ps.fib.lcbSttbfffn,
			ps.tablefd);
	wvGetSTTBF(&ps.anSttbfAssoc, ps.fib.fcSttbfAssoc, ps.fib.lcbSttbfAssoc, ps.tablefd);
      }


    for (i = 0; i < ps.fonts.nostrings; i++)
      {
	char  *fname;
	char  *buf;

	char *inptr =  (char *) ps.fonts.ffn[i].xszFfn;
	size_t inbytes;
	size_t outbytes;

	for (j = 0; ps.fonts.ffn[i].xszFfn[j]; j++);
	if (j == 0)
	  continue;
	inbytes = j * 2;
	outbytes = j * 4;
	fname = (char *) malloc(outbytes + 1);
	buf = fname;

	if (iconv(converter, &inptr, &inbytes, &buf, &outbytes))
	  {
	    printf("iconv problem %s\n", strerror(errno));
	    free(fname);
	    continue;
	  }


	buf[0] = 0;
	if (fname)
	  {
	    char *altname;
	    printf("Font:%s\n",fname);
	    altname = wvGetFontnameFromCode(&ps.fonts, ps.fonts.ffn[i].ixchSzAlt);
	    printf("Alternative Name:%s\n", altname);
	    //free(altname);
	    printf("Panose:");
	    for (k = 0; k < 10; k++)
	      printf(" %d", ((unsigned char *) &ps.fonts.ffn[i].panose)[k]);
	    printf("\nSignature:");
	    for (k = 0; k < 6; k++)
	      printf(" %x", ((int *) &ps.fonts.ffn[i].fs)[k]);
	    printf("\nPitch:%d\n", ps.fonts.ffn[i].prq);
	    printf("Family:%d\n", ps.fonts.ffn[i].ff);
	    printf("Weight:%d\n", ps.fonts.ffn[i].wWeight);
	    printf("Charset:%d\n", ps.fonts.ffn[i].chs);
	    printf("Truetype:%s\n", ps.fonts.ffn[i].fTrueType ? "True" : "False");

	  }
      }
}
