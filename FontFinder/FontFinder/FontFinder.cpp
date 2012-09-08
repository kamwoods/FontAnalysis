// FontFinder.cpp : Defines the entry point for the console application.
//

#include <Windows.H>
#include <StdIO.H>
#include <ConIO.H>

	
static FILE *out ;



int CALLBACK EnumFontFamiliesExProc( ENUMLOGFONTEX *lpelfe, NEWTEXTMETRICEX *lpntme, int FontType, LPARAM lParam )

{
        fwprintf( out, L"%s\n", lpelfe->elfFullName );
        return 1;
}


int main(int argc, char* argv[])
{
        HDC hDC = GetDC( NULL );
		errno_t   err ;

		err = _wfopen_s(&out, L"font_dump.txt", L"w, ccs=UNICODE" ) ;

        LOGFONT lf = { 0, 0, 0, 0, 0, 0, 0, 0, DEFAULT_CHARSET, 0, 0, 0,
         0, NULL };
         EnumFontFamiliesEx( hDC, &lf, (FONTENUMPROC)EnumFontFamiliesExProc, 0, 0 );
        ReleaseDC( NULL, hDC );
	    fclose( out ); 
        return 0;
}
