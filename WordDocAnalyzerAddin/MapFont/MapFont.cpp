// This is the main DLL file.
#include "stdafx.h"

#include <Windows.H>
#include <StdIO.H>
#include <ConIO.H>

#include "resource.h"
//#include "MapFont.h"

//using namespace std;



namespace Map_Font{

	class MapFont{
	void map_font( wchar_t* font_path)
	{
        HDC hDC = GetDC( NULL );		
		HFONT	hfont;
        TEXTMETRIC          tm;
        LOGFONT             lf;
		TCHAR				szFaceName[LF_FACESIZE];
		CHOOSEFONT			choose_font;

		LPCSTR path;
		int err;
		bool success;
		
        // load font 
		err = AddFontResource(font_path);
		SendMessage(HWND_BROADCAST,WM_FONTCHANGE,0,0);

		choose_font.lStructSize = sizeof(CHOOSEFONT);
		choose_font.hwndOwner = NULL;
		choose_font.lpLogFont = &lf;
		choose_font.Flags = CF_SCREENFONTS;
		success = ChooseFont( &choose_font );


		// unload font

		err = RemoveFontResource(font_path);
		SendMessage(HWND_BROADCAST,WM_FONTCHANGE,0,0);

		hfont = CreateFontIndirect(&lf);
		SelectObject(hDC, hfont);
		GetObject( GetCurrentObject( hDC, OBJ_FONT ), 
					sizeof(lf), 
					(LPVOID) &lf );
		GetTextFace( hDC, sizeof(szFaceName), szFaceName );
		GetTextMetrics( hDC, &tm );
		//cout << "Font: " << "Arial" << "  Mapping: " << szFaceName << endl;
		
		ReleaseDC( NULL, hDC );

        return;
	}

	};


}