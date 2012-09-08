// FontMapping.cpp : Defines the entry point for the console application.
//

#include <Windows.H>
#include <StdIO.H>
#include <ConIO.H>
#include <iostream>
#include <gdiplus.h>

using namespace std ;
using namespace Gdiplus;

static FILE *out ;


class MapFont{

public:

	void map_font()
	{
        HDC hDC = GetDC( NULL );
		HFONT				hfont;
        TEXTMETRIC          tm;
        LOGFONT             lf;
		TCHAR				szFaceName[LF_FACESIZE];
		CHOOSEFONT			choose_font;
		OPENFILENAME		ofn;
		LPCTSTR lpszFilename ;
		int err;
		bool success;
		wchar_t szFile[260];       // buffer for file name

	while( true )
	{

		// Initialize OPENFILENAME
		ZeroMemory(&ofn, sizeof(ofn));
		ofn.lStructSize = sizeof(ofn);
		ofn.hwndOwner = NULL;
		ofn.lpstrFile = szFile;
		//
		// Set lpstrFile[0] to '\0' so that GetOpenFileName does not 
		// use the contents of szFile to initialize itself.
		//
		ofn.lpstrFile[0] = '\0';
		ofn.nMaxFile = sizeof(szFile);
		ofn.lpstrFilter = L"Font Files\0*.ttf\0";
		ofn.nFilterIndex = 1;
		ofn.lpstrFileTitle = NULL;
		ofn.nMaxFileTitle = 0;
		ofn.lpstrInitialDir = NULL;
		ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

		// Display the Open dialog box. 

		success = GetOpenFileName(&ofn);
		lpszFilename = szFile;
        // load font 

		err = AddFontResource(lpszFilename);
		SendMessage(HWND_BROADCAST,WM_FONTCHANGE,0,0);
		// invoke the Font Chooser Dialog

		choose_font.lStructSize = sizeof(CHOOSEFONT);
		choose_font.hwndOwner = NULL;
		choose_font.lpLogFont = &lf;
		choose_font.Flags = CF_SCREENFONTS;
		success = ChooseFont( &choose_font );


wcout<<"-----------------Returned from ChooseFont----------------------------"<<endl;
		wcout << "Font File: " <<lpszFilename<< endl<<endl;
		wcout<<"Logical Font:"<<endl;
		wcout<<"lfFaceName: "<<lf.lfFaceName<<endl;
		wcout<<"lfHeight: "<<lf.lfHeight<<endl;
		wcout<<"lfWidth: "<<lf.lfWidth<<endl;
		wcout<<"lfEscapement: "<<lf.lfEscapement<<endl;
		wcout<<"lfOrientation: "<<lf.lfOrientation<<endl;
		wcout<<"lfWeight: "<<lf.lfWeight<<endl;
		wcout<<"lfItalic: "<<lf.lfItalic<<endl;
		wcout<<"lfUnderline: "<<lf.lfUnderline<<endl;
		wcout<<"lfStrikeOut: "<<lf.lfStrikeOut<<endl;
		wcout<<"lfCharSet: "<<lf.lfCharSet<<endl;
		wcout<<"lfOutPrecision: "<<lf.lfOutPrecision<<endl;
		wcout<<"lfClipPrecision: "<<lf.lfClipPrecision<<endl;
		wcout<<"lfQuality: "<<lf.lfQuality<<endl;
		wcout<<"lfPitchAndFamily: "<<lf.lfPitchAndFamily<<endl;
wcout<<"---------------------------------------------------------------------"<<endl;
		wcout<<endl<<endl;



		// unload font

		err = RemoveFontResource(lpszFilename);
		SendMessage(HWND_BROADCAST,WM_FONTCHANGE,0,0);


		// create new font and retrieve information

		hfont = CreateFontIndirect(&lf);
		SelectObject(hDC, hfont);
		//GetObject( GetCurrentObject( hDC, OBJ_FONT ), 
					//sizeof(lf), 
					//(LPVOID) &lf );
		GetTextFace( hDC, sizeof(szFaceName), szFaceName );
		GetTextMetrics( hDC, &tm );
/*		
		wcout << "Font File: " <<lpszFilename<< endl<<endl;
		wcout<<"Logical Font:"<<endl;
		wcout<<"lfFaceName: "<<lf.lfFaceName<<endl;
		wcout<<"lfHeight: "<<lf.lfHeight<<endl;
		wcout<<"lfWidth: "<<lf.lfWidth<<endl;
		wcout<<"lfEscapement: "<<lf.lfEscapement<<endl;
		wcout<<"lfOrientation: "<<lf.lfOrientation<<endl;
		wcout<<"lfWeight: "<<lf.lfWeight<<endl;
		wcout<<"lfItalic: "<<lf.lfItalic<<endl;
		wcout<<"lfUnderline: "<<lf.lfUnderline<<endl;
		wcout<<"lfStrikeOut: "<<lf.lfStrikeOut<<endl;
		wcout<<"lfCharSet: "<<lf.lfCharSet<<endl;
		wcout<<"lfOutPrecision: "<<lf.lfOutPrecision<<endl;
		wcout<<"lfClipPrecision: "<<lf.lfClipPrecision<<endl;
		wcout<<"lfQuality: "<<lf.lfQuality<<endl;
		wcout<<"lfPitchAndFamily: "<<lf.lfPitchAndFamily<<endl;
		wcout<<endl<<endl;
*/
		wcout<<"Text Face: " << szFaceName << endl<<endl;
		wcout<<"Text Metrics:"<<endl;
		wcout<<"tmAscent: "<<tm.tmAscent<<endl;
		wcout<<"tmDescent: "<<tm.tmDescent<<endl;
		wcout<<"tmInternalLeading: "<<tm.tmInternalLeading<<endl;
		wcout<<"tmExternalLeading: "<<tm.tmExternalLeading<<endl;
		wcout<<"tmAveCharWidth: "<<tm.tmAveCharWidth<<endl;
		wcout<<"tmMaxCharWidth: "<<tm.tmMaxCharWidth<<endl;
		wcout<<"tmWeight: "<<tm.tmWeight<<endl;
		wcout<<"tmOverhang: "<<tm.tmOverhang<<endl;
		wcout<<"tmDigitizedAspectX: "<<tm.tmDigitizedAspectX<<endl;
		wcout<<"tmDigitizedAspectY: "<<tm.tmDigitizedAspectY<<endl;
		wcout<<"tmFirstChar: "<<(int)tm.tmFirstChar<<endl;
		wcout<<"tmLastChar: "<<(int)tm.tmLastChar<<endl;
		wcout<<"tmDefaultChar: "<<(int)tm.tmDefaultChar<<endl;
		wcout<<"tmBreakChar: "<<(int)tm.tmBreakChar<<endl;
		wcout<<"tmItalic: "<<tm.tmItalic<<endl;
		wcout<<"tmUnderlined: "<<tm.tmUnderlined<<endl;
		wcout<<"tmStruckOut: "<<tm.tmStruckOut<<endl;
		wcout<<"tmPitchAndFamily: "<<tm.tmPitchAndFamily<<endl;
		wcout<<"tmCharSet: "<<tm.tmCharSet<<endl;		
		wcout<<endl;

		char ch;

		cin>>ch;
		if( ch == 'q' )
			break;

	}

		ReleaseDC( NULL, hDC );

        return;
	}
};


void  main( int argc, char* argv[])
{
	MapFont *map = new MapFont();

	map->map_font();
	return;
}
