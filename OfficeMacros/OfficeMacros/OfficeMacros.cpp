// OfficeMacros.cpp : main project file.

#include "stdafx.h"

   #using <mscorlib.dll>

   #using "Office.dll"
   #using "Microsoft.Office.Interop.word.dll"
   #using "Microsoft.Office.Interop.excel.dll"
   #using "Microsoft.Office.Interop.powerpoint.dll"
   #using "Microsoft.Office.Interop.Access.dll"
   #using "Microsoft.vbe.interop.dll"

   using namespace System;
   //using namespace System::Diagnostics;
   using namespace System::Reflection;
   using namespace Microsoft::Office::Core;
   using namespace Microsoft::Office::Interop;

   #include <tchar.h>

   void PrintMenu();
   void AutoCallAccess();
   void AutoCallExcel();
   void AutoCallPowerPoint();
   void AutoCallWord();
   void CallMacro(Object^ oApp, Object^ oArgs[]);

int main(array<System::String ^> ^args)
   {
   	PrintMenu();
   	System::String ^ s = Console::ReadLine();
   	while( !s->ToLower()->Equals(S"q") )
   	{
   		Int32 i;
   		try
   		{
   			i = Convert::ToInt32(s,10);
   		}
   		catch( Exception^ e )
   		{
   			goto print;
   		}

   		// Select the Office application to automate based on user input.

   		switch( i )
   		{
/*   		case 1:
   			AutoCallAccess();
   			break;
   		case 2:
   			AutoCallExcel();
   			break;
   		case 3:
   			AutoCallPowerPoint();
   			break;

*/
			case 4:
   			AutoCallWord();
   			break;
   		default:
   			;
   		}
   print:
   		PrintMenu();
   		s = Console::ReadLine();
   	}
   	return 0;
   }

   void PrintMenu()
   {
   	Console::WriteLine(S"\n\nEnter the number of the application you'd like to automate.");
   	Console::WriteLine(S"Enter 'q' to quit the application.\n");
   	Console::WriteLine(S"\t\t1. Microsoft Access");
   	Console::WriteLine(S"\t\t2. Microsoft Excel");
   	Console::WriteLine(S"\t\t3. Microsoft PowerPoint");
   	Console::WriteLine(S"\t\t4. Microsoft Word\n");
   	Console::Write(S"\tSelection:");
   }

   void AutoCallAccess()
   {
   	try{
   		//Start Access, make it visible, and open C:\Db1.mdb.
   		Console::WriteLine("\nStarting Microsoft Access...");
   		Access::ApplicationClass^ pAccess = gcnew Access::ApplicationClass();
   		pAccess->Visible = true;
   		pAccess->OpenCurrentDatabase("c:\\db1.mdb", false, "");
   		//Run the macros.
   		array<System::Object^, 3> oParams = {gcnew String("DoKbTest"), 
   			System::Reflection::Missing::Value};
   		CallMacro(pAccess, oParams);
   		oParams[0] = new String("DoKbTestWithParameter");
   		oParams[1] = new String("Hello From Visual C++ .NET (AutoCallAccess)");
   		CallMacro(pAccess, oParams);	
   		//Quit Access and clean up.
        pAccess->get_DoCmd()->Quit(Access::AcQuitOption::acQuitSaveNone);
   		System::Runtime::InteropServices::Marshal::ReleaseComObject(pAccess);
   		GC::Collect();
   	}
   	catch(Exception^ e)
   	{
   		Console::WriteLine(S"Error automating Access...");
   		Console::WriteLine(e->get_Message());
   	}
   }
/*
   void AutoCallExcel()
   {
   	try{
   		//Start Excel, make it visible, and open C:\Book1.xls.
   		System::Object* oMissing = System::Reflection::Missing::Value;
   		Console::WriteLine("\nStarting Microsoft Excel...");
        Excel::ApplicationClass* pExcel = new Excel::ApplicationClass();
   		pExcel->Visible = true;
        Excel::Workbooks* pBooks = pExcel->get_Workbooks();
   		Excel::_Workbook* pBook = pBooks->Open("c:\\book1.xls", oMissing, oMissing,
   			oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, 
   			oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

   		//Run the macros.
   		System::Object* oParams[] = {new String("DoKbTest"), oMissing};
   		CallMacro(pExcel, oParams);
   		oParams[0] = new String("DoKbTestWithParameter");
   		oParams[1] = new String("Hello From Visual C++ .NET (AutoCallExcel)");
   		CallMacro(pExcel, oParams);
   		//Quit Excel and clean up.
   		pBook->Close(false, oMissing, oMissing);
   		System::Runtime::InteropServices::Marshal::ReleaseComObject(pBook);
   		System::Runtime::InteropServices::Marshal::ReleaseComObject(pBooks);
   		pExcel->Quit();
   		System::Runtime::InteropServices::Marshal::ReleaseComObject(pExcel);
   		GC::Collect();
   	}
   	catch(Exception* e)
   	{
   		Console::WriteLine(S"Error automating Excel...");
   		Console::WriteLine(e->get_Message());
   	}
   }

   void AutoCallPowerPoint()
   {
   	try{
   		//Start PowerPoint, make it visible, and open C:\Pres1.ppt.
   		Console::WriteLine("\nStarting Microsoft PowerPoint...");
   		PowerPoint::ApplicationClass* pPPT = new PowerPoint::ApplicationClass();
        pPPT->Visible =  Microsoft::Office::Core::MsoTriState::msoTrue;
   		
   		PowerPoint::Presentations* pPresSet = pPPT->get_Presentations();
   		PowerPoint::_Presentation* pPres = pPresSet->Open("C:\\pres1.ppt", 
   			MsoTriState::msoFalse, 
   			MsoTriState::msoFalse,
   			MsoTriState::msoTrue);
   		
   		//Run the macros.
   		System::Object* oParams[] = {new String("'pres1.ppt'!DoKbTest"), 
   			System::Reflection::Missing::Value};
   		CallMacro(pPPT, oParams);
   		oParams[0] = new String("'pres1.ppt'!DoKbTestWithParameter");
   		oParams[1] = new String("Hello From Visual C++ .NET (AutoCallPowerPoint)");
   		CallMacro(pPPT, oParams);		
   		//Quit PowerPoint and clean up.
   		pPres->Close();
   		System::Runtime::InteropServices::Marshal::ReleaseComObject(pPres);
   		System::Runtime::InteropServices::Marshal::ReleaseComObject(pPresSet);
   		pPPT->Quit();
   		System::Runtime::InteropServices::Marshal::ReleaseComObject(pPPT);
   		GC::Collect();
   	}
   	catch(Exception* e)
   	{
   		Console::WriteLine(S"Error automating PowerPoint...");
   		Console::WriteLine(e->get_Message());
   	}
   }
*/
   void AutoCallWord()
   {
   	try{
   		//Start Word, make it visible, and open C:\Doc1.doc.
   		System::Object* oMissing = System::Reflection::Missing::Value;
   		Console::WriteLine("\nStarting Microsoft Word...");
   		Word::ApplicationClass* pWord = new Word::ApplicationClass;
   		pWord->Visible = true;
   		Word::Documents* pDocs = pWord->Documents;
   		System::Object* oFile = new System::Object;
   		oFile = S"c:\\doc1.doc";
   		Word::_Document* pDoc = pDocs->Open(&oFile, &oMissing,
   			&oMissing, &oMissing, &oMissing, &oMissing, &oMissing,
   			&oMissing, &oMissing, &oMissing, &oMissing, &oMissing,
   			&oMissing, &oMissing, &oMissing);

   		//Run the macros.
   		System::Object* oParams[] = {new String("DoKbTest"), oMissing};
   		CallMacro(pWord, oParams);
   		oParams[0] = new String("DoKbTestWithParameter");
   		oParams[1] = new String("Hello From Visual C++ .NET (AutoCallWord)");
   		CallMacro(pWord, oParams);		
   		//Quit Word and clean up.
   		pDoc->Close(&oMissing, &oMissing, &oMissing);
   		System::Runtime::InteropServices::Marshal::ReleaseComObject(pDoc);
   		System::Runtime::InteropServices::Marshal::ReleaseComObject(pDocs);
   		pWord->Quit(&oMissing, &oMissing, &oMissing);
   		System::Runtime::InteropServices::Marshal::ReleaseComObject(pWord);
   		GC::Collect();
   	}
   	catch(Exception* e)
   	{
   		Console::WriteLine(S"Error automating Word...");
   		Console::WriteLine(e->get_Message());
   	}
   }

   void CallMacro(Object* oApp, Object* oArgs[])
   {
   	Console::WriteLine("Calling Macro...");
   	oApp->GetType()->InvokeMember("Run",
   		BindingFlags(BindingFlags::Default | BindingFlags::InvokeMethod),
   		NULL, oApp, oArgs);

   }
		