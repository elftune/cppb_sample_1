//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Unit1.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
	: TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Button2Click(TObject *Sender)
{
	/*
		プロジェクト＞オプションで　ビルド後イベントに　copy $(ProjectDir)\*.xls* $(OUTPUTDIR)　を登録しておく
		これで、プロジェクトのフォルダにあるExcelファイルが実行ファイルのあるフォルダに自動コピーされる
		C++ Builderでは $(ProjectDir) のあとに \ が必要。VS(Visual Studio 20xx)では不要
		出力先は VSでは $(TargetDir)　C++ Builderでは$(OUTPUTDIR)

		フォルダをコピーしたい場合、VSではこうする(Dataというフォルダ)
		xcopy $(ProjectDir)Data $(TargetDir)Data /D/E/C/I/H/Y
	*/

	Variant ExcelApp = Variant::CreateObject("Excel.Application");
	ExcelApp.OlePropertySet("Visible", false); // 念のため
	ExcelApp.OlePropertySet("DisplayAlerts", false); // 念のため
	UnicodeString excelfile = ExtractFilePath(Application->ExeName) + Edit1->Text;


	Variant book = ExcelApp.OlePropertyGet("WorkBooks").OleFunction("Open", (OleVariant)excelfile);
	UnicodeString macro = "SampleMacro";
	Variant v = ExcelApp.OlePropertyGet("Application").OleFunction("Run", (OleVariant)macro,(OleVariant)"test.csv");

	// 配列0番目が整数、1番目が文字列と事前に分かっているのでこれで取得できる
	int i = v.GetElement(0);
	UnicodeString s = v.GetElement(1);
	ShowMessage(s);

	ExcelApp.OleProcedure("Quit");
	v = Unassigned;
	book = Unassigned;
	ExcelApp = Unassigned;
}
//---------------------------------------------------------------------------

