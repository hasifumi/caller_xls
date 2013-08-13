/*
	Excelからマクロの全モジュールを外部に吸い出すバッチ
*/


// excelのファイル名と，モジュールの吸出し先パスを指定
var file_dir = "C:\\Users\\fumio\\MyProject\\excel_vba\\loadableModules";
var file_name = "template.xls";
var file_path = file_dir + "\\" + file_name;
var vacuum_dir = file_dir + "\\macros";


// ブックを開く
var excel = WScript.CreateObject("Excel.Application");
excel.Visible = true;
excel.Workbooks.Open( file_path );
var book = excel.Workbooks( excel.Workbooks.Count ); 
	// JScript/WSH で，Excelファイルを読み書きしよう
	// http://d.hatena.ne.jp/language_and_engineering/20090717/p1


// ブック内のマクロの全モジュールをスキャンする
var cnt_module = 0;
var e = new Enumerator( book.VBProject.VBComponents );
for( ; ! e.atEnd() ;  e.moveNext() )
{
	// モジュールを取得
	var vba_module = e.item();
	
	// このモジュールの名前を取得
	var module_name = vba_module.Name;
		//WScript.Echo( module_name );
	
	// このモジュールのエクスポート先パスを決定
	var bas_path = vacuum_dir + "\\" + module_name + ".bas";
	
	// このモジュールをエクスポート
	vba_module.Export( bas_path );
		
	cnt_module ++;
}


// Excelを閉じて終了
excel.DisplayAlerts = false;
excel.Quit();
excel = null;

WScript.Echo(
	file_name 
	+ " から，全 " 
	+ cnt_module 
	+ " 個のモジュールを " 
	+ vacuum_dir 
	+ " 上に吸い出しました。" 
);
