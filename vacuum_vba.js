/*
	Excel����}�N���̑S���W���[�����O���ɋz���o���o�b�`
*/


// excel�̃t�@�C�����ƁC���W���[���̋z�o����p�X���w��
var file_dir = "C:\\Users\\fumio\\MyProject\\excel_vba\\loadableModules";
var file_name = "template.xls";
var file_path = file_dir + "\\" + file_name;
var vacuum_dir = file_dir + "\\macros";


// �u�b�N���J��
var excel = WScript.CreateObject("Excel.Application");
excel.Visible = true;
excel.Workbooks.Open( file_path );
var book = excel.Workbooks( excel.Workbooks.Count ); 
	// JScript/WSH �ŁCExcel�t�@�C����ǂݏ������悤
	// http://d.hatena.ne.jp/language_and_engineering/20090717/p1


// �u�b�N���̃}�N���̑S���W���[�����X�L��������
var cnt_module = 0;
var e = new Enumerator( book.VBProject.VBComponents );
for( ; ! e.atEnd() ;  e.moveNext() )
{
	// ���W���[�����擾
	var vba_module = e.item();
	
	// ���̃��W���[���̖��O���擾
	var module_name = vba_module.Name;
		//WScript.Echo( module_name );
	
	// ���̃��W���[���̃G�N�X�|�[�g��p�X������
	var bas_path = vacuum_dir + "\\" + module_name + ".bas";
	
	// ���̃��W���[�����G�N�X�|�[�g
	vba_module.Export( bas_path );
		
	cnt_module ++;
}


// Excel����ďI��
excel.DisplayAlerts = false;
excel.Quit();
excel = null;

WScript.Echo(
	file_name 
	+ " ����C�S " 
	+ cnt_module 
	+ " �̃��W���[���� " 
	+ vacuum_dir 
	+ " ��ɋz���o���܂����B" 
);
