@echo off

rem
rem ���̃o�b�`�̐���
rem

rem �ݒ莖��
set HOGE="�ϐ��̒l"

rem ���̃o�b�`�����݂���t�H���_���J�����g��
pushd %0\..
cls

START CSCRIPT vacuum_vba.js

pause
exit