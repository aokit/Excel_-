Attribute VB_Name = "x2p_Lib1"
' -*- coding:shift_jis-dos -*-

'./x2p_Lib1.bas
' �����Ӂ�
' �@���̃X�N���v�g�ł͊֐���`���Ă���̂ŕW�����W���[���ɒu���Ȃ��Ƌ@�\���Ȃ��B
' �@���̂��߃t�@�C���̐擪�Ɂ@Attribute VB_Name = "x2p_Lib1"�@�ƋL�q���Ă����B
' �@�i�O���t�@�C���Ƃ��ĕҏW���Ă���Ƃ��̂݌�����j

Function �Z���̌Œ�F(�Z��)
     Dim a
     '// a = �Z��.Interior.ColorIndex
     '// a = �Z��.FormatConditions.Interior.Color ' - ����͂��܂��s���Ȃ�
     '// �@�@�����t���w�i�F�𓾂����Ƃ��Ă�
     '// �@�@���[�U�֐��Ƃ��Ď��s����ƃG���[���N����̂Ŏg���Ȃ�
     '// a = �Z��.DisplayFormat.Interior.Color

     a = �Z��.Interior.ColorIndex
     �Z���̌Œ�F = a
End Function
