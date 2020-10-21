'ThisWorkbook Module

' 2020/10/18 12:04 シートを指定してVBAモジュールを積み込めるようにした。
' libdef.txt に記載しておく。
' 2020/10/20 12:45 BAS コードをインポートするモジュールをコード内で明示的に
' 指定していないときにはデフォルトで Module1 が使われる。これに対応した。
' シートの作成が間に合わないためエラーが起きている。
' 仮説：シートを追加しているので、Worksheet.Activate()がキックされてしまっている。
' この段階では、追加してもキックしないようにするか、副作用を消さないとだめだ。
' 2020/10/21 19:55 シートを Add しても CodeName が無いためのインポート失敗を解決。
' CodeNameが有効になるまでに、VBProject へのアクセスが最低１回必要だった、とのこと。
' ---
' https://social.msdn.microsoft.com/Forums/expression/ja-JP
' /d8862af1-4790-44a8-a6b8-89fc9681f7af/excel-12398-codename
' ?forum=officesupportteamja
'      Dim tmp As String : tmp = ThisWorkbook.VBProject.Name


' ==== https://blog.goo.ne.jp/end-u/e/b1e41190650cbb1ff2c70bafe92a6dce ====
' = 1/3 =
Option Explicit
Dim flg     As Boolean
' Dim shCount As Long
Public shCount As Long
Public booted As Boolean
Dim shName  As String
' booting = True
' =========================================================================


' Text Scripting on VBA v1.0.0
' last update: 2013-01-03
' HATANO Hirokazu
'
' Detail  :  http://rsh.csh.sh/text-scripting-vba/
' See Also:  http://d.hatena.ne.jp/language_and_engineering/20090731/p1

' Option Explicit


'----------------------------- Consts ---------------

'ライブラリリストの設定 (設置フォルダはワークブックと同じディレクトリ)
Const FILENAME_LIBLIST As String = "libdef.txt" 'ライブラリリストのファイル名
Const FILENAME_EXPORT As String = "ThisWorkbook-sjis.cls" 'エクスポート clsファイル名

'ワークブック オープン時に実行する(True) / しない(False)
Const ENABLE_WORKBOOK_OPEN As Boolean = True
'Const ENABLE_WORKBOOK_OPEN As Boolean = False

'ショートカットキー
Const SHORTKEY_RELOAD As String = "^r" 'ctrl + r


'----------------------------- Workbook_open() ---------------

'ワークブック オープン時に実行
Private Sub Workbook_Open()
  shCount = Sheets.count ' = 2/3 =' ==== https://blog.goo.ne.jp/... ====
  If ENABLE_WORKBOOK_OPEN = False Then
    Exit Sub
  End If
  
  Call setShortKey
  Call reloadModule
  ' オープン直後は『直前』のワークシート数が現在のシート数と一致はしないが
  ' 『ワークシートの追加』ではないので識別するために flg をセットしておく。
  flg = True
  shCount = Sheets.count
  Debug.Print shCount
  ' MsgBox "ブックのシート数：" & shCount

End Sub


'ワークブック クローズ時に実行
Private Sub Workbook_BeforeClose(Cancel As Boolean)
  Call clearShortKey
End Sub


'----------------------------- public Subs/Functions ---------------

Public Sub reloadModule()

  '手動リロード用 Public関数
  
  Dim msgError As String
  msgError = loadModule("." & Application.PathSeparator & FILENAME_LIBLIST)
  
  If Len(msgError) > 0 Then
    MsgBox msgError
  End If
End Sub


Public Sub exportThisWorkbook()
  'ThisWorkbook 手動export用 Public関数
  Call exportModule("ThisWorkbook", FILENAME_EXPORT)
End Sub


'----------------------------- main Subs/Functions ---------------

Private Function loadModule(ByVal pathConf As String) As String
  'Main: モジュールリストファイルに書いてある外部ライブラリを読み込む。

  '1. 全モジュールを削除
  Dim isClear As Boolean
  isClear = clearModules
  
  If isClear = False Then
    loadModule = "Error: 標準モジュールの全削除に失敗しました。"
    Exit Function
  End If
  
  
  '2. モジュールリストファイルの存在確認
  ' 2.1. モジュールリストファイルの絶対パスを取得
  pathConf = absPath(pathConf)
  
  ' 2.2. 存在チェック
  Dim isExistList As Boolean
  isExistList = checkExistFile(pathConf)
  
  If isExistList = False Then
    loadModule = "Error: ライブラリリスト" & pathConf & "が存在しません。"
    Exit Function
  End If


  '3. モジュールリストファイルの読み込み&配列化
  Dim arrayModules As Variant
  arrayModules = list2array(pathConf)
  
  If UBound(arrayModules) = 0 Then
    loadModule = "Error: ライブラリリストに有効なモジュールの記述が存在しません。"
    Exit Function
  End If

  
  '4. 各モジュールファイル読み込み
  Dim i As Integer
  Dim msgError As String
  msgError = ""
  
  ' 配列は0始まり。(最大値: 配列個数-1)
  For i = 0 To UBound(arrayModules) - 1
    Dim pathModule As String
    pathModule = arrayModules(i)
    
    '4.1. モジュールリストファイルの存在確認
    ' 4.1.1. モジュールリストファイルの絶対パスを取得
    pathModule = absPath(pathModule)
  
    ' 4.1.2. 存在チェック
    Dim isExistModule As Boolean
    isExistModule = checkExistFile(pathModule)
  
    '4.2. モジュール読み込み
    If isExistModule = True Then
      ' ThisWorkbook.VBProject.VBComponents.Import pathModule
      importModule (pathModule)
    Else
      msgError = msgError & pathModule & " は存在しません。" & vbCrLf
    End If
  Next i
  loadModule = msgError
  
End Function


'----------------------------- Functions / Subs ---------------

' シートモジュールとして外部ファイルに記載したスクリプトをシートを
' 指定してインポートする
' 
Private Sub importModule(ByVal pathModule As String)
  Dim aArr As Variant
  aArr = Split(pathModule, "》")
  If UBound(aArr) = 0 Then
    ThisWorkbook.VBProject.VBComponents.Import pathModule
  Else
    ' pathModule：G:\user01\Github\RECT_tsv\FixValue》値固定.bas
    ' aArr(0)   ：G:\user01\Github\RECT_tsv\FixValue
    ' aArr(1)   ：値固定.bas
    Dim aArr2 As Variant
    aArr2 = Split(aArr(1), ".")
    ' aArr2(0)：値固定
    Dim aSheet As Worksheet
    Set aSheet = Nothing
    On Error Resume Next ' エラーが発生しても次の行から実行.
    Set aSheet = ActiveWorkbook.Worksheets(aArr2(0))
    On Error GoTo 0 ' On Error Resume Next を使用して有効にしたエラー処理を無効にする.
    If aSheet Is Nothing Then
      ' シートがない場合の処理
      Debug.Print "モジュールのインポート指定先のシートを作成します：" & aArr2(0) 
      MsgBox "モジュールのインポート指定先のシートを作成します：" & aArr2(0) 
      Set aSheet = ActiveWorkbook.Worksheets.Add
      ' あらかじめ『コードを表示』しておくと、ここでシートができ、シートの名前がつき、
      ' CodeNameも準備されたのがわかる。ところが、それなしにファイルをダブルクリック
      ' で開くと、シートの名前がつくところまでは確認できるが、CodeNameがないままと
      ' なっている。
      ' https://social.msdn.microsoft.com/Forums/expression/ja-JP
      ' /d8862af1-4790-44a8-a6b8-89fc9681f7af/excel-12398-codename
      ' ?forum=officesupportteamja
      Dim tmp As String : tmp = ThisWorkbook.VBProject.Name
      aSheet.Name = aArr2(0)
      While asheet.CodeName = ""
         MsgBox "インポート指定先のシートが無名" & asheet.CodeName
      Wend
      Debug.Print "モジュールのインポート指定先のシートができました：" & aSheet.CodeName
      MsgBox "モジュールのインポート指定先のシートができました：" & aSheet.CodeName
    Else
      ' シートがある場合（ほんとはそのシートのモジュールを全部消してから）
      With ThisWorkbook.VBProject.VBComponents(aSheet.CodeName).CodeModule
        ' .DeleteLines 7
        Debug.Print .CountOfLines
        .DeleteLines 1, .CountOfLines
      End With
    End If
    ' ThisWorkbook.VBProject.VBComponents.(sCN).CodeModule.Import aArr(1)
    importModule2 pathModule, aSheet.CodeName
  End If
End Sub

Private Sub importModule2(ByVal pathModule As String, ByVal CodeName As String)
    ' まず、ここでいったん標準モジュールの直下にインポート
    Debug.Print pathModule
    ThisWorkbook.VBProject.VBComponents.Import (pathModule)
    ' "G:\user01\Github\RECT_tsv\FixValue》値固定.bas"
    ' ThisWorkbook.VBProject.VBComponents(sCN).CodeModule
    ' Debug.Print aInsert.Name

    ' インポート先の検出：
    ' VBComponents.Import によるインポートではシート名を指定することができない。
    ' インポート先は標準モジュール配下のオブジェクトとなりその名前は、通常は
    ' スクリプトの中のAttributeで、以下のように指定される。しかし、Attribute
    ' でも、シート名を指定することまではできないので、結局スクリプトのファイル名
    ' で指定する。結果として、インポート先となる標準モジュール配下のオブジェクト
    ' 名もファイル名から抽出する。
    ' Attribute VB_Name = "任意の識別子》シート名"
    ' ただし、上記の指定がないと、ファイル名に指定してあっても、そのオブジェクト
    ' 名は、 Module1 になる（上書きされる）
    ' 《逆に：Attribute VB_Name = "tmp" などと共通にしておいて、デフォルトの
    ' 　Module1 とは当たらないようにしておくのも手かもしれない。また、
    ' 　あらかじめモジュール名のコレクションを持っておいて、 .Importしたあとの
    ' 　差分を確認して、インポートされたモジュールを特定する方法もあるだろう》
    Dim aArr As Variant
    Dim aArr2 As Variant
    aArr = Split(pathModule, "\")
    aArr2 = Split(aArr(UBound(aArr)), ".")
    ' pathModule        ：G:\user01\Github\RECT_tsv\FixValue》値固定.bas
    ' aArr(UBound(aArr))：FixValue》値固定.bas
    ' aArr2(0)          ：FixValue》値固定　←これがオブジェクト名
    
    Dim Code As String
    Debug.Print CodeName
    Debug.Print aArr2(0) ' 標準モジュールの下のコード
    
    ' デフォルトでインポートされる先は Module1
    Dim aModule As Variant
    On Error Resume Next ' エラーが発生しても次の行から実行.
    Set aModule = ThisWorkbook.VBProject.VBComponents(aArr2(0))
    Set aModule = ThisWorkbook.VBProject.VBComponents("Module1")
    On Error GoTo 0 ' On Error Resume Next を使用して有効にしたエラー処理を無効にする.

    ' 移動したいシートモジュール
    Debug.Print CodeName
    MsgBox "つぎのシートにモジュールをコピーします：" & CodeName
    Dim asModule As Variant
    Set asModule = ThisWorkbook.VBProject.VBComponents(CodeName)

    ' 標準モジュールからシートモジュールへのコピー
    
    With aModule.CodeModule
        Code = .Lines(1, .CountOfLines)
    End With
    With asModule.CodeModule
        .AddFromString Code
    End With
    Application.VBE.ActiveVBProject.VBComponents.Remove aModule
    
End Sub



Private Sub exportModule(ByVal nameModule As String, ByVal nameFile As String)

  Dim component As Object
  For Each component In ThisWorkbook.VBProject.VBComponents
    
    If component.Name = nameModule Then
      component.Export ThisWorkbook.Path & Application.PathSeparator & nameFile
      MsgBox nameModule & " を " & ThisWorkbook.Path & Application.PathSeparator & nameFile & " として保存しました。"
    End If
    
  Next component

End Sub


'----------------------------- common Functions / Subs ---------------
Private Function clearModules() As Boolean
  '標準モジュール/クラスモジュール初期化(全削除)
  
  Dim component As Object
  For Each component In ThisWorkbook.VBProject.VBComponents
      
    '標準モジュール(Type=1) / クラスモジュール(Type=2)を全て削除
    If component.Type = 1 Or component.Type = 2 Then
      ThisWorkbook.VBProject.VBComponents.Remove component
    End If
    
  Next component
  
  '標準モジュール/クラスモジュールの合計数が0であればOK
  Dim cntBAS As Long
  cntBAS = countBAS()
  
  Dim cntClass As Long
  cntClass = countClasses()
        
  If cntBAS = 0 And cntClass = 0 Then
    clearModules = True
  Else
    clearModules = False
  End If

End Function


Private Function countBAS() As Long
  Dim count As Long
  count = countComponents(1) 'Type 1: bas
  countBAS = count
End Function


Private Function countClasses() As Long
  Dim count As Long
  count = countComponents(2) 'Type 2: class
  countClasses = count
End Function


Private Function countComponents(ByVal numType As Integer) As Long
  '存在する標準モジュール/クラスモジュールの数を数える
  
  Dim i As Long
  Dim count As Long
  count = 0
  
  With ThisWorkbook.VBProject
    For i = 1 To .VBComponents.count
      If .VBComponents(i).Type = numType Then
        count = count + 1
      End If
    Next i
  End With

  countComponents = count
End Function


Private Function absPath(ByVal pathFile As String) As String
  ' ファイルパスを絶対パスに変換
  
  Dim nameOS As String
  nameOS = Application.OperatingSystem
  
  'replace Win backslash(Chr(92))
  pathFile = Replace(pathFile, Chr(92), Application.PathSeparator)
  
  'replace Mac ":"Chr(58)
  pathFile = Replace(pathFile, ":", Application.PathSeparator)
  
  'replace Unix "/"Chr(47)
  pathFile = Replace(pathFile, "/", Application.PathSeparator)


  Select Case Left(pathFile, 1)
  
    'Case1. . で始まる場合(相対指定)
    Case ".":
  
      Select Case Left(pathFile, 2)
        
        ' Case1-1. 相対指定 "../" 対応
        Case "..":
          'MsgBox "Case1-1: " & pathFile
          absPath = ThisWorkbook.Path & Application.PathSeparator & pathFile
          Exit Function
    
        ' Case1-2. 相対指定 "./" 対応
        Case Else:
          'MsgBox "Case1-2: " & pathFile
          absPath = ThisWorkbook.Path & Mid(pathFile, 2, Len(pathFile) - 1)
          Exit Function
    
      End Select
    
    'Case2. 区切り文字で始まる場合 (絶対指定)
    Case Application.PathSeparator:
    
      ' Case2-1. Windows Network Drive ( chr(92) & chr(92) & "hoge")
      'MsgBox "Case2-1: " & pathFile
      If Left(pathFile, 2) = Chr(92) & Chr(92) Then
        absPath = pathFile
        Exit Function
      
      Else
      ' Case2-2. Mac/UNIX Absolute path (/hoge)
        absPath = pathFile
        Exit Function
      
      End If
    
  End Select


  'Case3. [A-z][0-9]で始まる場合 (Mac版Officeで正規表現が使えれば select文に入れるべき...)

  ' Case3-1.ドライブレター対応("c:" & chr(92) が "c" & chr(92) & chr(92)になってしまうので書き戻す)
  If nameOS Like "Windows *" And Left(pathFile, 2) Like "[A-z]" & Application.PathSeparator Then
    'MsgBox "Case3-1" & pathFile
    absPath = Replace(pathFile, Application.PathSeparator, ":", 1, 1)
    Exit Function
  End If
 
  ' Case3-2. 無指定 "filename"対応
  If Left(pathFile, 1) Like "[0-9]" Or Left(pathFile, 1) Like "[A-z]" Then
    absPath = ThisWorkbook.Path & Application.PathSeparator & pathFile
    Exit Function
  Else
    MsgBox "Error[AbsPath]: fail to get absolute path."
  
  End If

End Function


Private Function checkExistFile(ByVal pathFile As String) As Boolean

  On Error GoTo Err_dir
  If Dir(pathFile) = "" Then
    checkExistFile = False
  Else
    checkExistFile = True
  End If

  Exit Function

Err_dir:
  checkExistFile = False

End Function


'リストファイルを配列で返す(行頭が'(コメント)の行 & 空行は無視する)
Private Function list2array(ByVal pathFile As String) As Variant
    
  Dim nameOS As String
  nameOS = Application.OperatingSystem
        
  '1. リストファイルの読み取り
  Dim fp As Integer
  fp = FreeFile
  Open pathFile For Input As #fp
  
  '2. リストの配列化
  Dim arrayOutput() As String
  Dim countLine As Integer
  countLine = 0
  ReDim Preserve arrayOutput(countLine) ' 配列0で返す場合があるため
  
  Do Until EOF(fp)
    'ライブラリリストを1行ずつ処理
    Dim strLine As String
    Line Input #fp, strLine

    Dim isLf As Long
    isLf = InStr(strLine, vbLf)
    
    If nameOS Like "Windows *" And Not isLf = 0 Then
      'OSがWindows かつ リストに LFが含まれる場合 (ファイルがUNIX形式)
      'ファイル全体で1行に見えてしまう。
      
      Dim arrayLineLF As Variant
      arrayLineLF = Split(strLine, vbLf)
    
      Dim i As Integer
      For i = 0 To UBound(arrayLineLF) - 1
        '行頭が '(コメント) ではない & 空行ではない場合
        If Not Left(arrayLineLF(i), 1) = "'" And Len(arrayLineLF(i)) > 0 Then
      
          '配列への追加
          countLine = countLine + 1
          ReDim Preserve arrayOutput(countLine)
          arrayOutput(countLine - 1) = arrayLineLF(i)
        End If
      Next i
    
    Else
      'OSがWindows and ファイルがWindows形式 (変換不要)
      'OSがMacOS X and ファイルがUNIX形式 (変換不要)
      
      'OSがMacOS X and ファイルがWindows形式
      ' vbCrがモジュールファイル名を発見できなくなる。
      strLine = Replace(strLine, vbCr, "")
        
      '行頭が '(コメント) ではない & 空行ではない場合
      If Not Left(strLine, 1) = "'" And Len(strLine) > 0 Then
      
        '配列への追加
        countLine = countLine + 1
        ReDim Preserve arrayOutput(countLine)
        arrayOutput(countLine - 1) = strLine
      End If
    
    End If
  Loop

  '3. リストファイルを閉じる
  Close #fp
  
  '4. 戻り値を配列で返す
  list2array = arrayOutput
End Function


' ショートカットの設定 (Macでは Macro指定できないっぽい)
Private Sub setShortKey()
  If Application.OperatingSystem Like "Windows *" Then
    ' Application.MacroOptions Macro:="ThisWorkbook.reloadModule", ShortcutKey:=SHORTKEY_RELOAD
    Application.MacroOptions Macro:="ThisWorkbook.reloadModule", ShortcutKey:="M"
  
  Else
    ' Mac OS Xの場合の注意: ThisWorkbook.reloadModule関数を持つマクロファイルを複数開いていると、
    ' 最後に開いたマクロファイルの ThisWorkbook.reloadModule関数が呼び出される模様。
    ' (その場合、マクロ一覧から'該当マクロファイル!reloadModule' を呼び出してください。)
    Application.OnKey SHORTKEY_RELOAD, "ThisWorkbook.reloadModule"

  End If
  
End Sub


'ショートカット設定の削除 (Macでは Macro指定できないっぽい)
Private Sub clearShortKey()
  If Application.OperatingSystem Like "Windows *" Then
    Application.MacroOptions Macro:="ThisWorkbook.reloadModule", ShortcutKey:=""
  
  Else
    ' Mac OS Xの場合の注意: ThisWorkbook.reloadModule関数を持つマクロファイルを複数開いていると、
    ' 最後に開いたマクロファイルの ThisWorkbook.reloadModule関数がクリアされる可能性が高いと思われる(未検証)。
    Application.OnKey SHORTKEY_RELOAD, ""
  End If
  
End Sub


' ==== https://blog.goo.ne.jp/end-u/e/b1e41190650cbb1ff2c70bafe92a6dce ====
' = 3/3 =
'-------------------------------------------------
' 新規シートの追加は、flg を使って識別。
'-------------------------------------------------
Private Sub Workbook_NewSheet(ByVal Sh As Object)
    flg = True
End Sub

'-------------------------------------------------
'Private Sub Workbook_Open()
'    shCount = Sheets.count
'End Sub
'-------------------------------------------------

Private Sub Workbook_SheetActivate(ByVal Sh As Object)

   ' bookの booted 前はワークシートの活性に反応しない。
   If Not booted Then Exit sub

   Dim m As Long

   ' MsgBox shCount
   ' shCount を Public にしたのだが、ここで値を参照できていない。
   ' しかたないので、初回（shCountの値が０）のときには実行しない
   ' しかけにした。
   m = Sheets.count - shCount
   If (m > 0) And (shCount > 0) Then
      MsgBox m
      shName = Sh.Name
      ' ====
      ' ワークシートが追加されたとき実行される関数の名前
      ' ====
      Application.OnTime Now, Me.CodeName & ".test"
      If Sheets.count > shCount Then
         shCount = shCount + 1
         ' ActiveSheet.Next.Activate
         ' ここで呼んではだめ。
      End If
   ElseIf m < 0 Then
      ' shCount = shCount + m
      shCount = Sheets.count
      MsgBox shCount
   End If
End Sub
'-------------------------------------------------
' 新規シートの追加は、flg を使って識別し、なにもしない。
'----
Private Sub test()
    If flg Then
        flg = False
    Else
        MsgBox shName
        ' ====
        ' シートを追加したときに実行する関数はここから呼ぶ。
        ' ====
        ' Call RuledLineCode
        If Sheets.count > shCount Then
            ActiveSheet.Next.Activate
            ' shCountが実際のシート数と一致するまで次のシート
            ' のアクティベートをする。
            ' そうすると、イベントが発生して、ハンドラの
            ' Workbook_SheetActivate が実行されるので、その中
            ' で定義されているこの関数が結果的に再帰呼び出し
        End If
    End If
End Sub
' =========================================================================


