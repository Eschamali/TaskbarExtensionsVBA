# TaskbarProgress_VBA

VBAから、タスクバー拡張機能の以下の機能にアクセスするDLLファイルを提供します。マクロで、タスクバーに変化を持たせたいときにどうぞ。
- [ITaskbarList3::SetProgressState](https://learn.microsoft.com/ja-jp/windows/desktop/api/shobjidl_core/nf-shobjidl_core-itaskbarlist3-setprogressstate)
- [ITaskbarList3::SetProgressValue](https://learn.microsoft.com/ja-jp/windows/desktop/api/shobjidl_core/nf-shobjidl_core-itaskbarlist3-setprogressvalue)
- [ITaskbarList3::SetOverlayIcon](https://learn.microsoft.com/ja-jp/windows/desktop/api/shobjidl_core/nf-shobjidl_core-itaskbarlist3-setoverlayicon)

# DEMO
- SetProgressState
- SetProgressValue

| シチュエーション例 | 動作イメージ | 
| ---------------- | ------------ | 
| データ準備→処理中→終了 | ![alt text](doc/Demo1.gif)       | 
| 処理中→一時中断→再開 | ![alt text](doc/Demo2.gif)       |
| 処理中→エラー→終了 | ![alt text](doc/Demo3.gif)       |

タスクバー ボタンに表示される進行状況インジケーターの種類と状態を設定できます。

- SetOverlayIcon

| シチュエーション例 | 動作イメージ | 
| ---------------- | ------------ | 
| 検索中… | ![alt text](doc/Demo4.png)       | 
| ファイル移動中… | ![alt text](doc/Demo5.png)       | 
| 処理中→エラー→終了 | ![alt text](doc/Demo6.gif)       | 

こんな感じで、ステータスの表現が可能です。<br>
プログレスバーと合わせて表現すると良いと思います。

# Features

- DLLインポートにより、数行で手軽に進捗状況とステータスの表現が可能です。ユーザーフォーム作ってプログレスバーを埋め込んで、呼び出して…　という手間が省けます。
- ステータスに使えるアイコンソースファイルは、下記に対応しています
  - .icoファイル: 単独のアイコンファイル。
  - .exeファイル: 実行ファイル内に埋め込まれたリソースアイコン。
  - .dllファイル: DLL内に埋め込まれたリソースアイコン。

# Requirement

以下で検証済みです。

- Microsoft 365 Excel 64bit
- Windows 10 , 11 64bit

タスクバーのプログレスバー自体は、Windows 7から実装されたものですが当本人、所有していないためWin 10未満のOSは、未検証です…<br>
Office製品も同様です。

# Load DLL

WindowsAPIの「LoadLibrary関数」を使って、読み込みます。

```bas
hDll = LoadLibrary("TaskbarProgress.dll")
```

実際に使う場合は、"Excelファイル(.xlsm)の存在するディレクトリ"というような[動的な場所を設定する仕組み](https://liclog.net/vba-dll-create-5/)で読み込むことをおすすめします。

```bas
'動的にDLLを取得するためのWinAPI
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr

Private Sub Workbook_Open()

    Dim hDll As LongPtr
    Dim sFolderPath As String
    
    'DLLファイルを保存するフォルダパスを設定
    sFolderPath = ThisWorkbook.Path
    
    'DLLﾌｧｲﾙを読み込む
    hDll = LoadLibrary(sFolderPath & "\" & "TaskbarProgress.dll")　'DLLファイルフルパス

    debug.print hDll
End Sub
```

hDll の中身が、0 以外であれば読み込み、成功です。


# Usage

使用する際はまず、このように定義します。Office 2010以降なら、32bit,64bit 共通で下記で読み込み可能です。

```bas
Declare PtrSafe Sub SetTaskbarProgress Lib "TaskbarProgress.dll" (ByVal hwnd As LongPtr, ByVal current As Long, ByVal maximum As Long, ByVal Status As Long)
Declare PtrSafe Sub SetTaskbarOverlayIcon Lib "TaskbarProgress.dll" (ByVal hwnd As LongPtr, ByVal dllPath As LongPtr, ByVal iconIndex As Long, ByVal description As LongPtr)
```

## SetTaskbarProgress
サンプルコード
```bas
Sub TaskbarProgressTest()
    Dim hwnd As LongPtr
    Dim current As Long
    Dim maximum As Long
    Dim Status As Long
    
    ' ウィンドウハンドルを取得
    hwnd = Application.hwnd
    
    ' 進捗の設定
    current = 50
    maximum = 100
    Status = 2
    
    ' DLL関数の呼び出し
    SetTaskbarProgress hwnd, current, maximum, Status
End Sub
```

上記のサンプルを実行すると、このようになります。<br>
![alt text](doc/Demo7.png)

### 引数の説明

| 名称            | 説明                                                                             | 
| --------------- | -------------------------------------------------------------------------------- | 
| hwnd            | 適用させるウィンドウハンドルを指定。<br>基本は、Application.hwnd　を指定します。 | 
| current         | 現在の進捗値                                                                     | 
| maximum         | 最大(ゲージMax)とする値                                                          | 
| Status          | プログレスバーの状態(0,1,2,4,8 のいずれか)                                        | 

### Status について

説明は[こちらから引用](https://learn.microsoft.com/ja-jp/windows/win32/api/shobjidl_core/nf-shobjidl_core-itaskbarlist3-setprogressstate)しています

| 値  | 説明                                                                                                                                              | イメージ | 
| --- | ------------------------------------------------------------------------------------------------------------------------------------------------- | -------- | 
| 0   | TBPF_NOPROGRESS<br>進行状況の表示を停止し、ボタンを通常の状態に戻します。                                                                         |![alt text](doc/Demo8.png)| 
| 1   | TBPF_INDETERMINATE<br>進行状況インジケーターのサイズは拡大しませんが、タスク バー ボタンの長さに沿って繰り返し循環します。                        |![alt text](doc/Demo9.gif)| 
| 2   | TBPF_NORMAL<br>進行状況インジケーターのサイズは、完了した操作の推定量に比例して左から右に大きくなります。                                         |![alt text](doc/Demo7.png)| 
| 4   | TBPF_ERROR<br>進行状況インジケーターが赤に変わり、進行状況をブロードキャストしているいずれかのウィンドウでエラーが発生したことを示します。        |![alt text](doc/Demo10.png)| 
| 8   | TBPF_PAUSED<br>進行状況インジケーターが黄色に変わり、進行状況は現在いずれかのウィンドウで停止されていますが、ユーザーが再開できることを示します。 |![alt text](doc/Demo11.png)| 

## SetTaskbarOverlayIcon
サンプルコード
```bas
Sub SetOverlayIconFromDLLExample()
    Dim hwnd As LongPtr
    Dim dllPath As String
    Dim iconIndex As Long
    Dim description As String
    
    ' ウィンドウハンドルを取得
    hwnd = Application.hwnd
    
    ' 任意のアイコンデータがあるフルパス(ico,dll,exe に対応)
    'dllPath = "C:\Program Files\Microsoft Office\root\Office16\XLICONS.EXE"
    dllPath = "C:\Windows\System32\shell32.dll"
    'dllPath = "C:\Users\user\Downloads\sample.ico"

    ' アイコンのインデックス（DLL,exe内のアイコン番号）
    iconIndex = 240
    
    ' アイコンの説明テキスト
    description = "Custom Icon from DLL"
    
    ' DLL関数を呼び出し、タスクバーにオーバーレイアイコンを設定
    SetTaskbarOverlayIcon hwnd, StrPtr(dllPath), iconIndex, StrPtr(description)
End Sub
```

上記のサンプルをWin 11で実行すると、このようになります。<br>
![alt text](doc/Demo12.png)

### 引数の説明

| 名称            | 説明                                                                             | 
| --------------- | -------------------------------------------------------------------------------- | 
| hwnd            | 適用させるウィンドウハンドルを指定。<br>基本は、Application.hwnd　を指定します。 | 
| dllPath         | 任意のアイコンデータがあるフルパス<br>文字列へのポインターを利用するため、StrPtrを経由する必要があります | 
| iconIndex       | アイコンのインデックス（DLL,exe内のアイコン番号）<br>icoファイルの場合は、この設定を無視します。| 
| description     | アクセシビリティ用の代替テキスト<br>文字列へのポインターを利用するため、StrPtrを経由する必要があります | 

なお、dllPath を省略したり、空文字、False で渡されると、ステータスアイコンを除去します。

# Note
DLL側の処理は、ある程度のエラー処理を施していますが、決して完璧ではありません。<br>
そのため、DLLの関数を直接呼ぶのではなく、VBAの標準モジュール内のプロシージャを介して、呼び出すことを推奨します。最近のPCであれば、誤差レベルです。<br>
また、Win32APIの文字列関係はほとんど文字列へのポインターによって取得されるため、毎度 StrPtr を記述するのは大変なので、プロシージャを介した作りにするといいでしょう。