unit jh_Basic;

interface

uses Classes, SysUtils, Windows;

function CreateShortcutDesktop(aFullName, aLinkName: String): Boolean;

function DecimalFormat(aDecimal: Integer): String;
function GetProgramFilesPath32: TFileName;
function GetProgramFilesPath64: TFileName;
function GetCurrentPath: String;

function ExecAndWait(sExecutableFile: String;
  wWindowState: Word = SW_SHOWNORMAL): Boolean;
function ExecAndNowait(sExecutableFile: String;
  wWindowState: Word = SW_SHOWNORMAL): Boolean;

function IIF(v1, v2, v3: Boolean): Boolean; OverLoad;
function IIF(v1: Boolean; v2, v3: String): String; OverLoad;
function IIF(v1: Boolean; v2, v3: Integer): Integer; OverLoad;
function IIF(v1: Boolean; v2, v3: Smallint): Smallint; OverLoad;

function StreamToVariant(Stream: TStream): OleVariant;
function VariantToStream(aVar: OleVariant): TStream;

function RoundF(X: Extended; Decimal: integer = 0): Extended;

implementation

uses Forms, StrUtils, Registry, Variants, Math,
  ShlObj, ActiveX, ComObj;

function CreateShortcutDesktop(aFullName, aLinkName: String): Boolean;
var
  IObject : IUnknown;
  ISLink : IShellLink;
  IPFile : IPersistFile;
  PIDL : PItemIDList;
  InFolder : array[0..MAX_PATH] of Char;
  vLinkName: String;
begin
  {Use TargetName:=ParamStr(0) which returns the path and file name of the
  executing program to create a link to your Application}

  IObject := CreateComObject(CLSID_ShellLink) ;
  ISLink := IObject as IShellLink;
  IPFile := IObject as IPersistFile;

  with ISLink do
  begin
    SetPath(pChar(aFullName)) ;
    SetWorkingDirectory(pChar(ExtractFilePath(aFullName))) ;
    SetDescription(pChar(aLinkName));
  end;

  // if we want to place a link on the Desktop
  SHGetSpecialFolderLocation(0, CSIDL_DESKTOPDIRECTORY, PIDL) ;
  SHGetPathFromIDList(PIDL, InFolder) ;

  {
   or if we want a link to appear in
   some other, not-so-special, folder:
   InFolder := 'c:\SomeFolder'
  }

  vLinkName := String(InFolder) + '\' + aLinkName + '.lnk';
  try
    IPFile.Save(PWChar(vLinkName), false) ;
    Result := True;
  except
    Result := False;
  end;
end;

function DecimalFormat(aDecimal: Integer): String;
var
  vStr: String;
begin
  vStr := '#,##0';

  if aDecimal <> 0 then
    vStr := vStr + '.' + DupeString('0', aDecimal);

  Result := vStr;
end;

function GetProgramFilesPath32: TFileName;
var
  reg: TRegistry;
  vName: String;
begin
  if GetProgramFilesPath64 = '' then
    vName := 'ProgramFilesDir'
  else
    vName := 'ProgramFilesDir (x86)';

  reg := TRegistry.Create(KEY_READ);
  try
    reg.RootKey := HKEY_LOCAL_MACHINE;
    reg.OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion', False);
    Result := reg.ReadString(vName);
  finally
    reg.Free;
  end;
end;

function GetProgramFilesPath64: TFileName;
var
  reg: TRegistry;
begin
  reg := TRegistry.Create(KEY_READ);
  try
    reg.RootKey := HKEY_LOCAL_MACHINE;
    reg.OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion', False);
    Result := reg.ReadString('ProgramW6432Dir');
  finally
    reg.Free;
  end;
end;

function GetCurrentPath: String;
begin
  Result := ExtractFilePath(Forms.Application.ExeName);
end;

function ExecAndWait(sExecutableFile: String;
  wWindowState: Word = SW_SHOWNORMAL): Boolean;
var
  siInfo: TStartUpInfo;
  piInfo: TProcessInformation;
begin
  FillChar(siInfo, SizeOf(siInfo), #0);

  with siInfo do
  begin
    cb := SizeOf(siInfo);
    dwFlags := STARTF_USESHOWWINDOW;
    wShowWindow := wWindowState;
  end;

  Result := CreateProcess(NIL, pChar(sExecutableFile), NIL, NIL, False,
    CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS, NIL,
    pChar(ExtractFilePath(sExecutableFile)), siInfo, piInfo);

  if Result then
    WaitForSingleObject(piInfo.hprocess, INFINITE);

  // Example:
  // Run Windows calculator.
  // ExecuteAndWait('C:\Windows\system32\Calc.exe');
end;

function ExecAndNowait(sExecutableFile: String;
  wWindowState: Word = SW_SHOWNORMAL): Boolean;
var
  siInfo: TStartUpInfo;
  piInfo: TProcessInformation;
begin
  FillChar(siInfo, SizeOf(siInfo), #0);

  with siInfo do
  begin
    cb := SizeOf(siInfo);
    dwFlags := STARTF_USESHOWWINDOW;
    wShowWindow := wWindowState;
  end;

  Result := CreateProcess(NIL, pChar(sExecutableFile), NIL, NIL, False,
    CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS, NIL,
    pChar(ExtractFilePath(sExecutableFile)), siInfo, piInfo);

  // if Result then
  // WaitForSingleObject(piInfo.hprocess, INFINITE);

  // Example:
  // Run Windows calculator.
  // ExecuteAndWait('C:\Windows\system32\Calc.exe');
end;

function IIF(v1, v2, v3: Boolean): Boolean; OverLoad;
begin
  if v1 then
    Result := v2
  else
    Result := v3;
end;

function IIF(v1: Boolean; v2, v3: String): String; OverLoad;
begin
  if v1 then
    Result := v2
  else
    Result := v3;
end;

function IIF(v1: Boolean; v2, v3: Integer): Integer; OverLoad;
begin
  if v1 then
    Result := v2
  else
    Result := v3;
end;

function IIF(v1: Boolean; v2, v3: Smallint): Smallint; OverLoad;
begin
  if v1 then
    Result := v2
  else
    Result := v3;
end;

function StreamToVariant(Stream: TStream): OleVariant;
var
  p: Pointer;
begin
  Result := VarArrayCreate([0, Stream.Size - 1], varByte);
  p := VarArrayLock(Result);
  try
    Stream.Position := 0;
    Stream.Read(p^, Stream.Size);
  finally
    VarArrayUnlock(Result);
  end;
end;

function VariantToStream(aVar: OleVariant): TStream;
var
  Data: PByteArray;
  Size: Integer;
begin
  Result := TMemoryStream.Create;
  try
    Size := VarArrayHighBound(aVar, 1) - VarArrayLowBound(aVar, 1) + 1;
    if Size = 0 then
    begin
      if Result <> nil then
        Result.Free;
      Result := nil;
      exit;
    end;
    Data := VarArrayLock(aVar);
    try
      Result.Position := 0;
      Result.WriteBuffer(Data^, Size);
    finally
      VarArrayUnlock(aVar);
    end;
  except
    VarArrayUnlock(aVar);
    Result.Free;
    Result := nil;
  end;
end;

function RoundF(X: Extended; Decimal: integer = 0): Extended;
var
  PowerNum: Extended;

  function RoundI(X: Extended): Int64;
  begin
    if X < 0 then
      Result := Round(X - 0.0000001)
    else
      Result := Round(X + 0.0000001);
  end;
begin
  PowerNum := IntPower(10, Decimal);
  Result := RoundI(X * PowerNum) / PowerNum;
end;

end.
