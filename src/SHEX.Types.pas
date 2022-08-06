unit SHEX.Types;


interface

uses
  Windows, ShellApi,
  JPL.Strings,
  SHEX.Procs;

const
  NO_HANDLE = 0;


  SW_HIDE = 0;
  SW_SHOWNORMAL = 1;
  SW_NORMAL = SW_SHOWNORMAL;
  SW_SHOWMINIMIZED = 2;
  SW_SHOWMAXIMIZED = 3;
  SW_MAXIMIZE = SW_SHOWMAXIMIZED;
  SW_SHOWNOACTIVATE = 4;
  SW_SHOW = 5;
  SW_MINIMIZE = 6;
  SW_SHOWMINNOACTIVE = 7;
  SW_SHOWNA = 8;
  SW_RESTORE = 9;
  SW_SHOWDEFAULT = 10;
  SW_FORCEMINIMIZE = 11;

  URL_INFO_SHELLEXECUTE = 'https://docs.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shellexecutew';
  URL_INFO_SHELLEXECUTEINFO = 'https://docs.microsoft.com/en-us/windows/win32/api/shellapi/ns-shellapi-shellexecuteinfow';
  URL_INFO_SHOWCMD = 'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow#parameters';
  URL_INFO_WAIT_FINISH = 'https://docs.microsoft.com/en-us/windows/win32/api/synchapi/nf-synchapi-waitforsingleobject';
  URL_INFO_WAIT_IDLE = 'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-waitforinputidle';

type

  TWaitFor = (wfDisabled, wfFinish, wfIdle);

  TAppParams = record
    WindowHandle: HWND; // -hwnd --window-handle
    Operation: string;  // -o --operation
    &File: string;      // -f --file
    Parameters: string; // -p -- params
    Directory: string;  // -d --directory
    ShowCmd: integer;   // -s --show-cmd
    WaitFor: TWaitFor;  // -w --wait
    WaitTimeMs: Int64;  // -t --wait-time 1000 // wait 1 sec // Default: INFINITE
    procedure Init;
  end;

  TShowCmd = record
    ShowCmd: integer;
    StrIDs: array of string;
    Description: string;
    procedure Init(const AShowCmd: integer);
    function TryInitFromStr(const ShowCmdStr: string; IgnoreCase: Boolean = True): Boolean;
    function IDsAsStr(const Separator: string = ', '): string;
  end;

  // https://docs.microsoft.com/en-us/windows/win32/api/shellapi/ns-shellapi-shellexecuteinfow - hInstApp
  // https://docs.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shellexecutea - return value
  TShellExecuteError = record
    function IsShellExecuteError(const ErrNo: integer): Boolean;
    function Description(const ErrNo: integer): string;
  end;

  
implementation


{ TAppParams }

procedure TAppParams.Init;
begin
  WindowHandle := NO_HANDLE;
  Operation := 'open';
  &File := '';
  Parameters := '';
  Directory := '';
  ShowCmd := SW_SHOWNORMAL;
  WaitFor := wfDisabled;
  Self.WaitTimeMs := INFINITE;
end;


{ TShowCmd }

procedure TShowCmd.Init(const AShowCmd: integer);

  procedure AddData(ArrIDs: array of string; Desc: string = '');
  var
    i: integer;
  begin
    SetLength(StrIDs, Length(ArrIDs));
    for i := 0 to High(ArrIDs) do StrIDs[i] := ArrIDs[i];
    Description := Desc;
  end;

begin
  ShowCmd := AShowCmd;

  case ShowCmd of

    SW_HIDE:
      AddData(
        ['0', 'SW_HIDE', 'Hide'],
        'Hides the window and activates another window.'
      );

    SW_SHOWNORMAL:
      AddData(
        ['1', 'SW_SHOWNORMAL', 'ShowNormal', 'SW_NORMAL', 'Normal'],
        'Activates and displays a window. If the window is minimized or maximized, the system restores it ' +
        'to its original size and position. An application should specify this flag when displaying the window for the first time.'
      );

    SW_SHOWMINIMIZED:
      AddData(
        ['2', 'SW_SHOWMINIMIZED', 'ShowMinimized', 'Minimized'],
        'Activates the window and displays it as a minimized window.'
      );

    SW_SHOWMAXIMIZED:
      AddData(
        ['3', 'SW_SHOWMAXIMIZED', 'ShowMaximized', 'Maximized', 'Maximize'],
        'Activates the window and displays it as a maximized window.'
      );

    SW_SHOWNOACTIVATE:
      AddData(
        ['4', 'SW_SHOWNOACTIVATE', 'ShowNoActivate', 'NoActivate'],
        'Displays a window in its most recent size and position. This value is similar to SW_SHOWNORMAL, ' +
        'except that the window is not activated.'
      );

    SW_SHOW:
      AddData(
        ['5', 'SW_SHOW', 'Show'],
        'Activates the window and displays it in its current size and position.'
      );

    SW_MINIMIZE:
      AddData(
        ['6', 'SW_MINIMIZE', 'Minimize'],
        'Minimizes the specified window and activates the next top-level window in the Z order.'
      );

    SW_SHOWMINNOACTIVE:
      AddData(
        ['7', 'SW_SHOWMINNOACTIVE', 'ShowMinNoActive', 'MinNoActive'],
        'Displays the window as a minimized window. This value is similar to SW_SHOWMINIMIZED, except the window is not activated.'
      );

    SW_SHOWNA:
      AddData(
        ['8', 'SW_SHOWNA', 'ShowNA'],
        'Displays the window in its current size and position. This value is similar to SW_SHOW, except that the window is not activated.'
      );

    SW_RESTORE:
      AddData(
        ['9', 'SW_RESTORE', 'Restore'],
        'Activates and displays the window. If the window is minimized or maximized, the system restores it to its original ' +
        'size and position. An application should specify this flag when restoring a minimized window.'
      );

    SW_SHOWDEFAULT:
      AddData(
        ['10', 'SW_SHOWDEFAULT', 'ShowDefault'],
        'Sets the show state based on the SW_ value specified in the STARTUPINFO structure passed to the CreateProcess ' +
        'function by the program that started the application.'
      );

    SW_FORCEMINIMIZE:
      AddData(
        ['11', 'SW_FORCEMINIMIZE', 'ForceMinimize'],
        'Minimizes a window, even if the thread that owns the window is not responding. This flag should only be used when ' +
        'minimizing windows from a different thread.'
      );

  end;
end;

function TShowCmd.TryInitFromStr(const ShowCmdStr: string; IgnoreCase: Boolean = True): Boolean;
var
  xShowCmd: integer;
  scTemp: TShowCmd;
begin
  Result := False;
  for xShowCmd := 0 {SW_HIDE} to 11 {SW_FORCEMINIMIZE} do
  begin
    scTemp.Init(xShowCmd);
    if StrInArray(ShowCmdStr, scTemp.StrIDs, IgnoreCase) then
    begin
      Init(xShowCmd);
      Result := True;
      Break;
    end;
  end;
end;

function TShowCmd.IDsAsStr(const Separator: string = ', '): string;
var
  i: integer;
begin
  Result := '';
  for i := 0 to High(StrIDs) do
  begin
    Result := Result + StrIDs[i];
    if i < High(StrIDs) then Result := Result + Separator;
  end;
end;


{ TShellExecuteError }

function TShellExecuteError.Description(const ErrNo: integer): string;
begin
  case ErrNo of
    SE_ERR_FNF: Result := 'File not found';
    SE_ERR_PNF: Result := 'Path not found';
    SE_ERR_ACCESSDENIED: Result := 'Access denied';
    SE_ERR_OOM: Result := 'Out of memory';
    SE_ERR_DLLNOTFOUND: Result := 'Dynamic-link library not found';
    SE_ERR_SHARE: Result := 'Cannot share an open file';
    SE_ERR_ASSOCINCOMPLETE: Result := 'File association information not complete';
    SE_ERR_DDETIMEOUT: Result := 'DDE operation timed out';
    SE_ERR_DDEFAIL: Result := 'DDE operation failed';
    SE_ERR_DDEBUSY: Result := 'DDE operation is busy';
    SE_ERR_NOASSOC: Result := 'File association not available';
  else
    Result := '';
  end;
end;

function TShellExecuteError.IsShellExecuteError(const ErrNo: integer): Boolean;
begin
  case ErrNo of
    SE_ERR_FNF, SE_ERR_PNF, SE_ERR_ACCESSDENIED, SE_ERR_OOM, SE_ERR_DLLNOTFOUND,
    SE_ERR_SHARE, SE_ERR_ASSOCINCOMPLETE, SE_ERR_DDETIMEOUT, SE_ERR_DDEFAIL, SE_ERR_DDEBUSY, SE_ERR_NOASSOC
    : Result := True;
  else
    Result := False;
  end;
end;

end.
