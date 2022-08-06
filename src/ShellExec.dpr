program ShellExec;

{
  Jacek Pazera
  https://www.pazera-software.com
  https://github.com/jackdp

  ------------------------------------------------------
  ShellExec - Performs an operation on a specified file.
  A "console wrapper" for the ShellExecuteEx function.
  ------------------------------------------------------

  Links:
  https://docs.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shellexecuteexw
  https://docs.microsoft.com/en-us/windows/win32/api/shellapi/ns-shellapi-shellexecuteinfow
  https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow#parameters
  https://docs.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shellexecutew
  https://docs.microsoft.com/en-us/windows/win32/api/synchapi/nf-synchapi-waitforsingleobject
  https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-waitforinputidle
}


{$APPTYPE CONSOLE}

// Disable extended RTTI
{$IF CompilerVersion >= 21.0} // >= Delphi 2010
  {$WEAKLINKRTTI ON}
  {$RTTI EXPLICIT METHODS([]) PROPERTIES([]) FIELDS([])}
{$ENDIF}

{$SetPEFlags 1}   // IMAGE_FILE_RELOCS_STRIPPED
{$SetPEFlags $20} // IMAGE_FILE_LARGE_ADDRESS_AWARE



uses
  Windows,
  SysUtils,
  JPL.Console,
  SHEX.App in 'SHEX.App.pas',
  SHEX.Types in 'SHEX.Types.pas',
  SHEX.Procs in 'SHEX.Procs.pas';

var
  App: TApp;


procedure MyExitProcedure;
begin
  if Assigned(App) then
  begin
    App.Done;
    FreeAndNil(App);
  end;
end;


{$R *.res}


// ---------------------------- ENTRY POINT ------------------------------
begin

  {$IFDEF DEBUG}
  ReportMemoryLeaksOnShutdown := True;
  {$ENDIF}


  App := TApp.Create;
  try

    try

      App.ExitProcedure := @MyExitProcedure;
      App.Init;
      App.Run;
      if Assigned(App) then App.Done;

    except

      on E: Exception do
      begin
        Writeln(E.ClassName, ': ', E.Message);
        if GetLastError <> 0 then Writeln('OS Error No. ', GetLastError, ': ', SysErrorMessage(GetLastError));
        ExitCode := TConsole.ExitCodeError;
      end;

    end;

  finally
    if Assigned(App) then FreeAndNil(App);
  end;

end.

