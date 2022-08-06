unit SHEX.App;

interface

uses
  Windows, ShellApi,
  SysUtils,
  JPL.Console, JPL.ConsoleApp, JPL.CmdLineParser,
  JPL.Strings, JPL.TStr, JPL.Conversion,
  JPL.DefinitionList,
  //JPL.Win.VersionInfo,
  SHEX.Types, SHEX.Procs;

type


  TApp = class(TJPConsoleApp)
  private
    AppParams: TAppParams;
    DefList_Operations: TJPDefinitionList;
    DefList_ShowCmd: TJPDefinitionList;
    DefList_WaitFor: TJPDefinitionList;
    clOperation: string;
    clURL: string;
    clShowCmd: string;
    clWaitFor: string;
  public
    procedure Init;
    procedure Run;
    procedure Done;

    procedure RegisterOptions;
    procedure ProcessOptions;

    procedure PerformMainAction;

    procedure DisplayHelpAndExit(const ExCode: integer);
    procedure DisplayShortUsageAndExit(const Msg: string; const ExCode: integer);
    procedure DisplayBannerAndExit(const ExCode: integer);
    procedure DisplayMessageAndExit(const Msg: string; const ExCode: integer);
  end;



implementation



{$region '                    Init                              '}

procedure TApp.Init;
const
  SEP_LINE = '------------------------------------------------------------';
var
  ShowCmdRec: TShowCmd;
  i, x: integer;
  st: string;
begin
  //----------------------------------------------------------------------------

  AppName := 'ShellExec';
  MajorVersion := 1;
  MinorVersion := 0;
  Self.Date := EncodeDate(2022, 8, 6);
  FullNameFormat := '%AppName% %MajorVersion%.%MinorVersion% [%OSShort% %Bits%-bit] (%AppDate%)';
  Description := 'Performs an operation on a specified file or directory or URL.';
  LicenseName := 'Freeware, Open Source';
  License := 'This program is completely free. You can use it without any restrictions, also for commercial purposes.' + ENDL +
    'The program''s source files are available at https://github.com/jackdp/ShellExec';
  Author := 'Jacek Pazera';
  HomePage := 'https://www.pazera-software.com/products/shellexec/';
  HelpPage := HomePage;
  SourcePage := 'https://github.com/jackdp/ShellExec';
  TrimExtFromExeShortName := True;

  HintBackgroundColor := TConsole.clLightGrayBg;
  HintTextColor := TConsole.clBlackText;

  clOperation := 'cyan';
  clURL := 'lightblue';
  clShowCmd := 'darkyellow';
  clWaitFor := 'lime';

  AppParams.Init;

  // ---------------------- Operations ---------------------
  DefList_Operations := TJPDefinitionList.Create;
  DefList_Operations.TermPadding := 0;
  DefList_Operations.DescriptionPadding := 2;
  DefList_Operations.MaxTermLineLength := 60;
  DefList_Operations.MaxDescriptionLineLength := 90;

  DefList_Operations.Add(
    ColorTag('open', clOperation),
    'Opens the item specified by the FILE parameter. The item can be a file or folder or URL.'
  );

  DefList_Operations.Add(
    ColorTag('edit', clOperation),
    'Launches an editor and opens the document for editing. If FILE is not a document file, the ShellExec will fail.'
  );

  DefList_Operations.Add(
    ColorTag('explore', clOperation),
    'Explores a folder specified by FILE.'
  );

  DefList_Operations.Add(
    ColorTag('find', clOperation),
    'Initiates a search beginning in the directory specified by -d option.'
  );

  DefList_Operations.Add(
    ColorTag('print', clOperation),
    'Prints the file specified by FILE. If FILE is not a document file, the program fails.'
  );

  DefList_Operations.Add(
    ColorTag('runas', clOperation),
    'Launches an application as Administrator. User Account Control (UAC) will prompt the user for consent to ' +
    'run the application elevated or enter the credentials of an administrator account used to run the application.'
  );

  DefList_Operations.Add(
    ColorTag('properties', clOperation),
    'Displays the file or folder''s properties.'
  );

  // ------------------------- ShowCmd -------------------------
  DefList_ShowCmd := TJPDefinitionList.Create;
  DefList_ShowCmd.TermPadding := 0;
  DefList_ShowCmd.DescriptionPadding := 2;
  DefList_ShowCmd.MaxTermLineLength := 200; // dużo, bo tag
  DefList_ShowCmd.MaxDescriptionLineLength := 90;

  for i := SW_HIDE to SW_FORCEMINIMIZE do
  begin
    ShowCmdRec.Init(i);

    st := '';
    for x := 0 to High(ShowCmdRec.StrIDs) do
      st := st + ColorTag(ShowCmdRec.StrIDs[x], clShowCmd) + ' | ';
    st := TrimFromEnd(st, ' | ');

    DefList_ShowCmd.Add(st, ShowCmdRec.Description);
  end;


  // ------------------------- Wait for -------------------------
  DefList_WaitFor := TJPDefinitionList.Create;
  DefList_WaitFor.TermPadding := 0;
  DefList_WaitFor.DescriptionPadding := 2;
  DefList_WaitFor.MaxTermLineLength := 60;
  DefList_WaitFor.MaxDescriptionLineLength := 90;

  DefList_WaitFor.Add(
    ColorTag('finish', clWaitFor),
    'Wait until the running program (specified in the FILE parameter) finishes its operation, or until the time-out interval (-t option) has elapsed.'
  );

  DefList_WaitFor.Add(
    ColorTag('idle', clWaitFor),
    'Waits until the specified program has finished processing its initial input and is waiting for user input with no input pending, '+
    'or until the time-out interval (-t option) has elapsed.'
  );


  //-----------------------------------------------------------------------------

  TryHelpStr := ENDL + 'Try <color=white,black>' + ExeShortName + ' --help</color> for more information.';


  ShortUsageStr :=
    ENDL +
    'Usage: ' + ExeShortName +
    ' <color=yellow>FILE</color> ' +
    '[-o ' + ColorTag('Operation', clOperation) + '] ' +
    '[-p Parameters] [-d Directory] ' +
    '[-s ' + ColorTag('ShowCommand', clShowCmd) + '] ' +
    '[-w [' + ColorTag('finish', clWaitFor) + '|' + ColorTag('idle', clWaitFor) + ']] ' +
    '[-t TIME] [-hwnd WindowHandle] ' +
    '[-h] [-V] [--license] [--home] [--source]' + ENDL2 +
    'Options are case-sensitive. Options in square brackets are optional.';


  ExtraInfoStr :=
    ENDL +
    SEP_LINE + ENDL +
    '<color=yellow>FILE</color>' + ENDL +
    'File name or URL or object on which to execute the specified operation.' + ENDL +

    SEP_LINE + ENDL +
    'OPERATIONS' + ENDL2 +
    DefList_Operations.AsString(1) + ENDL2 +
    'More information (lpOperation): ' + ColorTag(URL_INFO_SHELLEXECUTE, clURL) + ENDL +
    'More information (lpVerb): ' + ColorTag(URL_INFO_SHELLEXECUTEINFO, clURL) + ENDL +

    SEP_LINE + ENDL +
    'SHOW COMMANDS' + ENDL2 +
    DefList_ShowCmd.AsString(1) + ENDL2 +
    'More information: ' + ColorTag(URL_INFO_SHOWCMD, clURL) + ENDL +

    SEP_LINE + ENDL +
    'WAIT FOR' + ENDL2 +
    DefList_WaitFor.AsString(1) + ENDL2 +
    'More information: ' + ENDL +
    '  ' + ColorTag(URL_INFO_WAIT_FINISH, clURL) + ENDL +
    '  ' + ColorTag(URL_INFO_WAIT_IDLE, clURL) + ENDL +

    SEP_LINE + ENDL +
    'EXIT CODES' + ENDL2 +
    '0 - Success' + ENDL +
    'Any other value - Error' +

    '';



  ExamplesStr :=
    SEP_LINE + ENDL +
    'EXAMPLES' + ENDL2 +
    '  Starts the system calculator:' + ENDL +
    '    ' + ExeShortName + ' calc' + ENDL2 +
    '  Opens the "hosts" file in the system text editor, Notepad:' + ENDL +
    '    ' + ExeShortName + ' notepad -o runas -p "C:\Windows\System32\drivers\etc\hosts"' + ENDL2 +
    '  Displays the given website in the default browser:' + ENDL +
    '    ' + ExeShortName + ' https://example.com' + ENDL2 +
    '  Opens the "win.ini" file in the text editor associated with the INI files:' + ENDL +
    '    ' + ExeShortName + ' "C:\Windows\win.ini"' + ENDL2 +
    '  Opens the JPG file in the default graphic viewer:' + ENDL +
    '    ' + ExeShortName + ' "D:\pictures\my picture.jpg"' + ENDL2 +
    '  Opens the BAT file for editing in the default editor and waits for the program to finish:' + ENDL +
    '    ' + ExeShortName + ' "D:\batch_files\test.bat" -o edit -w=finish';

  //------------------------------------------------------------------------------

end;
{$endregion Init}


{$region '                    Run & Done                               '}
procedure TApp.Run;
begin
  inherited;

  RegisterOptions;
  Cmd.Parse;
  ProcessOptions;
  if Terminated then Exit;

  PerformMainAction; // <----- the main procedure
end;

procedure TApp.Done;
begin
  if Assigned(DefList_Operations) then DefList_Operations.Free;
  if Assigned(DefList_ShowCmd) then DefList_ShowCmd.Free;
  if Assigned(DefList_WaitFor) then DefList_WaitFor.Free;
end;
{$endregion Run & Done}


{$region '                    RegisterOptions                   '}
procedure TApp.RegisterOptions;
const
  MAX_LINE_LEN = 110;
var
  Category: string;
begin

  Cmd.CommandLineParsingMode := cpmCustom;
  Cmd.UsageFormat := cufWget;
  Cmd.AcceptAllNonOptions := True;


  // ------------ Registering command-line options -----------------

  Category := 'main';

  Cmd.RegisterOption(
    'o', 'operation', cvtRequired, False, False,
    'Action to be performed. The default action is ' + ColorTag('open', clOperation) + '. See description below.',
    'STR', Category
  );

  //Cmd.RegisterOption('f', 'file', cvtRequired, False, False, 'File or object on which to execute the specified operation.', 'FILE', Category);

  Cmd.RegisterOption(
    'p', 'params', cvtRequired, False, False,
    'If <color=yellow>FILE</color> specifies an executable file, this parameter is a string that specifies the parameters ' +
    'to be passed to the application.',
    'STR', Category
  );

  Cmd.RegisterOption('d', 'directory', cvtRequired, False, False, 'Working directory for the action.', 'DIR', Category);

  Cmd.RegisterOption(
    's', 'show-cmd', cvtRequired, False, False,
    'This parameter specifies how an application is to be displayed when it is opened. ' +
    'The default show command is ' + ColorTag('SW_SHOWNORMAL', clShowCmd) + '. ' +
    'See description below.',
    'STR', Category
  );

  Cmd.RegisterOption(
    'w', 'wait', cvtOptional, False, False,
    ColorTag('finish', clWaitFor) + ' OR ' + ColorTag('idle', clWaitFor) + '. Default: finish.        ' +
    'If you do not specify a waiting time in the "-t" option, INFINITE will be used  (-t=infinite). ' +
    'See description below.',
    'FOR', Category
  );

  Cmd.RegisterOption(
    't', 'wait-time', cvtRequired, False, False,
    'Time-out interval (for the "-w" option). Available time units: ms (milliseconds), s (seconds), m (minutes), h (hours), d (days). ' +
    'If no unit is specified, milliseconds will be used. The value INFINITE means no time limit.',
    'TIME', Category
  );

  Cmd.RegisterOption(
    'hwnd', 'window-handle', cvtRequired, False, False,
    'A handle to the parent window used for displaying an UI or error messages.', 'INT', Category
  );



  Category := 'info';

  Cmd.RegisterOption('h', 'help', cvtNone, False, False, 'Show this help.', '', Category);
  Cmd.RegisterShortOption('?', cvtNone, False, True, '', '', '');
  Cmd.RegisterOption('V', 'version', cvtNone, False, False, 'Show application version.', '', Category);
  Cmd.RegisterLongOption('license', cvtNone, False, False, 'Display program license.', '', Category);
  Cmd.RegisterLongOption('home', cvtNone, False, False, 'Opens program home page in the default browser.', '', Category);
  Cmd.RegisterLongOption('source', cvtNone, False, False, 'Opens the program page on GitHub with the program''s source files.', '', Category);


  UsageStr :=
    ENDL +
    'Options:' + ENDL + Cmd.OptionsUsageStr('  ', 'main', MAX_LINE_LEN, '  ', 30) + ENDL2 +
    'Info:' + ENDL + Cmd.OptionsUsageStr('  ', 'info', MAX_LINE_LEN, '  ', 30);


  Cmd.RegisterLongOption('dev-info', cvtNone, False, True, '', '', '');

end;
{$endregion RegisterOptions}


{$region '                    ProcessOptions                    '}
procedure TApp.ProcessOptions;
var
  x: integer;
  s: string;
  //evi: TExeVersionInfo;
  sc: TShowCmd;
begin

  // Hidden option: --dev-info
  if Cmd.IsLongOptionExists('dev-info') then
  begin
    Writeln(AppFullName);
    Writeln('Compiled with Delphi ' + CompilerVersionToDelphiName(CompilerVersion));
    Writeln('System.CompilerVersion: ', ftosEx(CompilerVersion, '0.0', '.'));
    Writeln('System.RTLVersion: ', ftosEx(RTLVersion, '0.0', '.'));
    //Writeln('Bitness: ' + APP_BITS_STR + '-bit');

    {$IFDEF DEBUG}s := 'Yes';{$ELSE}s := 'No';{$ENDIF}
    Writeln('DEBUG: ' + s);
//    if evi.ReadFromFile(ParamStr(0)) then
//    begin
//      Writeln('VERSION INFO');
//      //Writeln('  File version: ' + evi.FileVersion.AsString);
//      Writeln(evi.StringInfo.AsString(False, False, '  '));
//    end;

    Terminate;
    Exit;
  end;

  // ---------------------------- Invalid options -----------------------------------

  if Cmd.ErrorCount > 0 then
  begin
    DisplayShortUsageAndExit(Cmd.ErrorsStr, TConsole.ExitCodeSyntaxError);
    Exit;
  end;


  //------------------------------------ Help ---------------------------------------

  if (ParamCount = 0) or (Cmd.IsLongOptionExists('help')) or (Cmd.IsOptionExists('?')) then
  begin
    DisplayHelpAndExit(TConsole.ExitCodeOK);
    Exit;
  end;


  //---------------------------------- Home -----------------------------------------

  if Cmd.IsLongOptionExists('home') then
  begin
    GoToHomePage;
    Terminate;
    Exit;
  end;

  //---------------------------------- Source -----------------------------------------

  if Cmd.IsLongOptionExists('source') then
  begin
    GoToSourcePage;
    Terminate;
    Exit;
  end;


  //------------------------------- Version ------------------------------------------

  if Cmd.IsOptionExists('version') then
  begin
    DisplayBannerAndExit(TConsole.ExitCodeOK);
    Exit;
  end;


  //------------------------------- License ------------------------------------------

  if Cmd.IsLongOptionExists('license') then
  begin
    TConsole.WriteTaggedTextLine('<color=white,black>' + LicenseName + '</color>');
    DisplayLicense;
    Terminate;
    Exit;
  end;


  // -------------- Operation -------------

  if Cmd.IsOptionExists('o') then
    AppParams.Operation := Cmd.GetOptionValue('o');


  // ---------------- Parameters -------------

  if Cmd.IsOptionExists('p') then
  begin
    AppParams.Parameters := Cmd.GetOptionValue('p');
  end;


  // --------------- Directory -----------------

  if Cmd.IsOptionExists('d') then
  begin
    AppParams.Directory := Cmd.GetOptionValue('d');
  end;


  // ------------ Show command ----------------

  if Cmd.IsOptionExists('s') then
  begin
    s := Cmd.GetOptionValue('s');
    if not sc.TryInitFromStr(s, True) then
    begin
      DisplayError('Invalid value for the -s option!');
      ExitCode := TConsole.ExitCodeError;
      Terminate;
      Exit;
    end;
    AppParams.ShowCmd := sc.ShowCmd;
  end;


  // -------------- Wait for ---------------

  if Cmd.IsOptionExists('w') then
  begin
    s := Cmd.GetOptionValue('w');
    if s = '' then AppParams.WaitFor := wfFinish // if empty then use default wfFinish
    else
    begin
      s := LowerCase(Cmd.GetOptionValue('w'));
      if s = 'finish' then AppParams.WaitFor := wfFinish
      else if s = 'idle' then AppParams.WaitFor := wfIdle
      else
      begin
        DisplayError('Invalid value for the -w option: ' + s);
        TConsole.WriteTaggedTextLine('Expected: ' + ColorTag('finish', clWaitFor) + ' OR ' + ColorTag('idle', clWaitFor));
        ExitCode := TConsole.ExitCodeError;
        Terminate;
        Exit;
      end;
    end;
  end;


  // -------------------- Wait time --------------------

  if Cmd.IsOptionExists('t') then
    if not Cmd.IsOptionExists('w') then
      DisplayWarning('The value given in the "-t" option will be ignored. Reason: "-w" option not specified.')
    else
    begin
      s := LowerCase(Cmd.GetOptionValue('t'));
      if s = 'infinite' then AppParams.WaitTimeMs := INFINITE
      else
        if not TryGetMilliseconds(s, AppParams.WaitTimeMs, tuMillisecond) then
        begin
          DisplayError('Invalid value for the -t option: ' + s);
          Writeln('Cannot convert specified value to time.');
          ExitCode := TConsole.ExitCodeError;
          Terminate;
          Exit;
        end;
    end;



  // ------------ Window handle -------------

  if Cmd.IsOptionExists('hwnd') then
  begin
    s := Cmd.GetOptionValue('hwnd');
    if TStr.StartsWith('0x', s, True) then s := '$' + Copy(s, 3, Length(s));
    if not TryStrToInt(s, x) then
    begin
      DisplayError('Invalid value for the -hwnd option!' + ENDL + '"' + s + '" is not a valid integer value!');
      ExitCode := TConsole.ExitCodeError;
      Terminate;
      Exit;
    end;
  end;


  //---------------------------- FILE --------------------------
  if Cmd.UnknownParamCount = 0 then
  begin
    DisplayError('File name/URL not specified!');
    ExitCode := TConsole.ExitCodeError;
    Terminate;
    Exit;
  end
  else if Cmd.UnknownParamCount > 1 then
  begin
    DisplayError('Only one file name/URL was expected!');
    ExitCode := TConsole.ExitCodeError;
    Terminate;
    Exit;
  end
  else
    AppParams.&File := Cmd.UnknownParams[0].ParamStr;


end;

{$endregion ProcessOptions}



{$region '                    PerformMainAction                     '}
procedure TApp.PerformMainAction;
var
  sei: TShellExecuteInfo;
  see: TShellExecuteError;
begin
  if Terminated then Exit;

  FillChar(sei, SizeOf(sei), 0);
  sei.cbSize := SizeOf(sei);
  sei.fMask := SEE_MASK_NOCLOSEPROCESS or SEE_MASK_DOENVSUBST;
  sei.wnd := AppParams.WindowHandle;
  sei.lpDirectory := PChar(AppParams.Directory);
  sei.lpFile := PChar(AppParams.&File);
  sei.lpParameters := PChar(AppParams.Parameters);
  sei.lpVerb := PChar(AppParams.Operation);
  sei.nShow := AppParams.ShowCmd;

  //////////////////////////////////////////////////////////

  if ShellExecuteEx(@sei) then

  // ShellExecuteEx OK!
  begin
    case AppParams.WaitFor of
      wfFinish: WaitForSingleObject(sei.hProcess, DWORD(AppParams.WaitTimeMs));
      wfIdle: WaitForInputIdle(sei.hProcess, DWORD(AppParams.WaitTimeMs));
    end;
    CloseHandle(sei.hProcess);
    ExitCode := TConsole.ExitCodeOK;
  end

  else

  // ShellExecuteEx failed!
  begin
    if see.IsShellExecuteError(sei.hInstApp) then
    begin
      DisplayError('ShellExecute error no. ' + itos(sei.hInstApp) + ': ' + see.Description(sei.hInstApp));
      ExitCode := sei.hInstApp;
    end
    else
    begin
      if GetLastError <> 0 then
      begin
        DisplayError('OS Error No. ' + itos(GetLastError) + ': ' + SysErrorMessage(GetLastError));
        ExitCode := GetLastError;
      end
      else
        ExitCode := TConsole.ExitCodeError;
    end;
  end;

end;
{$endregion PerformMainAction}


{$region '                    Display... procs                  '}
procedure TApp.DisplayHelpAndExit(const ExCode: integer);
begin
  DisplayBanner;
  DisplayShortUsage;
  DisplayUsage;
  DisplayExtraInfo;
  DisplayExamples;

  ExitCode := ExCode;
  Terminate;
end;

procedure TApp.DisplayShortUsageAndExit(const Msg: string; const ExCode: integer);
begin
  if Msg <> '' then Writeln(Msg);
  DisplayShortUsage;
  DisplayTryHelp;
  ExitCode := ExCode;
  Terminate;
end;

procedure TApp.DisplayBannerAndExit(const ExCode: integer);
begin
  DisplayBanner;
  ExitCode := ExCode;
  Terminate;
end;

procedure TApp.DisplayMessageAndExit(const Msg: string; const ExCode: integer);
begin
  Writeln(Msg);
  ExitCode := ExCode;
  Terminate;
end;
{$endregion Display... procs}



end.
