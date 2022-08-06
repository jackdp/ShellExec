unit SHEX.Procs;


interface

uses
  SysUtils,
  JPL.Strings, JPL.Conversion;


function ColorTag(const Text, Colors: string): string;
function CompilerVersionToDelphiName(const CompilerVer: Extended): string;



implementation


function ColorTag(const Text, Colors: string): string;
begin
  Result := '<color=' + Colors + '>' + Text + '</color>';
end;

function CompilerVersionToDelphiName(const CompilerVer: Extended): string;
var
  x: integer;
begin
  x := Round(CompilerVer * 10);

  if x > 350 then Result := 'Alexandria+'
  else
    case x of
      140: Result := '6';
      150: Result := '7';
      160: Result := '8 .NET';
      170: Result := '2005';
      180: Result := '2006';
      185: Result := '2007 Win32';
      190: Result := '2007 .NET';
      200: Result := '2009';
      210: Result := '2010';
      220: Result := 'XE';
      230: Result := 'XE2';
      240: Result := 'XE3';
      250: Result := 'XE4';
      260: Result := 'XE5';
      270: Result := 'XE6';
      280: Result := 'XE7';
      290: Result := 'XE8';
      300: Result := '10.0 Seattle';
      310: Result := '10.1 Berlin';
      320: Result := '10.2 Tokyo';
      330: Result := '10.3 Rio';
      340: Result := '10.4 Sydney';
      350: Result := '11.0 Alexandria';
    else
      Result := '';
    end;
end;



end.

