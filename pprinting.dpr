program pprinting;

uses
  Forms,
  printing in 'printing.pas' {Labeldruck},
  Utilities in 'Utilities.pas',
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TLabeldruck, Labeldruck);
  Application.Run;
end.
