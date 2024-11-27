program pLabeldruck;

uses
  Forms,
  Labeldruck in 'Labeldruck.pas' {Form7},
  Utilities in '..\common\Utilities.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm7, Form7);
  Application.Run;
end.
