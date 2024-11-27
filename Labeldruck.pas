unit Labeldruck;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, DB, ADODB, Mask, JvExMask, JvSpin, Menus;

type
  TForm7 = class(TForm)
    ADOConnection1: TADOConnection;
    qryDb: TADOQuery;
    DataSource1: TDataSource;
    edEinsenderNummer: TLabeledEdit;
    btnQuery: TButton;
    btnPrint: TButton;
    Memo1: TMemo;
    edAnzahlKopien: TJvSpinEdit;
    Label1: TLabel;
    btnNeueAbfrage: TButton;
    cbBarcode: TCheckBox;
    Button1: TButton;
    procedure btnQueryClick(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure btnNeueAbfrageClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form7: TForm7;

implementation

{$R *.dfm}

uses Utilities;

procedure TForm7.btnQueryClick(Sender: TObject);
var
   s, Name : string;
begin
   s := 'SELECT SampName, SampUserCode FROM Sample WHERE SampNr=' + edEinsenderNummer.Text + ';' ;
   qryDb.SQL.Text := s;
   qryDb.Open;
   if qryDb.RecordCount > 0 then
    begin
        Name := qryDb.Fields.FieldByName('SampName').AsString;
        Memo1.Lines.Add('Probenname = ' + Name);
        Memo1.Lines.Add('UserCode = ' + qryDb.Fields.FieldByName('SampUserCode').AsString);
        Memo1.Lines.Add('------');
        Memo1.Lines.Add('Drucken: Druckt Anzahl Etiketten.');
        Memo1.Lines.Add('Neue Abfrage: Neue Abfrage starten (Ohne Drucken).');
        btnQuery.Enabled:= false;
        btnPrint.Enabled:= true;
    end
   else
    begin
        ShowMessage(edEinsenderNummer.Text + ': Keine Probe gefunden!');
        Exit;
    end;
    edEinsenderNummer.Enabled:= false;
end;


procedure TForm7.btnNeueAbfrageClick(Sender: TObject);
  begin
   Memo1.Clear;                     //Löscht alles im Memo-Feld
   edEinsenderNummer.Clear;         //Löscht Nummern-Eingabe-Feld
   edEinsenderNummer.Enabled:= true;//Aktiviert Nummern-Eingabe-Feld
   btnQuery.Enabled:= true;         //Aktiviert Abfrage-Button
   btnPrint.Enabled:= false;        //Deaktiviert Druck-Button
   cbBarcode.Checked:= false;       //Deaktiviert Checkbox
   edAnzahlKopien.Clear;
  end;

function StringToCaseSelect
   (Selector : string;
CaseList: array of string): Integer;
var cnt: integer;
begin
   Result:=-1;
   for cnt:=0 to Length(CaseList)-1 do
begin
     if CompareText(Selector, CaseList[cnt]) = 0 then
     begin
       Result:=cnt;
       Break;
     end;
   end;
end;

{
Usage:

case StringToCaseSelect('Delphi',
      ['About','Borland','Delphi']) of
   0:ShowMessage('You''ve picked About') ;
   1:ShowMessage('You''ve picked Borland') ;
   2:ShowMessage('You''ve picked Delphi') ;
end;
}

procedure TForm7.btnPrintClick(Sender: TObject);
  var
   SampleName, SampleUserCode, Count, Number : string;
   Barcode: boolean;
  begin
   SampleName := Copy(qryDb.Fields.FieldByName('SampName').AsString,0,20);   //Copy: Beschränken auf 20 Zeichen
   SampleUserCode := Copy(qryDb.Fields.FieldByName('SampUserCode').AsString,0,20);
   Number := edAnzahlKopien.Text;            
   case StringToCaseSelect(Number,           //Umwandeln von Text zu den Anzahl Kopien
      ['2','4','6','8','10']) of
      0: Count := '1';
      1: Count := '2';
      2: Count := '3';
      3: Count := '4';
      4: Count := '5';
   end;
   Barcode := cbBarcode.Checked;  //Barcode wird gedruckt
   PrintLabel('Citizen_CLP_631', 'HD ' + edEinsenderNummer.Text + ' - ', SampleName, //Printer, L1, L2
               SampleUserCode,  'Q000' + Count , edEinsenderNummer.Text, Barcode); //L3, Labelzahl, Barcodewert, Barcode
   btnPrint.Enabled:= false;
  end;

procedure TForm7.Button1Click(Sender: TObject);
begin
 Close;
end;
end.
