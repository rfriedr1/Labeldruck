unit printing;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ComCtrls, ExtCtrls, Mask, JvExMask, JvSpin, DB, ADODB, Dialogs, Grids,
  JvExGrids, JvStringGrid, JvComponentBase, JvHidControllerClass, jpeg,
  JvFormPlacement, JvJCLUtils, JvDialogs, FireDAC.UI.Intf, FireDAC.VCLUI.Wait,
  FireDAC.Stan.Intf, FireDAC.Comp.UI, Vcl.DBGrids, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.Phys.Intf, FireDAC.Stan.Def, FireDAC.Stan.Pool,
  FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.MySQL, FireDAC.Phys.MySQLDef,
  FireDAC.VCLUI.Login, FireDAC.VCLUI.Error, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt.Intf, FireDAC.DApt, FireDAC.Comp.DataSet, FireDAC.Comp.Client;

type
  TlblStatus = (busy, saved, notsaved);

const
  Version = 'v1.6 2016-09-15';

type
  TLabeldruck = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    btnCancel: TButton;
    edtSampleNrStart: TLabeledEdit;
    btnFind: TButton;
    cbBarcode: TCheckBox;
    btnPrint: TButton;
    btnReset: TButton;
    edspCopyCount: TJvSpinEdit;
    DataSource1: TDataSource;
    lbledLabel1Line1: TLabeledEdit;
    lbledLabel1Line3: TLabeledEdit;
    lbledLabel1bc: TLabeledEdit;
    lbledLabel1Line2: TLabeledEdit;
    lblDistance: TLabel;
    LabelRight: TLabel;
    lbledLabel2Line1: TLabeledEdit;
    lbledLabel2Line2: TLabeledEdit;
    lbledLabel2Line3: TLabeledEdit;
    lbledLabel2bc: TLabeledEdit;
    btnResetLeft: TButton;
    btnSave: TButton;
    btnResetRight: TButton;
    edMm: TEdit;
    btnConvert: TButton;
    edInch: TEdit;
    mm: TLabel;
    Label2: TLabel;
    lblNotSaved: TLabel;
    lblStatus: TLabel;
    Image1: TImage;
    lblCheck11: TLabel;
    lblCheck12: TLabel;
    lblCheck13: TLabel;
    lblCheck1bc: TLabel;
    lblCheck21: TLabel;
    lblCheck22: TLabel;
    lblCheck23: TLabel;
    lblCheck2bc: TLabel;
    lbledLabel1Line4: TLabeledEdit;
    lbledLabel2Line4: TLabeledEdit;
    DataSource2: TDataSource;
    TabSheet3: TTabSheet;
    Button1: TButton;
    grdPrintSet: TJvStringGrid;
    JvFormStorage1: TJvFormStorage;
    edtPrepStart: TJvSpinEdit;
    edtTargetStart: TJvSpinEdit;
    lblNprep: TLabel;
    lblNtar: TLabel;
    tbsGraphLabel: TTabSheet;
    StrGrdPrintGraph: TJvStringGrid;
    btnLoad: TButton;
    lblPrintPos: TLabel;
    btnPrintGraph: TButton;
    btnNextMAMS: TButton;
    Label6: TLabel;
    Label_Copies: TLabel;
    WaitCursor: TFDGUIxWaitCursor;
    GroupBox1: TGroupBox;
    edtSampleNrEnd: TLabeledEdit;
    DBGrid1: TDBGrid;
    FDConnection1: TFDConnection;
    FDGUIxWaitCursor1: TFDGUIxWaitCursor;
    FDGUIxLoginDialog1: TFDGUIxLoginDialog;
    FDGUIxErrorDialog1: TFDGUIxErrorDialog;
    FDQueryDBSample: TFDQuery;
    FDQueryDbUser: TFDQuery;
    FDPhysMySQLDriverLink1: TFDPhysMySQLDriverLink;
    procedure btnCancelClick(Sender: TObject);
    procedure btnFindClick(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure btnResetClick(Sender: TObject);
    procedure btnResetLeftClick(Sender: TObject);
    procedure btnResetRightClick(Sender: TObject);
    procedure lbledLabel1Line1Change(Sender: TObject);
    procedure lbledLabel1Line2Change(Sender: TObject);
    procedure lbledLabel1Line3Change(Sender: TObject);
    procedure lbledLabel1bcChange(Sender: TObject);
    procedure lbledLabel2Line1Change(Sender: TObject);
    procedure lbledLabel2Line2Change(Sender: TObject);
    procedure lbledLabel2Line3Change(Sender: TObject);
    procedure lbledLabel2bcChange(Sender: TObject);
    procedure btnConvertClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnHelpClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btnLoadClick(Sender: TObject);
    procedure btnPrintGraphClick(Sender: TObject);
    //procedure PrintTimerTimer(Sender: TObject);
    procedure PrintGraphClick(Sender: TObject);
    procedure btnNextMAMSClick(Sender: TObject);
    procedure lblPrintPosClick(Sender: TObject);
    procedure edtSampleNrStartChange(Sender: TObject);
    procedure FDConnection1AfterConnect(Sender: TObject);
  private
    User_Label, target_nr, prep_nr: string;
    counter, PrintPosition: integer;
    PrintFinished: boolean;
    function GetMaxPrepNrBySampleNr(sample_nr: integer): integer;
    function GetMaxTargetNrBySampleNr(sample_nr: integer): integer;
    procedure ChangeLabel(status: TlblStatus);
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Labeldruck: TLabeldruck;

implementation

{$R *.dfm}

uses Utilities, ConvUtils;

var
  L1L1, L1L2, L1L3, L1L4, L1bc, L2L1, L2L2, L2L3, L2L4, L2bc: string;

function StringToCaseSelect //DruckAnzahl str to case
  (Selector: string;
  CaseList: array of string): Integer;
var cnt: integer;
begin
  Result := -1;
  for cnt := 0 to Length(CaseList) - 1 do
  begin
    if CompareText(Selector, CaseList[cnt]) = 0 then
    begin
      Result := cnt;
      Break;
    end;
  end;
end;


procedure TLabeldruck.btnFindClick(Sender: TObject);
// query the database for the sample info
var
  s, t, Name: string;
  Ntar, Nprep: integer;
begin
  case StringToCaseSelect(edtSampleNrStart.Text,['']) of   //Wenn sample_nr leer -> abbrechen
    0: Exit;
  end;

  //btnQuery.Enabled := false;

  if edtSampleNrStart.Text=edtSampleNrEnd.Text then
  Begin
  // run this when Start and End Sample Nr are the same aka only one label is being printed
    //Query
    t := 'Select sample_nr, last_name, user_label, user_label_nr from user_t inner join project_t on project_t.user_nr=user_t.user_nr ' +
      'inner join sample_t on project_t.project_nr=sample_t.project_nr where sample_nr=' + edtSampleNrStart.Text + ';';
    FDQueryDbUser.SQL.Text := t;
    FDQueryDbUser.Open;

    if FDQueryDbUser.RecordCount > 0 then
    begin
    // something was returned from the database. record exists
    // display info in memo field

      // get max prep number of this SampleNr
      Nprep := GetMaxPrepNrBySampleNr(StrToInt(edtSampleNrStart.Text));
      if Nprep > 0 then
        lblNprep.Caption := 'max: ' + IntToStr(Nprep);
      // gte max number of available targets
      Ntar := GetMaxTargetNrBySampleNr(StrToInt(edtSampleNrStart.Text));
      if Ntar > 0 then
        lblNtar.Caption := 'max: ' + IntToStr(Ntar);
      edtPrepStart.Enabled := true;
      edtTargetStart.Enabled := true;

      btnPrint.Enabled := true;
    end
    else
      begin
      //the sample number was not found in the database
        ShowMessage('No records found!!');
        btnPrint.Enabled := false;
        //edInput.Enabled := false;
        Exit;
      end;
     end
  Else
    Begin
    // start and End Sample Number are not the same
    // display multiple sample info in Memo field
    // run query and ask for all samples in this range
    t := 'Select sample_nr, last_name, user_label, user_label_nr from user_t inner join project_t on project_t.user_nr=user_t.user_nr ' +
            'inner join sample_t on project_t.project_nr=sample_t.project_nr where sample_nr BETWEEN ' + edtSampleNrStart.Text + ' AND ' + edtSampleNrEnd.Text + ';';
    //ShowMessage(t);
    FDQueryDbUser.SQL.Text := t;
    FDQueryDbUser.Open;
    if FDQueryDbUser.RecordCount > 0 then
      Begin
      btnPrint.Enabled := true;
      End;
    End;

    DBGrid1.Columns[0].Width:=80;
    DBGrid1.Columns[1].Width:=100;
    DBGrid1.Columns[2].Width:=100;
    DBGrid1.Columns[3].Width:=100;

    ChangeLabel(saved);
    L1L1 := lbledLabel1Line1.Text;
    L1L2 := lbledLabel1Line2.Text;
    L1L3 := lbledLabel1Line3.Text;
    L1L4 := lbledLabel1Line4.Text;
    L1bc := lbledLabel1bc.Text;
    L2L1 := lbledLabel2Line1.Text;
    L2L2 := lbledLabel2Line2.Text;
    L2L3 := lbledLabel2Line3.Text;
    L2L4 := lbledLabel2Line4.Text;
    L2bc := lbledLabel2bc.Text;
    grdPrintSet.Cells[0, 1] := 'Label 1';
    grdPrintSet.Cells[0, 2] := 'Label 2';
    grdPrintSet.Cells[1, 0] := 'Linie 1';
    grdPrintSet.Cells[2, 0] := 'Linie 2';
    grdPrintSet.Cells[3, 0] := 'Linie 3';
    grdPrintSet.Cells[4, 0] := 'Barcode';
    grdPrintSet.Cells[1, 1] := L1L1;
    grdPrintSet.Cells[2, 1] := L1L2;
    grdPrintSet.Cells[3, 1] := L1L3;
    grdPrintSet.Cells[4, 1] := L1bc;
    grdPrintSet.Cells[5, 1] := L1L4;
    grdPrintSet.Cells[1, 2] := L2L1;
    grdPrintSet.Cells[2, 2] := L2L2;
    grdPrintSet.Cells[3, 2] := L2L3;
    grdPrintSet.Cells[4, 2] := L2bc;
    grdPrintSet.Cells[5, 2] := L2L4;
    //edInput.Enabled := false;

end;

procedure TLabeldruck.btnResetClick(Sender: TObject);
begin
  edtSampleNrStart.Clear; //Lˆscht Nummern-Eingabe-Feld
  edtSampleNrStart.Enabled := true; //Aktiviert Nummern-Eingabe-Feld
  btnFind.Enabled := true; //Aktiviert Abfrage-Button
  btnPrint.Enabled := false; //Deaktiviert Druck-Button
  cbBarcode.Checked := false; //Deaktiviert Checkbox
  edspCopyCount.Value := 2; //Anzahl zur¸cksetzten
  grdPrintSet.Clear;
  edtPrepStart.Value := 1;
  edtTargetStart.Value := 1;
  edtPrepStart.Enabled := false;
  edtTargetStart.Enabled := false;
  lblNprep.Caption := ' ';
  lblNtar.Caption := ' ';
end;



procedure TLabeldruck.btnResetLeftClick(Sender: TObject); //Wiederherstellen Label1 Einstellungen
begin
  lbledLabel1Line1.Text := '0006';
  lbledLabel1Line2.Text := '0006';
  lbledLabel1Line3.Text := '0006';
  lbledLabel1bc.Text := '0017';
end;

procedure TLabeldruck.btnResetRightClick(Sender: TObject); //Wiederherstellen Label2 Einstellungen
begin
  lbledLabel2Line1.Text := '0138';
  lbledLabel2Line2.Text := '0138';
  lbledLabel2Line3.Text := '0138';
  lbledLabel2bc.Text := '0150';
end;

procedure TLabeldruck.btnSaveClick(Sender: TObject);
  function CheckNumber(str: string): boolean;
  var dummy: integer;
  begin
    try
      begin
        dummy := StrToInt(str);
        Result := false;
      end;
    except
      Result := true;
    end;
  end;
begin
  begin
    if CheckNumber(lbledLabel1Line1.Text) = true then
    begin
     // lblCheck11.Visible := true;
      lbledLabel1Line1.EditLabel.Color := clRed;
    end
    else begin
      lblCheck11.Visible := false;
      lbledLabel1Line1.EditLabel.Color := clBtnFace;
    end;
// eigentlich unnˆtig for v := 1 to 3 ... Bernd Fragen!
    begin
      if CheckNumber(lbledLabel1Line2.Text) = true then
      begin
   //     lblCheck12.Visible := true;
        lbledLabel1Line2.EditLabel.Color := clRed;
      end
      else begin
        lblCheck12.Visible := false;
        lbledLabel1Line2.EditLabel.Color := clBtnFace;
      end;
    end;
    begin
      if CheckNumber(lbledLabel1Line3.Text) = true then
      begin
    //    lblCheck13.Visible := true;
        lbledLabel1Line3.EditLabel.Color := clRed;
      end
      else begin
        lblCheck13.Visible := false;
        lbledLabel1Line3.EditLabel.Color := clBtnFace;
      end;
    end;
    begin
      if CheckNumber(lbledLabel1Line4.Text) = true then
      begin
   //     lblCheck21.Visible := true;
        lbledLabel2Line1.EditLabel.Color := clRed;
      end
      else begin
        lblCheck21.Visible := false;
        lbledLabel2Line1.EditLabel.Color := clBtnFace;
      end;
    end;
    begin
      if CheckNumber(lbledLabel1bc.Text) = true then
      begin
   //     lblCheck1bc.Visible := true;
        lbledLabel1bc.EditLabel.Color := clRed;
      end
      else begin
     //   lblCheck1bc.Visible := false;
        lbledLabel1bc.EditLabel.Color := clBtnFace;
      end;
    end;
    begin
      if CheckNumber(lbledLabel2Line1.Text) = true then
      begin
      //  lblCheck21.Visible := true;
        lbledLabel2Line1.EditLabel.Color := clRed;
      end
      else begin
        lblCheck21.Visible := false;
        lbledLabel2Line1.EditLabel.Color := clBtnFace;
      end;
    end;


// eigentlich unnˆtig for v := 1 to 3 ... Bernd Fragen!

    begin
      if CheckNumber(lbledLabel2Line2.Text) = true then
      begin
      //  lblCheck22.Visible := true;
        lbledLabel2Line2.EditLabel.Color := clRed;
      end
      else begin
        lblCheck22.Visible := false;
        lbledLabel2Line2.EditLabel.Color := clBtnFace;
      end;
    end;
    begin
      if CheckNumber(lbledLabel2Line3.Text) = true then
      begin
      //  lblCheck23.Visible := true;
        lbledLabel2Line3.EditLabel.Color := clRed;
      end
      else begin
     //   lblCheck23.Visible := false;
        lbledLabel2Line3.EditLabel.Color := clBtnFace;
      end;
    end;
    begin
      if CheckNumber(lbledLabel2Line4.Text) = true then
      begin
     //   lblCheck21.Visible := true;
        lbledLabel2Line4.EditLabel.Color := clRed;
      end
      else begin
        lblCheck21.Visible := false;
        lbledLabel2Line4.EditLabel.Color := clBtnFace;
      end;
    end;
    begin
      if CheckNumber(lbledLabel2bc.Text) = true then
      begin
        lblCheck2bc.Visible := true;
        lbledLabel2bc.EditLabel.Color := clRed;
      end
      else begin
        lblCheck2bc.Visible := false;
        lbledLabel2bc.EditLabel.Color := clBtnFace;
      end;
    end;
  end;
  ChangeLabel(saved);
  L1L1 := lbledLabel1Line1.Text;
  L1L2 := lbledLabel1Line2.Text;
  L1L3 := lbledLabel1Line3.Text;
  L1L4 := lbledLabel1Line4.Text;
  L1bc := lbledLabel1bc.Text;
  L2L1 := lbledLabel2Line1.Text;
  L2L2 := lbledLabel2Line2.Text;
  L2L3 := lbledLabel2Line3.Text;
  L2L4 := lbledLabel1Line4.Text;
  L2bc := lbledLabel2bc.Text;
  grdPrintSet.Cells[0, 1] := 'Label 1';
  grdPrintSet.Cells[0, 2] := 'Label 2';
  grdPrintSet.Cells[1, 0] := 'Linie 1';
  grdPrintSet.Cells[2, 0] := 'Linie 2';
  grdPrintSet.Cells[3, 0] := 'Linie 3';
  grdPrintSet.Cells[4, 0] := 'Barcode';
  grdPrintSet.Cells[1, 1] := L1L1;
  grdPrintSet.Cells[2, 1] := L1L2;
  grdPrintSet.Cells[3, 1] := L1L3;
  grdPrintSet.Cells[4, 1] := L1bc;
  grdPrintSet.Cells[1, 2] := L2L1;
  grdPrintSet.Cells[2, 2] := L2L2;
  grdPrintSet.Cells[3, 2] := L2L3;
  grdPrintSet.Cells[4, 2] := L2bc;
  btnSave.Enabled := false;
end;


procedure TLabeldruck.Button1Click(Sender: TObject);
var
  UserLabel1, UserLabel2: string;
begin
  //Debug
  if Length(User_Label) > 20 then begin
    ShowMessage(User_Label);
    UserLabel1 := Copy(User_Label, 0, 18);
    UserLabel2 := Copy(User_Label, 19, Length(User_Label));
    ShowMessage(UserLabel1 + '-->1  2<---' + UserLabel2);
  end;
end;

procedure TLabeldruck.PrintGraphClick(Sender: TObject);
var
  i: integer;
  UserLabel, UserLabel1, UserLabel2, User, UserLabelNr, Count, Number, PrintString: string;
  Barcode: boolean;
begin

  L1L1 := lbledLabel1Line1.Text;
  L1L2 := lbledLabel1Line2.Text;
  L1L3 := lbledLabel1Line3.Text;
  L1L4 := lbledLabel1Line4.Text;
  L1bc := lbledLabel1bc.Text;
  L2L1 := lbledLabel2Line1.Text;
  L2L2 := lbledLabel2Line2.Text;
  L2L3 := lbledLabel2Line3.Text;
  L2L4 := lbledLabel2Line4.Text;
  L2bc := lbledLabel2bc.Text;

for i := 1 to StrGrdPrintGraph.RowCount-1 do begin
    PrintPosition := 1;
    with StrGrdPrintGraph do begin
      //i := PrintPosition;
      lblPrintPos.Caption := 'Printing... ' + IntToStr(PrintPosition);
      User_Label := Cells[5, i];
      //ShowMessage('User label gemacht1');
      if Length(User_Label) > 20 then begin
        UserLabel1 := Copy(User_Label, 0, 18);
        UserLabel2 := Copy(User_Label, 19, Length(User_Label));
      end
      else begin
        UserLabel2 := Copy(User_Label, 0, 20); //Copy: Beschr‰nken auf 20 Zeichen
        UserLabel1 := ' ';
      end;
      //ShowMessage('User label gemacht2');
      User := Copy(Cells[4, i], 0, 20);
      Count := '1';
      PrintString := Cells[0, i] + '.' + Cells[1, i] + '.' + Cells[2, i] + ' ' + Cells[3, i];
      Barcode := false; //Barcode wird nicht gedruckt
      //ShowMessage('Vor dem Printstring');
      (*PrintLabelVar('Citizen_CLP_631', PrintString, UserLabel1, UserLabel2, //Printer, L1, L2
      User, 'Q000' + Count, edInput.Text, L1L1, L1L2, L1L3, L1L4, L1bc,
      L2L1, L2L2, L2L3, L2L4, L2bc, Barcode); *)

      PrintLabelVar('Citizen_CLP_631', PrintString, UserLabel1, UserLabel2, //Printer, L1, L2
        User, 'Q000' + Count, Cells[0, i], L1L1, L1L2, L1L3, L1L4, L1bc,
        L2L1, L2L2, L2L3, L2L4, L2bc, Barcode);

        sleep(500);
    end;
    end;
end;

procedure TLabeldruck.ChangeLabel(status: TlblStatus);
begin
  lblStatus.Visible := false;
  case status of
    busy: begin
    //
      end;
    saved: begin
        lblStatus.Color := clGreen;
        lblStatus.Caption := 'Druck-Einstellungen gespeichert';
      end;
    notsaved: begin
        lblStatus.Color := clRed;
        lblStatus.Caption := 'Druck-Einstellungen NICHT gespeichert';
      end;
  end;
  //lblStatus.Visible := true;
end;

procedure TLabeldruck.edtSampleNrStartChange(Sender: TObject);
begin
 // add the value to another edit field
 edtSampleNrEnd.Text:=edtSampleNrStart.Text;
 edtPrepStart.Value:=1;
 edtTargetStart.Value:=1;
end;

procedure TLabeldruck.FDConnection1AfterConnect(Sender: TObject);
begin
  // a connection to the database has been established
  // safe parameters in file

  // FDConnection1.Params.SaveToFile('connections.ini');       // password is saved as plain text!!!!!!!
end;

procedure TLabeldruck.FormShow(Sender: TObject);
begin
  Labeldruck.Caption := 'Label Printing ' + Version;

  // perform some kind of startup procedure
  // maybe checking the database connection
  // and the connection to the server for the graphitization labels

end;

function TLabeldruck.GetMaxPrepNrBySampleNr(sample_nr: integer): integer;
var
  s: string;
begin
  with FDQueryDbSample do
  begin
    Close;
    SQL.Text := 'SELECT Max(prep_nr) FROM preparation_t WHERE sample_nr=' + IntToStr(sample_nr) + ';';
    s := SQL.Text;
    Open;
    Result := Fields.Fields[0].AsInteger;
  end;
end;

function TLabeldruck.GetMaxTargetNrBySampleNr(sample_nr: integer): integer;
var
  s: string;
begin
  with FDQueryDbSample do
  begin
    Close;
    SQL.Text := 'SELECT Max(target_nr) FROM target_t WHERE sample_nr=' + IntToStr(sample_nr) + ';';
    s := SQL.Text;
    Open;
    Result := Fields.Fields[0].AsInteger;
  end;
end;

procedure TLabeldruck.lbledLabel1bcChange(Sender: TObject);
begin
  ChangeLabel(notsaved);
  btnSave.Enabled := true;
end;

procedure TLabeldruck.lbledLabel1Line1Change(Sender: TObject);
begin
  ChangeLabel(notsaved);
  btnSave.Enabled := true;
end;

procedure TLabeldruck.lbledLabel1Line2Change(Sender: TObject);
begin
  ChangeLabel(notsaved);
  btnSave.Enabled := true;
end;

procedure TLabeldruck.lbledLabel1Line3Change(Sender: TObject);
begin
  ChangeLabel(notsaved);
  btnSave.Enabled := true;
end;

procedure TLabeldruck.lbledLabel2bcChange(Sender: TObject);
begin
  ChangeLabel(notsaved);
  btnSave.Enabled := true;
end;

procedure TLabeldruck.lbledLabel2Line1Change(Sender: TObject);
begin
  ChangeLabel(notsaved);
  btnSave.Enabled := true;
end;

procedure TLabeldruck.lbledLabel2Line2Change(Sender: TObject);
begin
  ChangeLabel(notsaved);
  btnSave.Enabled := true;
end;

procedure TLabeldruck.lbledLabel2Line3Change(Sender: TObject);
begin
  ChangeLabel(notsaved);
  btnSave.Enabled := true;
end;

procedure TLabeldruck.lblPrintPosClick(Sender: TObject);
begin

end;

(*procedure TLabeldruck.PrintTimerTimer(Sender: TObject);
// print timer runs avery 1 second if enabled
var
  i: integer;
  UserLabel, UserLabel1, UserLabel2, User, UserLabelNr, Count, Number, PrintString: string;
  Barcode: boolean;
begin
  PrintTimer.Enabled := false; //disable timer
  dec(counter);  //timer runs 3 times before the actual print happens (counter<0)
  if counter <= 0 then begin
  // send data to printer
    with StrGrdPrintGraph do begin
      i := PrintPosition;
      lblPrintPos.Caption := 'Drucke... ' + IntToStr(PrintPosition);
      User_Label := Cells[5, i]; //Access violation happens here
      if Length(User_Label) > 20 then begin
        UserLabel1 := Copy(User_Label, 0, 18);
        UserLabel2 := Copy(User_Label, 19, Length(User_Label));
      end else begin
        UserLabel2 := Copy(User_Label, 0, 20); //Copy: Beschr‰nken auf 20 Zeichen
        UserLabel1 := ' ';
      end;
      User := Copy(Cells[4, i], 0, 20);
      Count := '1';
      PrintString := Cells[0, i] + '.' + Cells[1, i] + '.' + Cells[2, i] + ' ' + Cells[3, i];
      Barcode := false; //Barcode wird nicht gedruckt
      PrintLabelVar('Citizen_CLP_631', PrintString, UserLabel1, UserLabel2, //Printer, L1, L2
        User, 'Q000' + Count, edInput.Text, L1L1, L1L2, L1L3, L1L4, L1bc,
        L2L1, L2L2, L2L3, L2L4, L2bc, Barcode);
      inc(PrintPosition);
      if i=RowCount - 1 then
        PrintFinished := true;

    end;
  end;
  if not PrintFinished then
    PrintTimer.Enabled := true //if print is not finished (within the first two runs of the routine) run timer again, otherwise keep time disabled
  else begin
    Counter := 3;
    lblPrintPos.Caption := 'Druckvorgang beendet.';
    btnLoad.Enabled := true;
  end;
end;  *)

procedure TLabeldruck.btnConvertClick(Sender: TObject); //Umrechnen von mm in inches
const
  conversion: double = 3.937;
var
  mmNum: double;
  aRes: double;
  Res: string;
begin
  case StringToCaseSelect(edMm.Text, //Wenn leer -> abbrechen
    ['']) of
    0: Exit;
  end;
  mmNum := StrToFloat(edMm.Text);
  aRes := (mmNum * conversion);
  Res := format('%5.0f', [aRes]);
  edInch.Text := Res;
end;

procedure TLabeldruck.btnHelpClick(Sender: TObject);
begin
  ShowMessage('.');
end;

procedure TLabeldruck.btnLoadClick(Sender: TObject);
var
  i, j: integer;
  s, t, fpath: string;
  Ntar, Nprep: integer;
begin
  lblPrintPos.Caption := 'Searching for Dataset... please wait.';
    fpath := ExtractFilePath(Application.ExeName) + 'Labels\label';
  with StrGrdPrintGraph do begin
    ColCount := 6;
    RowCount := 11;
    Cells[0, 0] := 'Sample';
    Cells[1, 0] := 'prep';
    Cells[2, 0] := 'tar';
    Cells[3, 0] := 'R';
    Cells[4, 0] := 'Name';
    Cells[5, 0] := 'user_label';
    //Sample
    for j := 1 to  RowCount - 1 do begin
      Cells[0, j] := ExtractWord(j, ReadIniPath(fpath, 'Data', 'Sample', ''), [' ']);
    end;
    //prep
    for j := 1 to  RowCount - 1 do begin
      Cells[1, j] := ExtractWord(j, ReadIniPath(fpath, 'Data', 'prep', ''), [' ']);
    end;
    //tar
    for j := 1 to  RowCount - 1 do begin
      Cells[2, j] := ExtractWord(j, ReadIniPath(fpath, 'Data', 'tar', ''), [' ']);
    end;
    //R
    for j := 1 to  RowCount - 1 do begin
      Cells[3, j] := ExtractWord(j, ReadIniPath(fpath, 'Data', 'Reactor', ''), [' ']);
    end;
    for i := 1 to RowCount - 1 do begin
      if not (Cells[0, i] = 'empty') then begin
        FDQueryDbUser.Close;
        t := 'Select last_name, user_label, user_label_nr from user_t inner join project_t on project_t.user_nr=user_t.user_nr ' +
          'inner join sample_t on project_t.project_nr=sample_t.project_nr where sample_nr=' + Cells[0, i] + ';';
        FDQueryDbUser.SQL.Text := t;
        FDQueryDbUser.Open;
        if FDQueryDbUser.RecordCount > 0 then
        begin
          Cells[5, i] := FDQueryDbUser.Fields.FieldByName('user_label').AsString + ' ' + FDQueryDbUser.Fields.FieldByName('user_label_nr').AsString;
          Cells[4, i] := FDQueryDbUser.Fields.FieldByName('last_name').AsString;
        end
        else begin
          ShowMessage('No record found for sample ' + Cells[0, i] + '.');
          lblPrintPos.Caption := 'Error!';
          Exit;
        end;
        FDQueryDbUser.Close;
      end;
    end;
  end;
  lblPrintPos.Caption := 'Label found.';
end;


procedure TLabeldruck.btnNextMAMSClick(Sender: TObject);
// query database and look for the next sample_nr
var
  n : integer;
begin
  n := StrToInt(edtSampleNrStart.Text);  // get sample_nr
  inc(n);                       // increment sample_nr by one
  edtSampleNrStart.Text := IntToStr(n);  // display new sample_nr
  btnFindClick(Self);          // query db
end;

procedure TLabeldruck.btnPrintClick(Sender: TObject);
// print sample labels
var
  UserLabel, UserLabel1, UserLabel2, User, UserLabelNr, Count, Number, PrintString: string;
  Barcode: boolean;
  i: integer;
begin

//      if Length(User_Label) > 20 then
//      begin
//        UserLabel1 := Copy(User_Label, 0, 18);
//        UserLabel2 := Copy(User_Label, 19, Length(User_Label));
//      end else
//      begin
//        UserLabel2 := Copy(User_Label, 0, 20); //Copy: Beschr‰nken auf 20 Zeichen
//        UserLabel1 := ' ';
//      end;

  If DBGrid1.DataSource.DataSet.RecordCount >0 Then
  Begin

      // go through all Records and print them
      DBGrid1.DataSource.DataSet.First;
      while DBGrid1.DataSource.DataSet.Eof=false do
      Begin
      //  UserLabelNr := Copy(FDQueryDbSample.Fields.FieldByName('User_Label_nr').AsString, 0, 20);
      // set dataset to first value
      User := Copy(DBGrid1.DataSource.DataSet.Fields.FieldByName('last_name').AsString, 0, 20);
      Number := edspCopyCount.Text;
      case StringToCaseSelect(Number, //Umwandeln von Text zu den Anzahl Kopien
        ['2', '4', '6', '8', '10']) of
        0: Count := '1';
        1: Count := '2';
        2: Count := '3';
        3: Count := '4';
        4: Count := '5';
      end;
      if Length(DBGrid1.DataSource.DataSet.Fields.FieldByName('user_label').AsString) > 20 then
        begin
          UserLabel1 := Copy(DBGrid1.DataSource.DataSet.Fields.FieldByName('user_label').AsString, 0, 18);
          UserLabel2 := Copy(DBGrid1.DataSource.DataSet.Fields.FieldByName('user_label').AsString, 19, Length(DBGrid1.DataSource.DataSet.Fields.FieldByName('user_label').AsString));
        end
      else
        begin
          UserLabel2 := Copy(DBGrid1.DataSource.DataSet.Fields.FieldByName('user_label').AsString, 0, 20); //Copy: Beschr‰nken auf 20 Zeichen
          UserLabel1 := ' ';
        end;
      PrintString := DBGrid1.DataSource.DataSet.Fields.FieldByName('sample_nr').AsString + '.' + edtPrepStart.Text + '.' + edtTargetStart.Text;
      Barcode := cbBarcode.Checked; //Barcode wird gedruckt
    {  PrintLabelVar('Citizen_CLP_631', 'MAMS ' + edInput.Text + ' - ', UserLabel, //Printer, L1, L2
        UserLabelNr, User, 'Q000' + Count, edInput.Text, L1L1, L1L2, L1L3, L1L4, L1bc,
        L2L1, L2L2, L2L3, L2L4, L2bc, Barcode);}
      PrintLabelVar('Citizen_CLP_631', PrintString, UserLabel1, UserLabel2, //Printer, L1, L2
        User, 'Q000' + Count, edtSampleNrStart.Text, L1L1, L1L2, L1L3, L1L4, L1bc,
        L2L1, L2L2, L2L3, L2L4, L2bc, Barcode);
      DBGrid1.DataSource.DataSet.Next;
  End;
  End;
end;

procedure TLabeldruck.btnPrintGraphClick(Sender: TObject);
// print graphitizer labels
var
  i: integer;
begin
 counter := 3;
 PrintPosition := 1;
 lblPrintPos.Caption := 'Druckvorgang gestartet.';
 //PrintTimer.Enabled := true;
end;

procedure TLabeldruck.btnCancelClick(Sender: TObject); //Programm schlieﬂen
begin
  Close;
end;
end.

