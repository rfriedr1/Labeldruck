unit Utilities;


interface
uses
  Windows, Winspool, SysUtils, StdCtrls, ExtCtrls, Forms, Graphics, IniFiles, Grids, DBGrids, Dialogs;


type
  TDoc_Info_1 = record
    pDocName: pChar;
    pOutputFile: pChar;
    pDataType: pChar;
  end;

function StringToCaseSelect (Selector: string; CaseList: array of string): Integer;

procedure GridDeleteRow(RowNumber: Integer; Grid: TstringGrid);
procedure ResizeStringGrid(strg: TStringGrid; pW: integer);

procedure PrintLabel(PrinterName, Line1, Line2, Line3, Count, SampNr: string; DoBarcode: boolean);
procedure PrintLabelVar(PrinterName, Line1, Line2, Line3, Line4, Count, SampNr, L1L1, L1L2, L1L3, L1L4, L1L5, L1bc,
  L2L1, L2L2, L2L3, L2L4, L2L5, L2bc: string; DoBarcode: boolean; ReturnSample, CNIsotopA, ProjectNr, PrepMethods, LabelsTopOffset: string);

procedure SaveFileIniStr(const FName, Section, Key, strValue: string);
procedure SaveIniPath(const Path, Section, Key, strValue: string);

function ReadFileIniStrPath(const FName, Section, Key, sDefault: string): string;
function ReadIniPath(const Path, Section, Key, sDefault: string): string;

function ReplaceUmlaute(s: string): string;

{procedure ColourLabel(const lab: TLabel; const valid: boolean); //DigitChecker
function CheckIsInteger(const edit: TLabeledEdit; const lab: TLabel;   //DigitChecker
                        const min: integer = -maxint; const max: integer = maxint): boolean;
         }

implementation

const
   //Linie 1 - 4
  yAddressLine1: string = '0046'; // inch; Abstand in Y-Richtung (Linie 1) '0046'
  yAddressLine2: string = '0033'; // inch; Abstand in Y-Richtung (Linie 2) '0032'
  yAddressLine3: string = '0022'; // inch; Abstand in Y-Richtung (Linie 3) '0021'
  yAddressLine4: string = '0012'; // inch; Abstand in Y-Richtung (Linie 4) '0010'
  yAddressLine5: string = '0003'; // inch; Abstand in Y-Richtung (Linie 5)
  Label1xAddressLine1: string = '0011'; // inch; Abstand in X-Richtung (Label 1, Linie 1)
  Label1xAddressLine2: string = '0011'; // inch; Abstand in X-Richtung (Label 1, Linie 2,3)
  Label2xAddressLine1: string = '0140'; // inch; Abstand in X-Richtung (Label 2, Linie 1)
  Label2xAddressLine2: string = '0140'; // inch; Abstand in X-Richtung (Label 2, Linie 2,3)
   //Barcode
  //bcyAddress: string = '0006'; //inch, Abstand in Y-Richtung (barcode), barcode at the bottom
  bcyAddress: string = '0039'; //inch, Abstand in Y-Richtung (barcode), barcode at the top
  bcLabel1xAddress: string = '0015'; //inch
  bcLabel2xAddress: string = '0148'; //inch
  ReturnSampleyAddress: string = '0048';

procedure PrintRAW(PrinterName, PrintText: String);
// Printtext must be 8bit ascii and not simple 'string'
// maybe use AnsiString instead of UTF8String
// maybe also WritePrinter(Handle, @PrintText[1], Length(PrintText), N);

var
  DocInfo: TDoc_Info_1;
  Handle: THandle;
  //DocInfo: TDocInfo1;
  PrintTextAnsi: AnsiString;
  N: DWORD;

begin
  if not OpenPrinter(PChar(PrinterName), Handle, nil) then
  begin
    ShowMessage('Problem printing on: ' + PrinterName + ' (Handle: ' + Handle.ToString + ', Error: ' + SysErrorMessage(GetLastError()) + ')');
    Exit;
  end;
  with DocInfo do
  begin
    pDocName := PChar('test doc');
    pOutputFile := nil;
    pDataType := 'RAW';
  end;
  StartDocPrinter(Handle, 1, @DocInfo);
  StartPagePrinter(Handle);

  PrintTextAnsi:=PrintText; // change type from string to AnsiString
  //WritePrinter(Handle, PChar(PrintText), Length(PrintText), N);  //write to printer, change this
  WritePrinter(Handle, @PrintTextAnsi[1], Length(PrintTextAnsi), N);

  EndPagePrinter(Handle);
  EndDocPrinter(Handle);
  ClosePrinter(Handle);
end;

procedure PrintLabel(PrinterName, Line1, Line2, Line3, Count, SampNr: string; DoBarcode: boolean);
var
  s: AnsiString;   // String has to be 8bit AnsiString and not 'string'

  function PrinterStart: string;
  begin
    Result := #$02 + #$1B + 'G1' + #$0D + #$0A + //Selects proper command set for your application (48)
      #$02 + 'XD' + #$0D + #$0A + //Selects default module. If default is set with other command moduleparameter, the module selected with this command is used.(45)
      #$02 + 'qC' + #$0D + #$0A + //Clears default module.
      #$02 + 'Q' + #$0D + #$0A + //Clearing all memory module contents
      #$02 + 'n' + #$0D + #$0A + //inch system
      #$02 + 'c0000' + #$0D + #$0A + //paper length for continuous paper
      #$02 + 'e' + #$0D + #$0A + //edge sensor selection
      #$02 + 'M0548' + #$0D + #$0A + //Maximum label length
      #$02 + 'S3' + #$0D + #$0A + //Paper feed speed (40)
      #$02 + 'L' + #$0D + #$0A + //specifying printing contents setting start
      'D11' + #$0D + #$0A + //setting pixel size (57)
      'A2' + #$0D + #$0A + //Specifies development method for character and bar code.(54)|
      'C0000' + #$0D + #$0A +
      'ySWE'; //Selection of a TrueType font symbol set (47)
  end;

  function LineHeader: string;
  begin
    Result := #$0D + #$0A + #$1B + 'P00' + #$0D + #$0A; //P00:Specifying space between characters
  end;

  function BarcodeHeader: string;
  begin
    Result := #$0D + #$0A + '1o22016'; //Barcode
  end;

  function TextHeader: string;
  begin
    Result := '1911A08'; //1: no rotation; 9: Font Number; 11: Horizontal/Vertikal expansion; A08: Font size
  end;

  function TextEnd: string;
  begin
    Result := #$0D + #$0A +
      'sALABEL' + #$0D + #$0A + //Stores label format into memory module and ends label format.(67)
      #$02 + 'V0' + #$0D + #$0A + //With this command, memory switch contents can be changed temporarily.(43)
      #$02 + 'O0218' + #$0D + #$0A + //Setting printing position (35)
      #$02 + 'f294' + #$0D + #$0A + //peeling (cutting) position (23)
      #$02 + 'L' + #$0D + #$0A + //specifying printing contents setting start
      'rLABEL' + #$0D + #$0A + //Vielleicht: Detects label position automatically by reflective paper sensor(?)(39)
      'D11' + #$0D + #$0A + //setting pixel size (57)
      'H15' + #$0D + #$0A + //Sets print density (heat energy is applied to print head).(58)
      'P3' + #$0D + #$0A + //Setting printable area speed (60)
      'S3' + #$0D + #$0A + //Sets unprintable area speed. (66)
      'p3' + #$0D + #$0A; //Setting backfeed speed (61)
  end;

  function PrinterEnd: string;
  begin
    Result := #$0D + #$0A + 'E' + #$0D + #$0A; //Ends label format mode and prints
  end;

begin
  if DoBarcode = false then //CheckBox deaktiviert
  begin
    s := PrinterStart +
      LineHeader + TextHeader + yAddressLine1 + Label1xAddressLine1 + Line1 +
      LineHeader + TextHeader + yAddressLine3 + Label1xAddressLine2 + Line3 +
      LineHeader + TextHeader + yAddressLine2 + Label1xAddressLine2 + Line2 + //Labelwechsel
      LineHeader + TextHeader + yAddressLine1 + Label2xAddressLine1 + Line1 +
      LineHeader + TextHeader + yAddressLine3 + Label2xAddressLine2 + Line3 +
      LineHeader + TextHeader + yAddressLine2 + Label2xAddressLine2 + Line2 +
      TextEnd + Count + PrinterEnd; //Count: "Q000X" Setting number of prints (63)
  end
  else begin
    s := PrinterStart +
      LineHeader + TextHeader + yAddressLine1 + Label1xAddressLine1 + Line1 +
      BarcodeHeader + bcyAddress + bcLabel1xAddress + SampNr +
      LineHeader + TextHeader + yAddressLine3 + Label1xAddressLine2 + Line3 +
      LineHeader + TextHeader + yAddressLine2 + Label1xAddressLine2 + Line2 + //Labelwechsel
      LineHeader + TextHeader + yAddressLine1 + Label2xAddressLine1 + Line1 +
      BarcodeHeader + bcyAddress + bcLabel2xAddress + SampNr +
      LineHeader + TextHeader + yAddressLine3 + Label2xAddressLine2 + Line3 +
      LineHeader + TextHeader + yAddressLine2 + Label2xAddressLine2 + Line2 +
      TextEnd + Count + PrinterEnd;
  end;
  PrintRAW(PrinterName, s);
end;

// #############################################################
// Prozedur für pLabelDruckTabPages
// #$0D + #$0A = CR LF aka Windows Linebreak
// #############################################################
procedure PrintLabelVar(PrinterName, Line1, Line2, Line3, Line4, Count, SampNr, L1L1, L1L2, L1L3, L1L4, L1L5, L1bc,
  L2L1, L2L2, L2L3, L2L4, L2L5, L2bc: string; DoBarcode: boolean; ReturnSample, CNIsotopA, ProjectNr, PrepMethods, LabelsTopOffset: string);
var
  //t: string;
  t: AnsiString;   //use UTF8String that is UTF8 (8bit)


  function PrinterVarStart: string;
  begin
    Result := #$02 + #$1B + 'G1' + #$0D + #$0A + //Selects proper command set for your application (48)
      #$02 + 'XD' + #$0D + #$0A + //Selects default module. If default is set with other command moduleparameter, the module selected with this command is used.(45)
      #$02 + 'qC' + #$0D + #$0A + //Clears default module.
      #$02 + 'Q' + #$0D + #$0A + //Clearing all memory module contents
      #$02 + 'n' + #$0D + #$0A + //inch system
      #$02 + 'c0000' + #$0D + #$0A + //paper length for continuous paper
      #$02 + 'e' + #$0D + #$0A + //edge sensor selection
      #$02 + 'M0548' + #$0D + #$0A + //Maximum label length
      #$02 + 'S2' + #$0D + #$0A + //Paper feed speed (S3)
      #$02 + 'L' + #$0D + #$0A + //specifying printing contents setting start
      'D11' + #$0D + #$0A + //setting pixel size (57)
      'A2' + #$0D + #$0A + //Specifies development method for character and bar code.(54)|
      'C0000' + #$0D + #$0A +
      'ySWE'; //Selection of a TrueType font symbol set (47)
  end;

  function LineVarHeader: string;
  begin
    Result := #$0D + #$0A + #$1B + 'P00' + #$0D + #$0A; //P00:Specifying space between characters
  end;

  function FixedLineVarHeader1: string;
  // command to draw a line
  // 1X11,000,row(xxxx), column(xxxx)
  begin
    // Result := #$0D + #$0A + '1X1100000450013l01030001' + #$0D + #$0A; //dicke Linie
    Result := #$0D + #$0A + '1X1100000450013l00750001' + #$0D + #$0A; //dicke kurze Linie
  end;

  function FixedLineVarHeader2: string;
  begin
    // Result := #$0D + #$0A + '1X1100000450145l01030001' + #$0D + #$0A; //dicke Linie; 1X11000: fixed for Lines, 0045:X, 0145:Y, l01:size, 030:horiz. ausdehn., 001:vertikale ausd
    // Result := #$0D + #$0A + '1X1100000450145l01030001' + #$0D + #$0A; //dicke Linie; 1X11000: fixed for Lines, 0045:X, 0145:Y, l01:size, 025:horiz. ausdehn., 001:vertikale ausd
    Result := #$0D + #$0A + '1X1100000450145l0750001' + #$0D + #$0A;  // dicke kurze Linie
  end;

  function BarcodeVarHeader: string;
  // general layout of code: rotate(1), font(W1c), thick(5), narrow(5), height (000)
  begin
    //Result := #$0D + #$0A + '1W1c33000'; //DataMatrix: Modulbreite 3  -> also set L1bc = 0109 and L2bc = 0242
    Result := #$0D + #$0A + '1W1c66000'; //DataMatrix-Font: W1c, thick: 6, Narrow: 6, height: 000, also set L1bc = 0099 and L2bc = 0232
    //Result := #$0D + #$0A + '1W1c99000'; //DataMatrix: Modulbreite 9 -> also set L1bc = 0089 and L2bc = 0222
    //Result := #$0D + #$0A + '1o22016';  //Barcode    93
  end;

  function BarcodeVarEnd: string;
  begin
    Result := '2000000000';
  end;

  function ReturnSampleVarHeader: string;
  begin
    Result := '1911A04'; //1: no rotation; 9: Font Number; 11: Horizontal/Vertikal expansion; A04: Font size
  end;

  function ReturnSampleText(ReturnSample: string): string;
  begin
    // showmessage(ReturnSample);
    if ReturnSample = '1' then
      Begin
        Result := 'Ret';
      end
    else
      begin
        Result := '';
      end;
  end;

  function CNIsotopAVarEnd: string;
  begin
    //Result := '2000000000';
  end;

    function CNIsotopAVarHeader: string;
  begin
    Result := '1911A05'; //1: no rotation; 9: Font Number; 11: Horizontal/Vertikal expansion; A04: Font size (smallest)
  end;

  function CNIsotopAText(CNIsotopA: string): string;
  begin
    if CNIsotopA = '1' then
      Begin
        Result := 'CN'; //1: no rotation; 9: Font Number; 11: Horizontal/Vertikal expansion; A08: Font size
        if ReturnSample = '1' then Result:= ' | ' + Result;  // add a ',' between this and ReturnSample Flag
      end
    else
      begin
        Result := '';
      end;
  end;

  function ReturnSampleVarEnd: string;
  begin
    //Result := '2000000000';
  end;

  function TextVarHeader: string;
  // character field defintion (page 1-76) that defines how to print the text
  begin
    Result := '1911A08'; //1: no rotation; 9: Font Number; 1: Horizontal Expansion; 1:Vertikal Expansion expansion; A08: Font size in pt (only valid if font is set to 9)
  end;

  function TextVarHeaderProjectNr: string;
  // character field defintion (page 1-76) that defines how to print the text
  begin
    Result := '1911A06'; //1: no rotation; 9: Font Number; 1: Horizontal Expansion; 1:Vertikal Expansion expansion; A06: Font size in pt (only valid if font is set to 9)
  end;

  function TextVarEnd: string;
  begin
    Result := #$0D + #$0A +
      'sALABEL' + #$0D + #$0A + //Stores label format into memory module and ends label format.(67)
      #$02 + 'V0' + #$0D + #$0A + //With this command, memory switch contents can be changed temporarily.(43)
      #$02 + 'O' + LabelsTopOffset + #$0D + #$0A + //Setting printing position (35) 0120-0320 allowed   used to be: 'O0218'
      #$02 + 'f294' + #$0D + #$0A + //peeling (cutting) position (23)
      #$02 + 'L' + #$0D + #$0A + //specifying printing contents setting start
      'rLABEL' + #$0D + #$0A + //Vielleicht: Detects label position automatically by reflective paper sensor(?)(39)
      'D11' + #$0D + #$0A + //setting pixel size (57)
      'H15' + #$0D + #$0A + //Sets print density (heat energy is applied to print head).(58)
      'P3' + #$0D + #$0A + //Setting printable area speed (60)
      'S3' + #$0D + #$0A + //Sets unprintable area speed. (66)
      'p3' + #$0D + #$0A; //Setting backfeed speed (61)
  end;

  function PrinterVarEnd: string;
  begin
    Result := #$0D + #$0A + 'E' + #$0D + #$0A; //Ends label format mode and prints
  end;


  // generte the string of commands that will be send to the printer
  // PrinterVarStart: general printer setup
  // LineVarHeader: Specify general setting for this line
  // TextVarHeader: specify how exactly the text is being printed (font, size)
  // yAddress: y-Position (row number on label)
  // xAdress: x-Position (column number on label) e.g. L1L1
  // Linex: actual Text to be printed
  begin
  if DoBarcode = false then //CheckBox deaktiviert, don't print barcode
  begin
    t := PrinterVarStart +
      LineVarHeader + TextVarHeader + yAddressLine1 + L1L1 + Line1 +
      LineVarHeader + ReturnSampleVarHeader + ReturnSampleyAddress + '0070' + ReturnSampleText(ReturnSample) + CNIsotopAText(CNIsotopA) +
      LineVarHeader + TextVarHeader + yAddressLine4 + L1L4 + Line4 +
      LineVarHeader + TextVarHeader + yAddressLine3 + L1L3 + Line3 +
      LineVarHeader + TextVarHeader + yAddressLine2 + L1L2 + Line2 +
      LineVarHeader + TextVarHeaderProjectNr + yAddressLine5 + L1L5 + ProjectNr + ' / ' + PrepMethods +      // print project_nr and prep-methods in the same line
      FixedLineVarHeader1 + //Labelwechsel zu Label2
      LineVarHeader + TextVarHeader + yAddressLine1 + L2L1 + Line1 +
      LineVarHeader + ReturnSampleVarHeader + ReturnSampleyAddress + '0205' + ReturnSampleText(ReturnSample) + CNIsotopAText(CNIsotopA) +
      LineVarHeader + TextVarHeader + yAddressLine4 + L2L4 + Line4 +
      LineVarHeader + TextVarHeader + yAddressLine3 + L2L3 + Line3 +
      LineVarHeader + TextVarHeader + yAddressLine2 + L2L2 + Line2 +
      LineVarHeader + TextVarHeaderProjectNr + yAddressLine5 + L2L5 + ProjectNr + ' / ' + PrepMethods +         // print project_nr and prep-methods in the same line
      FixedLineVarHeader2 +
      TextVarEnd + Count + PrinterVarEnd; //Count: "Q000X" Setting number of prints (63)
  end
  else
  begin
    t := PrinterVarStart +
      LineVarHeader + TextVarHeader + yAddressLine1 + L1L1 + Line1 +
      LineVarHeader + ReturnSampleVarHeader + ReturnSampleyAddress + '0070' + ReturnSampleText(ReturnSample) + CNIsotopAText(CNIsotopA) +
      BarcodeVarHeader + bcyAddress + L1bc + BarcodeVarEnd + SampNr +
      LineVarHeader + TextVarHeader + yAddressLine4 + L1L4 + Line4 +
      LineVarHeader + TextVarHeader + yAddressLine3 + L1L3 + Line3 +
      LineVarHeader + TextVarHeader + yAddressLine2 + L1L2 + Line2 +
      LineVarHeader + TextVarHeaderProjectNr + yAddressLine5 + L1L5 + ProjectNr + ' / ' + PrepMethods +      // print project_nr and prep-methods in the same line
      FixedLineVarHeader1 + //Labelwechsel
      LineVarHeader + TextVarHeader + yAddressLine1 + L2L1 + Line1 +
      LineVarHeader + ReturnSampleVarHeader + ReturnSampleyAddress + '0205' + ReturnSampleText(ReturnSample) + CNIsotopAText(CNIsotopA) +
      BarcodeVarHeader + bcyAddress + L2bc + BarcodeVarEnd + SampNr +
      LineVarHeader + TextVarHeader + yAddressLine4 + L2L4 + Line4 +
      LineVarHeader + TextVarHeader + yAddressLine3 + L2L3 + Line3 +
      LineVarHeader + TextVarHeader + yAddressLine2 + L2L2 + Line2 +
      LineVarHeader + TextVarHeaderProjectNr + yAddressLine5 + L2L5 + ProjectNr + ' / ' + PrepMethods +     // print project_nr and prep-methods in the same line
      FixedLineVarHeader2 +
      TextVarEnd + Count + PrinterVarEnd;
  end;
  //ShowMessage(t);

  // send commands to printer
  if NOT (PrinterName = '') then
    begin
      PrintRAW(PrinterName, t);
    end;
end;

{procedure ColourLabel(const lab: TLabel; const valid: boolean);
const
colours: array[boolean] of TColor = (clRed, clWindowText);
begin
if Assigned(lab) then
lab.Font.Color := colours[valid];
end;

function CheckIsInteger(const edit: TLabeledEdit; const lab: TLabel; const min: integer = -MaxInt; const max: integer = maxint): boolean;
var
v: integer;
begin
try
v := StrToInt(edit.Text);
result := (min <= v) and (v <= max);
except
result := false;
end;
ColourLabel(lab, result);
end;          }

procedure ResizeStringGrid(strg: TStringGrid; pW: integer); //grid, parent
var
  i, w: integer;
begin
  with strg do begin
    w := 0;
    for I := 0 to ColCount - 1 do w := w + ColWidths[i] + 1; // 1 = gridline
    w := w + 7;
    ColWidths[ColCount - 1] := ColWidths[ColCount - 1] + pW - w;
  end;
end;

procedure GridDeleteRow(RowNumber: Integer; Grid: TstringGrid);
var
  i: Integer;
begin
  Grid.Row := RowNumber;
  if (Grid.Row = Grid.RowCount - 1) then
    { On the last row}
    Grid.RowCount := Grid.RowCount - 1
  else
  begin
    { Not the last row}
    for i := RowNumber to Grid.RowCount - 1 do
      Grid.Rows[i] := Grid.Rows[i + 1];
    Grid.RowCount := Grid.RowCount - 1;
  end;
end;

function StringToCaseSelect (Selector: string; CaseList: array of string): Integer;
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

procedure SaveFileIniStr(const FName, Section, Key, strValue: string);
var
  IniF: TIniFile;
begin
  IniF := TIniFile.Create(FName);
  IniF.WriteString(Section, Key, strValue);
  IniF.Free;
end;

procedure SaveIniPath(const Path, Section, Key, strValue: string);
var
  FName: string;
begin
  FName := ChangeFileExt(Path, '.INI');
  SaveFileIniStr(FName, Section, Key, strValue);
end;

function ReadIniPath(const Path, Section, Key, sDefault: string): string;
var
  FName: string;
begin
  FName := ChangeFileExt(Path, '.INI');
  result := ReadFileIniStrPath(FName, Section, Key, sDefault);
end;

function ReadFileIniStrPath(const FName, Section, Key, sDefault: string): string;
var
  IniF: TIniFile;
begin
  IniF := TIniFile.Create(FName);
  result := IniF.ReadString(Section, Key, sDefault);
  IniF.Free;
end;

function ReplaceUmlaute(s: string): string;
var i: integer;
begin
  result := '';
  for i := 1 to length(s) do
  begin
    Case s[i] of
    'ä': result := result+'ae';
    'ü': result := result+'ue';
    'ö': result := result+'oe';
    'ß': result := result+'ss';
    else result := result+s[i];
    end;
  end;
end;

end.

