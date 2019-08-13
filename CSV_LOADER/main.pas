unit main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, StdCtrls, ExtCtrls, ComCtrls, Buttons, DB, ADODB, FileCtrl, Registry ;

type
  TMainForm = class(TForm)
    StringGrid1: TStringGrid;
    Panel1: TPanel;
    Load_All_Btn: TButton;
    Panel2: TPanel;
    FileListBox: TFileListBox;
    tablename_ed: TEdit;
    ADOConnection1: TADOConnection;
    Qdelete: TADOQuery;
    Qinsert: TADOQuery;
    Load_Cur_Btn: TButton;
    spid_ed: TEdit;
    QSchema: TADOQuery;
    Panel3: TPanel;
    Patch_Ed: TEdit;
    BitBtn1: TBitBtn;
    Label1: TLabel;
    Log_memo: TMemo;
    ServerSet_Btn: TButton;
    Label2: TLabel;
    Label3: TLabel;
    ok_lb: TLabel;
    false_lb: TLabel;
    count_lb: TLabel;

    procedure Load_All_BtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure FileListBoxClick(Sender: TObject);
    procedure Load_Cur_BtnClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ServerSet_BtnClick(Sender: TObject);
    procedure ADOConnection1AfterConnect(Sender: TObject);
    procedure ADOConnection1BeforeDisconnect(Sender: TObject);
  private
    procedure ReadTabFile(FN: TFileName; FieldSeparator:Char; SG: TStringGrid);
    procedure ShowFileList;
    function  SaveToBase(SG: TStringGrid; TableName: string; spid: string):boolean;
    procedure SaveSettings;
    procedure LoadSettings;
    procedure Log(Mes:string);
    function  CheckConnect:boolean;
    procedure SetConLabel(Server, Base: string);
    procedure setFlag(f:smallint);
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;



implementation
  uses clipbrd, ADOConEd;
{$R *.dfm}

procedure TMainForm.setFlag(f:smallint);
begin
  case f of
    0: begin
         ok_lb.Font.color    := clGrayText;
         false_lb.Font.color := clGrayText;
       end;
    1: begin
         ok_lb.Font.color    := clLime;
         false_lb.Font.color := clGrayText;
       end;
    2: begin
         ok_lb.Font.color    := clGrayText;
         false_lb.Font.color := clRed;
       end;
     else
     begin
         ok_lb.Font.color    := clGrayText;
         false_lb.Font.color := clGrayText;
     end;

  end;

end;

procedure TMainForm.SetConLabel(Server, Base: string);
begin
  Label3.Caption := Server;
  Label2.Caption := Base;
end;

function TMainForm.CheckConnect:boolean;
begin;
  ADOConnection1.Close;
  ADOConnection1.Open;
  Log('Проверка подключения к БД');
  Result :=  ADOConnection1.Connected;
end;

procedure TMainForm.Log(Mes:string);
begin
  Log_memo.Lines.append(Mes);
end;

procedure TMainForm.ShowFileList;
begin
  if DirectoryExists(Patch_Ed.Text) then
    FileListBox.Directory := Patch_Ed.Text;
end;

procedure TMainForm.ADOConnection1AfterConnect(Sender: TObject);
var str :TStringList;
begin

  str := TStringList.Create;
  try

    str.Delimiter := ';';
    str.DelimitedText := ADOConnection1.ConnectionString;
    SetConLabel(str.Values['Source'], str.Values['Catalog']);
  finally
    str.Free;
  end;
end;

procedure TMainForm.ADOConnection1BeforeDisconnect(Sender: TObject);
begin
  SetConLabel('','');
end;

procedure TMainForm.BitBtn1Click(Sender: TObject);
begin
  ShowFileList;
end;

procedure TMainForm.FileListBoxClick(Sender: TObject);
var s: string;
begin
  Screen.Cursor := crHourGlass;
  ReadTabFile(FileListBox.Items[FileListBox.ItemIndex], ';', StringGrid1);

  s := FileListBox.Items[FileListBox.ItemIndex];
  s := Copy(s, 1, Length(s)-4);
  tablename_ed.text := s;

  Screen.Cursor := crDefault;
  setFlag(0);
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  SaveSettings;
end;

procedure TMainForm.FormShow(Sender: TObject);
begin
  Patch_ed.Text := ExtractFilePath(ParamStr(0));
  LoadSettings;
  ShowFileList;
  CheckConnect;
  setFlag(0);
end;

procedure TMainForm.Load_All_BtnClick(Sender: TObject);
var i: integer;
var s: string;
begin

  if not CheckConnect then
  begin
    Log('Не удалось подключиться к БД');
    Exit;
  end;

  Screen.Cursor := crHourGlass;
  StringGrid1.Enabled := false;
  for i := 0 to FileListBox.Items.Count - 1 do
  begin
    s := FileListBox.Items[i];
    s := Copy(s, 1, Length(s)-4);
    tablename_ed.text := s;
    ReadTabFile(FileListBox.Items[i], ';', StringGrid1);
    Log('Загрузка файла: '+FileListBox.Items[i]);
    SaveToBase(StringGrid1, tablename_ed.text, spid_ed.Text);
    sleep(1000);
  end;
  StringGrid1.Enabled := true;
  Screen.Cursor := crDefault;
end;

procedure TMainForm.Load_Cur_BtnClick(Sender: TObject);
begin
  Log('Загрузка файла: '+FileListBox.Items[FileListBox.ItemIndex]);
  StringGrid1.Enabled := false;
  SaveToBase(StringGrid1, tablename_ed.text, spid_ed.Text);
  StringGrid1.Enabled := true;
end;

procedure TMainForm.ReadTabFile(FN: TFileName; FieldSeparator: Char; SG: TStringGrid);
var 
  i: Integer; 
  S: string;
  T: string;
  Colonne, ligne: Integer; 
  Les_Strings: TStringList;
  CountCols: Integer;
  CountLines: Integer; 
  TabPos: Integer; 
  StartPos: Integer; 
  InitialCol: Integer; 
begin
  for i:=0 to SG.RowCount-1 do
    SG.Rows[i].Clear;

  SG.ColCount := 1;
  SG.RowCount := 1;

  Les_Strings := TStringList.Create;
  CountCols :=0;
  try
    // Load the file, Datei laden 
    Les_Strings.LoadFromFile(FN);

    // Get the number of rows, Anzahl der Zeilen ermitteln 
    CountLines := Les_Strings.Count + SG.FixedRows;

    // Get the number of columns, Anzahl der Spalten ermitteln
    T := Les_Strings[0];
    for i := 0 to Length(T) - 1 do Inc(CountCols,
    Ord(IsDelimiter(FieldSeparator, T, i)));
    Inc(CountCols, 1 + SG.FixedCols);

    // Adjust Grid dimensions, Anpassung der Grid-Gro?e 
    if CountLines > SG.RowCount then SG.RowCount := CountLines; 
    if CountCols > SG.ColCount then SG.ColCount := CountCols; 

    // Initialisierung 
    InitialCol := SG.FixedCols - 1;
    Ligne := SG.FixedRows - 1; 

    // Iterate through all rows of the table 
    // Schleife durch allen Zeilen der Tabelle 
    for i := 0 to Les_Strings.Count - 1 do
    begin
      Colonne := InitialCol; 
      Inc(Ligne); 
      StartPos := 1;
      S := Les_Strings[i]+';'; 
      TabPos := Pos(FieldSeparator, S); 
      repeat 
        Inc(Colonne); 
        SG.Cells[Colonne, Ligne] := Copy(S, StartPos, TabPos - 1);
        S := Copy(S, TabPos + 1, 999); 
        TabPos := Pos(FieldSeparator, S); 
      until TabPos = 0; 
    end; 
  finally
    Les_Strings.Free; 
  end;
end;

function TMainForm.SaveToBase(SG: TStringGrid; TableName: string; spid: string): boolean;
var r: integer;
var c: integer;
var fieldstr: string;
var recstr: string;
var z,s,ss: string;

var
DATA_TYPE, NUMERIC_PRECISION,NUMERIC_SCALE,CHARACTER_MAXIMUM_LENGTH, column_Default:string;

var loadflag :boolean;

begin

  loadflag := false;

  if TableName <> '' then
  begin

    if not CheckConnect then
    begin
      Log('Не удалось подключиться к БД');
      Exit;
    end;

    Qdelete.Close;
    Qdelete.SQL.Clear;
    Qdelete.SQL.Add('delete from '+TableName+' where spid = '+spid);
    Qdelete.ExecSQl;

    fieldstr := SG.Cells[0, 0];
    for c := 1 to SG.ColCount - 1 do
    begin
      s := SG.Cells[c, 0];
      if s <> '' then
      begin

        QSchema.Close;
        QSchema.SQL.Clear;
        QSchema.SQL.Add('select DATA_TYPE, NUMERIC_PRECISION,NUMERIC_SCALE,CHARACTER_MAXIMUM_LENGTH, column_Default');
        QSchema.SQL.Add('from INFORMATION_SCHEMA.COLUMNS');
        QSchema.SQL.Add('where TABLE_NAME = '+chr(39)+tablename_ed.Text+chr(39));
        QSchema.SQL.Add('and COLUMN_NAME = '+chr(39)+s+chr(39));

        QSchema.Open;

        if not QSchema.IsEmpty then
        begin

          DATA_TYPE                := '';
          NUMERIC_PRECISION        := '';
          NUMERIC_SCALE            := '';
          CHARACTER_MAXIMUM_LENGTH := '';
          column_Default           := '';

          DATA_TYPE                := QSchema.FieldByName('DATA_TYPE').AsString;
          NUMERIC_PRECISION        := QSchema.FieldByName('NUMERIC_PRECISION').AsString;
          NUMERIC_SCALE            := QSchema.FieldByName('NUMERIC_SCALE').AsString;
          CHARACTER_MAXIMUM_LENGTH := QSchema.FieldByName('CHARACTER_MAXIMUM_LENGTH').AsString;
          column_Default           := QSchema.FieldByName('column_Default').AsString;

          z := ',';

          if
          //это поле identity - почистим его
          ((column_Default = '') and (DATA_TYPE = 'int'))
          //это поле не найдено в БД - почистим его
          or (DATA_TYPE = '')
          then
          begin
            //это поле identity - почистим его
            for r := 1 to SG.RowCount - 1 do
            begin
              SG.Cells[c, r] := '';
            end;
          end
          else
          begin

            fieldstr := fieldstr + ', ' + s;

            if ((DATA_TYPE = 'numeric')
            or  (DATA_TYPE = 'int')
            or  (DATA_TYPE = 'money'))
            then
             for r := 1 to SG.RowCount - 1 do
             begin
               ss := StringReplace(SG.Cells[c, r], ',','.', [rfReplaceAll]);
               ss := StringReplace(ss, ' ','', [rfReplaceAll]);
               SG.Cells[c, r] := 'convert('+DATA_TYPE+'('+NUMERIC_PRECISION+','+NUMERIC_SCALE+'),'+ss+')'+z;
             end
            else
            if DATA_TYPE = 'float' then
             for r := 1 to SG.RowCount - 1 do
             begin
               ss := StringReplace(SG.Cells[c, r], ',','.', [rfReplaceAll]);
               ss := StringReplace(ss, ' ','', [rfReplaceAll]);
               SG.Cells[c, r] := 'convert('+DATA_TYPE+'('+NUMERIC_PRECISION+'),'+ss+')'+z;
             end
            else
            if DATA_TYPE = 'tinyint'
            then
             for r := 1 to SG.RowCount - 1 do
             begin
               ss := StringReplace(SG.Cells[c, r], ',','.', [rfReplaceAll]);
               SG.Cells[c, r] := SG.Cells[c, r]+z;
             end

            else
            if ((DATA_TYPE = 'char')
            or  (DATA_TYPE = 'varchar')
            or  (DATA_TYPE = 'text'))
            then
             for r := 1 to SG.RowCount - 1 do
             begin
              SG.Cells[c, r] := 'convert('+DATA_TYPE+'('+CHARACTER_MAXIMUM_LENGTH+'),'+chr(39)+SG.Cells[c, r]+chr(39)+')'+z;
             end

            else
            if ((DATA_TYPE = 'smalldatetime')
            or  (DATA_TYPE = 'datetime'))
            then
             for r := 1 to SG.RowCount - 1 do
             begin
               SG.Cells[c, r] := 'convert('+DATA_TYPE+','+chr(39)+SG.Cells[c, r]+chr(39)+',103)'+z;
             end;
          end;

        end;
      end;
    end;

   //уберем последнюю запятую
   for r := 1 to SG.RowCount - 1 do
   begin
     ss := SG.Cells[SG.ColCount - 1, r];
     ss := Copy(ss, 1, Length(ss) - 1);
     SG.Cells[SG.ColCount - 1, r] := ss;
   end;

   //Формируем запрос на Insert для каждой строки
    if fieldstr <> '' then    

    for r := 1 to SG.RowCount - 1 do
    begin

      //count_lb.caption := inttostr(r);
      QInsert.Close;
      QInsert.SQL.Clear;
      RecStr := spid+',';
      for c := 1 to SG.ColCount - 1 do
      begin
        s := SG.Cells[c, r];
          RecStr := RecStr + s;
      end;
      QInsert.SQL.Add('Insert into '+Tablename);
      QInsert.SQL.Add('('+fieldstr+')');
      QInsert.SQL.Add('select '+recstr);
      try
        QInsert.ExecSQl;
      Except
      on E: Exception do
        begin
         Log('Ошибка: '+E.Message);
         Log('SQL   : '+QInsert.SQL.GetText);
         Clipboard.Clear;
         Clipboard.AsText := QInsert.SQL.GetText;
         Break;
        end;
      end;

    end;
    loadflag := true;
  end;

  if loadflag then
  begin
    for r:=0 to SG.RowCount-1 do
      SG.Rows[r].Clear;

    Log('Файл Загружен');
    setFlag(1);
  end
  else
  begin
    setFlag(2);
  end;

  Result := loadflag;
end;

procedure TMainForm.ServerSet_BtnClick(Sender: TObject);
begin
  ADOConnection1.Close;
  EditConnectionString(ADOConnection1);
  CheckConnect;
end;

procedure TMainForm.SaveSettings;
var Reg: TRegistry;
begin
  Reg:=TRegistry.Create;
  Reg.RootKey := HKEY_CURRENT_USER;
  Reg.OpenKey('\SOFTWARE\CSV_LOADER', true);

  Reg.WriteString('load_patch', Patch_Ed.text);
  Reg.WriteString('con_str', ADOConnection1.ConnectionString);

  Reg.Free;
end;

procedure TMainForm.LoadSettings;
var Reg: TRegistry;
begin
  Reg:=TRegistry.Create;
  Reg.RootKey := HKEY_CURRENT_USER;
  Reg.OpenKey('\SOFTWARE\CSV_LOADER', true);

  if reg.ValueExists('load_patch')
    then Patch_Ed.text := Reg.ReadString('load_patch');

  if reg.ValueExists('con_str')
    then ADOConnection1.ConnectionString := Reg.ReadString('con_str');

  Reg.free;
end;

end.
