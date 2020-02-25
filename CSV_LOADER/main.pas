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
    ADOConn: TADOConnection;
    Qdelete: TADOQuery;
    Qinsert: TADOQuery;
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
    QChkIdentityField: TADOQuery;
    RadioGroup1: TRadioGroup;
    RB_sel: TRadioButton;
    RB_all: TRadioButton;
    spid_cb: TCheckBox;

    procedure Load_BtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure FileListBoxClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ServerSet_BtnClick(Sender: TObject);
    procedure ADOConnAfterConnect(Sender: TObject);
    procedure ADOConnBeforeDisconnect(Sender: TObject);
    procedure spid_cbClick(Sender: TObject);
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

resourcestring
  DELETE_P_TABLE       = 'DELETE %s FROM %s where spid = %s';
  SELECT_TABLE_SCHEMA  = 'SELECT DATA_TYPE, NUMERIC_PRECISION, NUMERIC_SCALE, CHARACTER_MAXIMUM_LENGTH, COLUMN_DEFAULT FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = ''%s'' AND COLUMN_NAME = ''%s''';
  CHECK_IDENTITY_FIELD = 'IF EXISTS(SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = ''%s'' AND TABLE_SCHEMA = ''dbo'' AND COLUMNPROPERTY(OBJECT_ID(''%s''), COLUMN_NAME, ''ISIDENTITY'') = 1) SET IDENTITY_INSERT %s ON';
  INSERT_TO_TABLE      = 'INSERT INTO %s (%s) VALUES(%s)';

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
  ADOConn.Close;
  ADOConn.Open;

  Log('Проверка подключения к БД');

  Result :=  ADOConn.Connected;

  if Result then
    Log('Соединение с БД установлено')
  else
    Log('Нет соединение с БД');
end;

procedure TMainForm.Log(Mes:string);
begin
  Log_memo.Lines.append(Mes);
end;

procedure TMainForm.ShowFileList;
begin

  if DirectoryExists(Patch_Ed.Text) then
  begin
    FileListBox.Clear;
    FileListBox.Directory := Patch_Ed.Text;
    FileListBox.Update;
  end;

end;

procedure TMainForm.spid_cbClick(Sender: TObject);
begin
  spid_ed.Enabled := spid_cb.Checked;
end;

procedure TMainForm.ADOConnAfterConnect(Sender: TObject);
var str :TStringList;
begin

  str := TStringList.Create;
  try

    str.Delimiter := ';';
    str.DelimitedText := ADOConn.ConnectionString;
    SetConLabel(str.Values['Source'], str.Values['Catalog']);
  finally
    str.Free;
  end;
end;

procedure TMainForm.ADOConnBeforeDisconnect(Sender: TObject);
begin
  SetConLabel('', '');
end;

procedure TMainForm.BitBtn1Click(Sender: TObject);
begin
  ShowFileList;
end;

procedure TMainForm.FileListBoxClick(Sender: TObject);
var filename:  string;
begin
  filename := FileListBox.Items[FileListBox.ItemIndex];
  
  if not FileExists(filename) then
  begin
    Log('Ошибка: файл '+filename+' не найден.');
    ShowFileList;
    Exit;
  end;
  
  Screen.Cursor := crHourGlass;

  ReadTabFile(FileListBox.Items[FileListBox.ItemIndex], ';', StringGrid1);

  tablename_ed.text := ChangeFileExt(ExtractFileName(filename), '');

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

procedure TMainForm.Load_BtnClick(Sender: TObject);
var i: integer;
    s_t:string;
begin

  if not CheckConnect then
  begin
    Log('Не удалось подключиться к БД');
    Exit;
  end;

  Screen.Cursor       := crHourGlass;
  StringGrid1.Enabled := false;

  if spid_ed.Enabled then
    s_t := spid_ed.Text
  else
    s_t := '';

  for i := 0 to FileListBox.Items.Count - 1 do
  begin
    if FileListBox.Selected[i] or RB_all.Checked then
    begin
      tablename_ed.text := ChangeFileExt(ExtractFileName(FileListBox.Items[i]),'');

      ReadTabFile(FileListBox.Items[i], ';', StringGrid1);

      Log('Загрузка файла: '+FileListBox.Items[i]);

      if SaveToBase(StringGrid1, tablename_ed.text, s_t) then
      begin
        Log('Файл загружен');
        setFlag(1);
      end
      else
      begin
        Log('При загрузке файла возникла ошибка');
        setFlag(2);
      end;

      // Чтобы интерфейс не замерзал надолго
      Application.ProcessMessages;
      sleep(1000);
    end;
  end;
  StringGrid1.Enabled := true;
  Screen.Cursor := crDefault; 
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
var r, row_count:     integer;
    c, BATCH_SIZE:    integer;
    fieldstr, recstr: string;
    s, ss:            string;
    i_bad_delimeter:  integer;
    field_delimiter:  char;
    DATA_TYPE, NUMERIC_PRECISION,NUMERIC_SCALE,CHARACTER_MAXIMUM_LENGTH, column_Default: string;
begin
  Result          := False;
  field_delimiter := ',';

  if not Assigned(SG) then
  begin
    Log('Ошибка: не передан набор данных для загрузки');
    Exit;
  end;

  if TableName = '' then
  begin
    Log('Ошибка: не определено имя таблицы');
    Exit;
  end;

  if not CheckConnect then
  begin
    Log('Не удалось подключиться к БД');
    Exit;
  end;

  if spid <> '' then
  begin
    Qdelete.Close;
    Qdelete.SQL.Clear;
    Qdelete.SQL.Add(Format(DELETE_P_TABLE, [TableName, TableName, spid]));

    try
      Qdelete.ExecSQl;
    except
    on E: Exception do
      begin
       Log('Ошибка: '+E.Message);
       Log('SQL   : '+Qdelete.SQL.GetText);
       Clipboard.Clear;
       Clipboard.AsText := Qdelete.SQL.GetText;
       Exit;
      end;
    end;
  end;

  for c := 0 to SG.ColCount - 1 do
  begin
    s := SG.Cells[c, 0];

    i_bad_delimeter := Pos(':', s);
    if i_bad_delimeter > 0 then
       s := copy(s, 0, i_bad_delimeter - 1);

    if s <> '' then
    begin
      QSchema.Close;
      QSchema.SQL.Clear;
      QSchema.SQL.Add(Format(SELECT_TABLE_SCHEMA, [tablename_ed.Text, s]));

      try
        QSchema.Open;
      Except
      on E: Exception do
        begin
          Log('Ошибка: '+E.Message);
          Log('SQL   : '+QSchema.SQL.GetText);
          Clipboard.Clear;
          Clipboard.AsText := QSchema.SQL.GetText;
          Exit;
        end;
      end;

      if QSchema.IsEmpty then
      begin
        Exit;
      end
      else
      begin
        DATA_TYPE                := QSchema.FieldByName('DATA_TYPE').AsString;
        NUMERIC_PRECISION        := QSchema.FieldByName('NUMERIC_PRECISION').AsString;
        NUMERIC_SCALE            := QSchema.FieldByName('NUMERIC_SCALE').AsString;
        CHARACTER_MAXIMUM_LENGTH := QSchema.FieldByName('CHARACTER_MAXIMUM_LENGTH').AsString;
        column_Default           := QSchema.FieldByName('column_Default').AsString;

        if c > 0 then
          fieldstr := fieldstr + field_delimiter + s
        else
          fieldstr := s;

        if (spid <> '') and (s = 'SPID') then
        begin
         for r := 1 to SG.RowCount - 1 do
         begin
           SG.Cells[c, r] := spid;
           if c < SG.ColCount - 1 then SG.Cells[c, r] := SG.Cells[c, r] + field_delimiter;
         end
        end
        else
        if DATA_TYPE = 'numeric'
        then
         for r := 1 to SG.RowCount - 1 do
         begin
           ss := StringReplace(SG.Cells[c, r], ',','.', [rfReplaceAll]);
           ss := StringReplace(ss, ' ', '', [rfReplaceAll]);
           SG.Cells[c, r] := 'convert('+DATA_TYPE+'('+NUMERIC_PRECISION+','+NUMERIC_SCALE+'),'+ss+')';
           if c < SG.ColCount - 1 then SG.Cells[c, r] := SG.Cells[c, r] + field_delimiter;
         end
        else
        if DATA_TYPE = 'money'
        then
         for r := 1 to SG.RowCount - 1 do
         begin
           ss := StringReplace(SG.Cells[c, r], ',','.', [rfReplaceAll]);
           ss := StringReplace(ss, ' ', '', [rfReplaceAll]);
           SG.Cells[c, r] := 'convert('+DATA_TYPE+','+ss+')';
           if c < SG.ColCount - 1 then SG.Cells[c, r] := SG.Cells[c, r] + field_delimiter;
         end
        else
        if DATA_TYPE = 'int'
        then
         for r := 1 to SG.RowCount - 1 do
         begin
           ss := StringReplace(SG.Cells[c, r], ',','.', [rfReplaceAll]);
           ss := StringReplace(ss, ' ', '', [rfReplaceAll]);
           SG.Cells[c, r] := 'convert('+DATA_TYPE+','+ss+')';
           if c < SG.ColCount - 1 then SG.Cells[c, r] := SG.Cells[c, r] + field_delimiter;
         end
        else
        if DATA_TYPE = 'float' then
         for r := 1 to SG.RowCount - 1 do
         begin
           ss := StringReplace(SG.Cells[c, r], ',','.', [rfReplaceAll]);
           ss := StringReplace(ss, ' ','', [rfReplaceAll]);
           SG.Cells[c, r] := 'convert('+DATA_TYPE+'('+NUMERIC_PRECISION+'),'+ss+')';
           if c < SG.ColCount - 1 then SG.Cells[c, r] := SG.Cells[c, r] + field_delimiter;
         end
        else
        if DATA_TYPE = 'tinyint'
        then
         for r := 1 to SG.RowCount - 1 do
         begin
           ss := StringReplace(SG.Cells[c, r], ',','.', [rfReplaceAll]);
           SG.Cells[c, r] := SG.Cells[c, r];
           if c < SG.ColCount - 1 then SG.Cells[c, r] := SG.Cells[c, r] + field_delimiter;
         end
        else
        if ((DATA_TYPE = 'char')
        or  (DATA_TYPE = 'varchar')
        or  (DATA_TYPE = 'text'))
        then
         for r := 1 to SG.RowCount - 1 do
         begin
          SG.Cells[c, r] := 'convert('+DATA_TYPE+'('+CHARACTER_MAXIMUM_LENGTH+'),'+chr(39)+SG.Cells[c, r]+chr(39)+')';
          if c < SG.ColCount - 1 then SG.Cells[c, r] := SG.Cells[c, r] + field_delimiter;
         end
        else
        if ((DATA_TYPE = 'smalldatetime')
        or  (DATA_TYPE = 'datetime'))
        then
         for r := 1 to SG.RowCount - 1 do
         begin
           SG.Cells[c, r] := 'convert('+DATA_TYPE+','+chr(39)+SG.Cells[c, r]+chr(39)+',103)';
           if c < SG.ColCount - 1 then SG.Cells[c, r] := SG.Cells[c, r] + field_delimiter;
         end;
      end;
    end;
    Application.ProcessMessages;
  end;

  //Если в таблице есть поле identity - разрешим заполнять его
  QChkIdentityField.Close;
  QChkIdentityField.SQL.Clear;
  QChkIdentityField.SQL.Add(Format(CHECK_IDENTITY_FIELD, [TableName, TableName, TableName]));

  try
    QChkIdentityField.ExecSQL;
  except
  on E: Exception do
    begin
      Log('Ошибка: '+E.Message);
      Log('SQL   : '+QChkIdentityField.SQL.GetText);
      Clipboard.Clear;
      Clipboard.AsText := QChkIdentityField.SQL.GetText;
      Exit;
    end;
  end;

  if fieldstr = '' then
  begin
    Log('Ошибка при формировании заголовка.');
    Exit;
  end;

  // Формируем запрос на Insert для каждой строки
  // Размер блока для единовременной вставки, чтобы не делать очень много единичных запросов
  BATCH_SIZE := 50;
  row_count  := 0;

  for r := 1 to SG.RowCount - 1 do
  begin
    inc(row_count);

    if row_count = 1 then
    begin
      QInsert.Close;
      QInsert.SQL.Clear;
    end;

    // Формируем строку данных для вставки
    RecStr := '';
    for c := 0 to SG.ColCount - 1 do
    begin
      RecStr := RecStr + SG.Cells[c, r];
    end;

    QInsert.SQL.Add(Format(INSERT_TO_TABLE, [TableName, fieldstr, RecStr]));

    if (row_count = BATCH_SIZE) or (r = SG.RowCount - 1) then
    begin
      try
        // Выполняем вставку строк в таблицу БД
        QInsert.ExecSQL;

        // Выполнили пакет, сбрасываем счетчик строк
        row_count := 0;
      except
        on E: Exception do
        begin
         Log('Ошибка: '+E.Message);
         Log('SQL   : '+QInsert.SQL.GetText);
         Clipboard.Clear;
         Clipboard.AsText := QInsert.SQL.GetText;    
         Exit;
        end;
      end;
    end;

    Application.ProcessMessages;
  end;

  Result := True;
end;

procedure TMainForm.ServerSet_BtnClick(Sender: TObject);
begin
  ADOConn.Close;
  EditConnectionString(ADOConn);
  CheckConnect;
end;

procedure TMainForm.SaveSettings;
var Reg: TRegistry;
begin
  Reg:=TRegistry.Create;
  Reg.RootKey := HKEY_CURRENT_USER;
  Reg.OpenKey('\SOFTWARE\CSV_LOADER', true);

  Reg.WriteString('load_patch', Patch_Ed.text);
  Reg.WriteString('con_str', ADOConn.ConnectionString);

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
    then ADOConn.ConnectionString := Reg.ReadString('con_str');

  Reg.free;
end;

end.
