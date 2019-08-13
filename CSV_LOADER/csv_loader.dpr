program csv_loader;

uses
  Forms,
  main in 'main.pas' {MainForm};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'Загрузка csv файлов';
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
