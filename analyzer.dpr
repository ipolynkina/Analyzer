program analyzer;

uses
  Forms,
  main in 'main.pas' {FormMain},
  ExcelAnalyzer in 'ExcelAnalyzer.pas',
  LogError in 'LogError.pas',
  info in 'info.pas' {FormInfo};

// info in 'info.pas' {FormInfo};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFormMain, FormMain);
  Application.CreateForm(TFormInfo, FormInfo);
  Application.Run;
end.
