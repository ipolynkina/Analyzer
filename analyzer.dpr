program analyzer;

uses
  Forms,
  main in 'main.pas' {FormMain},
  ExcelAnalyzer in 'ExcelAnalyzer.pas',
  LogError in 'LogError.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFormMain, FormMain);
  Application.Run;
end.
