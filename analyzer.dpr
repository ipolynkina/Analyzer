program analyzer;

uses
  Forms,
  main in 'main.pas' {FormMain},
  ExcelAnalyzer in 'ExcelAnalyzer.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFormMain, FormMain);
  Application.Run;
end.
