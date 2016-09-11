unit main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, ExcelAnalyzer;

type
  TFormMain = class(TForm)

    pnl_right: TPanel;
    pnl_top_right: TPanel;
    pnl_left: TPanel;
    pnl_top_left: TPanel;
    pnl_bottom: TPanel;

    mmo_list_files: TMemo;
    mmo_error_text: TMemo;

    btn_add_files: TButton;
    btn_run_analyze: TButton;

    dlg_add_files: TOpenDialog;

    procedure btn_add_filesClick(Sender: TObject);
    procedure btn_run_analyzeClick(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormMain: TFormMain;

implementation

{$R *.dfm}

procedure TFormMain.btn_add_filesClick(Sender: TObject);
var
  index : Integer;
begin
  mmo_list_files.Text := '';
  mmo_error_text.Text := '';
  
  dlg_add_files := TOpenDialog.Create(Self);
  dlg_add_files.Options := [ofAllowMultiSelect];
  if dlg_add_files.Execute then begin
    for index := 0 to dlg_add_files.Files.Count - 1 do begin
      mmo_list_files.Text := mmo_list_files.Text + ExtractFileName(dlg_add_files.Files[index]) + #13#10;
    end;
    btn_run_analyze.Enabled := True;
  end;
end;

procedure TFormMain.btn_run_analyzeClick(Sender: TObject);
var
  index : Integer;
  analyzer : Analyzer_for_TS_files;
begin
  btn_add_files.Enabled := False;
  btn_run_analyze.Enabled := False;

  analyzer := Analyzer_for_TS_files.Create();
  for index := 0 to dlg_add_files.Files.Count - 1 do begin
    analyzer.analyze_file(dlg_add_files.Files[index]);
  end;

  mmo_error_text.Text := analyzer.get_error_text();
  analyzer.Destroy();
  
  btn_add_files.Enabled := True;
end;

end.
