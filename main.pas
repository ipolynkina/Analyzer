unit main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, ExcelAnalyzer, Gauges, ImgList, Menus;

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
    
    g_progress_analyze: TGauge;
    
    il_pictures: TImageList;
    img_smile: TImage;
    pm_main_form: TPopupMenu;
    N1: TMenuItem;

    procedure btn_add_filesClick(Sender: TObject);
    procedure btn_run_analyzeClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure N1Click(Sender: TObject);

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
  g_progress_analyze.Progress := 0;

  analyzer := Analyzer_for_TS_files.Create();
  for index := 0 to dlg_add_files.Files.Count - 1 do begin
    analyzer.analyze_file(dlg_add_files.Files[index]);
    g_progress_analyze.Progress := 100 * (index + 1) div dlg_add_files.Files.Count;
  end;
  mmo_error_text.Text := analyzer.get_error_text();

  img_smile.Picture.Bitmap := nil;
  if analyzer.check_was_successful() then FormMain.il_pictures.GetBitmap(1, img_smile.Picture.Bitmap)
  else FormMain.il_pictures.GetBitmap(2, img_smile.Picture.Bitmap);

  analyzer.Destroy();
  btn_add_files.Enabled := True;
end;

procedure TFormMain.FormCreate(Sender: TObject);
begin
  FormMain.il_pictures.GetBitmap(0, img_smile.Picture.Bitmap);
  mmo_error_text.Text := 'Hello!';
  btn_run_analyze.Enabled := False;
end;

procedure TFormMain.FormDestroy(Sender: TObject);
begin
  dlg_add_files.Free();
end;

procedure TFormMain.N1Click(Sender: TObject);
begin
  Close();
end;

end.
