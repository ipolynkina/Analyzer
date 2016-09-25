unit main;

interface

uses
  SysUtils, Classes, Graphics, Controls, Forms, Dialogs, ExtCtrls,
  StdCtrls, ExcelAnalyzer, Gauges, ImgList, Menus;

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
    get_info: TMenuItem;
    decorator: TMenuItem;
    exit: TMenuItem;

    procedure FormCreate(Sender: TObject);
    procedure btn_add_filesClick(Sender: TObject);
    procedure btn_run_analyzeClick(Sender: TObject);
    procedure get_infoClick(Sender: TObject);
    procedure exitClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);

  end;

var
  FormMain: TFormMain;

implementation

uses info;

{$R *.dfm}

procedure TFormMain.FormCreate(Sender: TObject);
begin
  FormMain.il_pictures.GetBitmap(0, img_smile.Picture.Bitmap);
  mmo_error_text.Text := 'Hello!';
  Application.HintHidePause := 5000;

  dlg_add_files := TOpenDialog.Create(Self);
  dlg_add_files.Options := [ofAllowMultiSelect];
  dlg_add_files.Filter := 'Excel|*.xls';

  btn_run_analyze.Enabled := False;
end;

procedure TFormMain.btn_add_filesClick(Sender: TObject);
var
  index : Integer;
begin
  mmo_list_files.Text := '';
  mmo_error_text.Text := '';

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
  analyzer : TSAnalyzer;
begin
  btn_add_files.Enabled := False;
  btn_run_analyze.Enabled := False;
  g_progress_analyze.Progress := 0;

  analyzer := TSAnalyzer.Create();
  for index := 0 to dlg_add_files.Files.Count - 1 do begin
    try
      analyzer.analyze_file(dlg_add_files.Files[index]);
    except
      on error_report : Exception do begin
        ShowMessage('Error: ' + error_report.Message + '. Sorry');
        Close();
      end;
    end;
    g_progress_analyze.Progress := 100 * (index + 1) div dlg_add_files.Files.Count;
  end;

  mmo_error_text.Text := analyzer.get_error_text();
  img_smile.Picture.Bitmap := nil;
  if analyzer.check_is_successful() then FormMain.il_pictures.GetBitmap(1, img_smile.Picture.Bitmap)
  else FormMain.il_pictures.GetBitmap(2, img_smile.Picture.Bitmap);

  analyzer.Destroy();
  btn_add_files.Enabled := True;
end;

procedure TFormMain.get_infoClick(Sender: TObject);
begin
  FormMain.il_pictures.GetBitmap(3, FormInfo.img_info.Picture.Bitmap);
  FormInfo.ShowModal();
end;

procedure TFormMain.exitClick(Sender: TObject);
begin
  Close();
end;

procedure TFormMain.FormDestroy(Sender: TObject);
begin
  dlg_add_files.Free();
end;

end.
