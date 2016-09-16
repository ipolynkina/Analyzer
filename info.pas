unit info;

interface

uses Classes, Controls, Forms, StdCtrls, ExtCtrls;

type
  TFormInfo = class(TForm)
    img_info: TImage;
    lbl_description: TLabel;
    lbl_version: TLabel;
    lbl_release: TLabel;
    lbl_contact: TLabel;
  end;

var
  FormInfo: TFormInfo;

implementation

{$R *.dfm}

end.
