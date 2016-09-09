unit LogError;

interface

uses SysUtils, Dialogs;

type
  Logger = class

  private
    error_text: string;

  public
    constructor Create();
    procedure record_error(filename: string; id_column: Integer);
  end;

implementation

constructor Logger.Create();
begin
  error_text := '';
end;

procedure Logger.record_error(filename: string; id_column: Integer);
var
  curr_error: string;
begin
  ;
  case id_column of
    1: curr_error := filename + ': incorrect data (see column ' + IntToStr(id_column) + ')' + #13#10;
  end;
  error_text := error_text + curr_error;
end;

end.

