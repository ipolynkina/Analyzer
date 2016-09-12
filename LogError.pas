unit LogError;

interface

uses SysUtils;

type
  Logger = class

    private
      error_text: String;
      file_has_error : Boolean;

    public
      constructor Create();
      procedure record_error(filename: String; id_column: Integer);
      function get_error_text() : String;
      function check_was_successful() : Boolean;

  end;

implementation

constructor Logger.Create();
begin
  error_text := '';
  file_has_error := False;
end;

procedure Logger.record_error(filename: String; id_column: Integer);
var
  curr_error: String;
begin;
  case id_column of
    1: curr_error := ': incorrect data ';
    2: curr_error := ': coefficients have different names ';
    3: curr_error := ': incorrect parameter 1 ';
    4: curr_error := ': incorrect parameter 2 ';
    5: curr_error := ': incorrect type ';
    6: curr_error := ': used the old name ';
    7, 8, 9: curr_error := ': different values ';
    10: curr_error := ': values are not activated ';
    11: curr_error := ': incorrect number shop ';
  end;

  error_text := error_text + filename + curr_error + '(see column ' + IntToStr(id_column) + ')' + #13#10;
  file_has_error := True;
end;

function Logger.get_error_text() : String;
begin
  if (error_text = '') then Result := 'OK'
  else Result := error_text;
end;

function Logger.check_was_successful() : Boolean;
begin
  if file_has_error then Result := False
  else Result := True;
end;

end.

