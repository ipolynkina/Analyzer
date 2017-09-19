unit ExcelAnalyzer;

interface

uses ComObj, LogError, SysUtils, DateUtils, RegExpr;

const FIRST_LINE_OF_DATA = 2;

type file_structure = (id_date = 1, id_koeff, id_parameter_1, id_parameter_2, id_type, id_nomenclature,
                       id_min_value, id_max_value, id_default_value, id_activation, id_number_shop,
                       id_subsidiary, id_business_line);

type
  TSAnalyzer = class

    private
      filename : String;
      excel : Variant;
      last_row : Integer;
      log_error : Logger;

      procedure initialization_file(filename : String);
      function file_structure_is_correct() : Boolean;
      procedure analyze_date();
      procedure analyze_koeff();
      procedure analyze_parameter_1();
      procedure analyze_parameter_2();
      procedure analyze_type();
      procedure analyze_nomenclature();
      procedure analyze_values();
      procedure analyze_activation();
      procedure analyze_numbering_shops();
      procedure analyze_subsidiary();
      procedure analyze_business_line();
      procedure smart_check(filename : String);

    public
      constructor Create();
      procedure analyze_file(filename : String);
      function get_error_text() : String;
      function check_is_successful() : Boolean;
      destructor Destroy(); override;

  end;

implementation

// ***************************** public ************************************* //

constructor TSAnalyzer.Create();
begin
  excel := CreateOleObject('Excel.Application');
  log_error := Logger.Create();
end;

// facade
procedure TSAnalyzer.analyze_file(filename : String);
begin
  initialization_file(filename);
  if file_structure_is_correct() then begin
    analyze_date();
    analyze_koeff();
    analyze_parameter_1();
    analyze_parameter_2();
    analyze_type();
    analyze_nomenclature();
    analyze_values();
    analyze_activation();
    analyze_numbering_shops();
    analyze_subsidiary();
    analyze_business_line();
    smart_check(filename);
  end;
end;

function TSAnalyzer.get_error_text() : String;
begin
  Result := log_error.get_error_text();
end;

function TSAnalyzer.check_is_successful() : Boolean;
begin
  Result := log_error.check_is_successful();
end;

destructor TSAnalyzer.Destroy();
begin
  excel.Application.Quit;
  log_error.Free();
end;

// ***************************** private ************************************ //

procedure TSAnalyzer.initialization_file(filename : String);
begin
  excel.Workbooks.Open(filename, 0, True);
  Self.filename := ExtractFileName(filename);
  last_row := excel.ActiveSheet.UsedRange.Rows.Count;
end;

function TSAnalyzer.file_structure_is_correct() : Boolean;
const MAX_INDEX_COLUMN = 13;
var last_column : Integer;
begin
  last_column := excel.ActiveSheet.UsedRange.Columns.Count;
  if last_column <= MAX_INDEX_COLUMN then begin
    Result := True;
  end
  else begin
    Result := False;
    log_error.record_error(filename, last_column);
  end;
end;
        
procedure TSAnalyzer.analyze_date();
const MAX_DAY_IN_MONTH = 31;
var
  index_row: Integer;
  user_date : TDateTime;
begin
  for index_row := FIRST_LINE_OF_DATA to last_row do begin
    if not (TryStrToDate(String(excel.Cells[index_row, ord(id_date)]), user_date)) or
           (Abs(DaysBetween(user_date, Now())) > MAX_DAY_IN_MONTH) then begin
      log_error.record_error(filename, ord(id_date));
      Break;
    end;
  end;
end;

procedure TSAnalyzer.analyze_koeff();
var
  index_row : Integer;
  user_input : String;
begin
  if (last_row - 1 > FIRST_LINE_OF_DATA) then begin
    user_input := String(excel.Cells[FIRST_LINE_OF_DATA, ord(id_koeff)]);
    for index_row := FIRST_LINE_OF_DATA + 1 to last_row do begin
      if (String(excel.Cells[index_row, ord(id_koeff)]) <> user_input) then begin
        log_error.record_error(filename, ord(id_koeff));
        Break;
      end;
    end
  end;
end;

procedure TSAnalyzer.analyze_parameter_1();
var
  index_row: Integer;
  user_input : String;
  reg_expr : TRegExpr;
begin
  reg_expr := TRegExpr.Create();
  
  for index_row := FIRST_LINE_OF_DATA to last_row do begin
    user_input := String(excel.Cells[index_row, ord(id_parameter_1)]);
    if (user_input = '') then Continue;
    reg_expr.InputString := user_input;
    reg_expr.Expression := '[0-9/]';
    if not (reg_expr.Exec()) then begin
      log_error.record_error(filename, ord(id_parameter_1));
      Break;
    end;
  end;
end;
   
procedure TSAnalyzer.analyze_parameter_2();
var
  index_row: Integer;
  user_input : String;
  reg_expr : TRegExpr;
begin
  reg_expr := TRegExpr.Create();

  for index_row := FIRST_LINE_OF_DATA to last_row do begin
    user_input := String(excel.Cells[index_row, ord(id_parameter_2)]);
    if (user_input = '') then Continue;
    reg_expr.InputString := user_input;
    reg_expr.Expression := '[0-9/]';
    if not (reg_expr.Exec()) then begin
      log_error.record_error(filename, ord(id_parameter_2));
      Break;
    end;
  end;

  reg_expr.Free();
end;

procedure TSAnalyzer.analyze_type();
const MIN_ID_TYPE = 0;
const MAX_ID_TYPE = 5;
var
  index_row, user_value : Integer;
begin
  for index_row := FIRST_LINE_OF_DATA to last_row do begin
    if not (TryStrToInt(excel.Cells[index_row, ord(id_type)], user_value) or
        (user_value < MIN_ID_TYPE) or (user_value > MAX_ID_TYPE)) then begin
      log_error.record_error(filename, ord(id_type));
      Break;
    end;
  end;
end;

procedure TSAnalyzer.analyze_nomenclature();
var
  index_row : Integer;
  user_input : String;
begin
  for index_row := FIRST_LINE_OF_DATA to last_row do begin
    user_input := String(excel.Cells[index_row, ord(id_nomenclature)]);
    if (user_input = '') then Continue;
    if (user_input = 'BF_INST_AGENT_NOTCAL') or (user_input = 'BF_ACC_ACTION_BAG') or
        (user_input = 'BF_INST_ACTION_NN') or (user_input = 'ACC_ON_CASH_ZONE  ') then begin
      log_error.record_error(filename, ord(id_nomenclature));
      Break;
    end;
  end;
end;

procedure TSAnalyzer.analyze_values();
var
  index_row : Integer;
  min_value_of_user, max_value_of_user, default_value_of_user : Double;
begin
  for index_row := FIRST_LINE_OF_DATA to last_row do begin
    if not (TryStrToFloat(excel.Cells[index_row, ord(id_min_value)], min_value_of_user)) then begin
      log_error.record_error(filename, ord(id_min_value));
      Break;
    end;

    if not (TryStrToFloat(excel.Cells[index_row, ord(id_max_value)], max_value_of_user)) then begin
      log_error.record_error(filename, ord(id_max_value));
      Break;
    end;

    if not (TryStrToFloat(excel.Cells[index_row, ord(id_default_value)], default_value_of_user)) then begin
      log_error.record_error(filename, ord(id_default_value));
      Break;
    end;

    if (min_value_of_user <> max_value_of_user) or (max_value_of_user <> default_value_of_user) then begin
      if (min_value_of_user = max_value_of_user) then log_error.record_error(filename, ord(id_default_value))
      else if (max_value_of_user = default_value_of_user) then log_error.record_error(filename, ord(id_min_value))
      else log_error.record_error(filename, ord(id_max_value));
      Break;
    end;
  end;
end;

procedure TSAnalyzer.analyze_activation();
const ACTIVATION_SIGN = '1';
var
  index_row : Integer;
begin
  for index_row := FIRST_LINE_OF_DATA to last_row do begin
    if (String(excel.Cells[index_row, ord(id_activation)]) <> ACTIVATION_SIGN) then begin
      log_error.record_error(filename, ord(id_activation));
      Break;
    end;
  end;
end;

procedure TSAnalyzer.analyze_numbering_shops();
var
  index_row : Integer;
  user_input : String;
  reg_expr_one_shop,  reg_expr_multi_shops: TRegExpr;
begin
  reg_expr_one_shop := TRegExpr.Create();
  reg_expr_multi_shops := TRegExpr.Create();

  for index_row := FIRST_LINE_OF_DATA to last_row do begin
    user_input := String(excel.Cells[index_row, id_number_shop]);
    if (user_input <> '') then begin
      reg_expr_one_shop.InputString := user_input;
      reg_expr_multi_shops.InputString := user_input + ',';

      reg_expr_one_shop.Expression := '^([A-B]{1}[0-9]{3})$';
      reg_expr_multi_shops.Expression := '^('#39'{1}[A-B]{1}[0-9]{3}'#39'{1}[,]{1}){1,}$';

      if not (reg_expr_one_shop.Exec()) and not (reg_expr_multi_shops.Exec()) then begin
        log_error.record_error(filename, ord(id_number_shop));
        Break;
      end;
    end;
  end;

  reg_expr_one_shop.Free();
  reg_expr_multi_shops.Free();
end;

procedure TSAnalyzer.analyze_subsidiary();
var
  index_row : Integer;
begin
  for index_row := FIRST_LINE_OF_DATA to last_row do begin
    if (String(excel.Cells[index_row, ord(id_subsidiary)]) <> '') then begin
      log_error.record_error(filename, ord(id_subsidiary));
      Break;
    end;
  end;
end;

procedure TSAnalyzer.analyze_business_line();
var
  index_row : Integer;
  user_input : String;
begin
  for index_row := FIRST_LINE_OF_DATA to last_row do begin
    user_input := String(excel.Cells[index_row, ord(id_business_line)]);
    if (user_input <> '') and (user_input <> 'PVZ') and (user_input <> 'PZVS') then begin
      log_error.record_error(filename, ord(id_business_line));
      Break;
    end;
  end;
end;

procedure TSAnalyzer.smart_check(filename : String);
var
  kpi_name : String;
  reg_expr : TRegExpr;
begin
  kpi_name := '';
  reg_expr := TRegExpr.Create();
  reg_expr.InputString := filename;
  reg_expr.Expression := '[KPI]+(\s)+(\d{1,2})';
  if (reg_expr.Exec()) then begin
    repeat
      kpi_name := kpi_name + reg_expr.Match[0];
    until not reg_expr.ExecNext;
    kpi_name := StringReplace(kpi_name, ' ', '', [rfReplaceAll]);
    if (kpi_name <> String(excel.Cells[FIRST_LINE_OF_DATA, ord(id_koeff)])) then begin
      log_error.record_error(filename, ord(id_koeff));
    end;
  end;
end;

end.
