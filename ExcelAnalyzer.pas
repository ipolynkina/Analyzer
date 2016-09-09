unit ExcelAnalyzer;

interface

uses ComObj, LogError, SysUtils, DateUtils;

type
  file_structure = (id_date = 1
                    );

type
  Analyzer_for_TS_files = class

    private
      filename : String;
      excel : Variant;
      last_row, last_column : Integer;
      log_error : Logger;

      procedure initialization_file(filename : String);
      procedure analyze_date();

    public
      constructor Create();
      procedure analyze_file(filename : String);
      destructor Destroy(); override;

  end;

implementation

// ***************************** public ************************************* //

constructor Analyzer_for_TS_files.Create();
begin
  excel := CreateOleObject('Excel.Application');
  log_error := Logger.Create();
end;

// facade
procedure Analyzer_for_TS_files.analyze_file(filename : String);
begin
  // TODO
  initialization_file(filename);
  analyze_date();
end;

destructor Analyzer_for_TS_files.Destroy();
begin
  excel.Application.Quit;
  log_error.Free;
end;

// ***************************** private ************************************ //

procedure Analyzer_for_TS_files.initialization_file(filename : String);
begin
  excel.Workbooks.Open(filename, 0, True);
  Self.filename := ExtractFileName(filename);
  last_row := excel.ActiveSheet.UsedRange.Rows.Count;
  //last_column := excel.ActivateSheet.UsedRange.Columns.Count;
end;

procedure Analyzer_for_TS_files.analyze_date();
const
  MAX_DAY_IN_MONTH = 31;
var
  index_row : Integer;
  user_input : String;
  user_date, today : TDateTime;
begin
  today := Now;
  for index_row := 2 to last_row do begin
    user_input := excel.Cells[index_row, ord(id_date)];
    if not (TryStrToDate(user_input, user_date)) or
    (Abs(DayOfTheYear(today) - DayOfTheYear(StrToDateTime(user_input))) > MAX_DAY_IN_MONTH)
    then begin
      log_error.record_error(filename, Ord(id_date));
      Break;
    end;
  end;
end;

end.
