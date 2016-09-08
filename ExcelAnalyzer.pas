unit ExcelAnalyzer;

interface

uses ComObj;

type
  Analyzer_for_TS_files = class

    private
      excel : Variant;
      last_row, last_column : Integer;
      filename : String;

    public
      constructor Create();
      procedure analyze_file(filename : String);
      destructor Destroy();

  end;

implementation

// ***************************** public ************************************* //

constructor Analyzer_for_TS_files.Create();
begin
  inherited;
  excel := CreateOleObject('Excel.Application');
end;

// facade
procedure Analyzer_for_TS_files.analyze_file(filename : String);
begin
  // TODO
end;

destructor Analyzer_for_TS_files.Destroy();
begin
  excel.Application.Quit;
  inherited;
end;

// ***************************** private ************************************ //

end.
