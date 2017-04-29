program pricecompare;

{$mode objfpc}{$H+}

uses {$IFDEF UNIX} {$IFDEF UseCThreads}
  cthreads, {$ENDIF} {$ENDIF}
  Classes,
  SysUtils,
  CustApp,
  fpspreadsheet,
  laz_fpspreadsheet { you can add units after this };

type

  { TPriceCompare }

  TPriceCompare = class(TCustomApplication)
  protected
    procedure DoRun; override;
  public
    constructor Create(TheOwner: TComponent); override;
    destructor Destroy; override;
    procedure WriteHelp; virtual;
    procedure appendFile(fileName: string);
  end;

  { TPriceCompare }
var
  Application: TPriceCompare;
  resFile: string;

  procedure TPriceCompare.DoRun;
  var
    ErrorMsg: string;
    i: integer;
  begin
    // quick check parameters
    ErrorMsg := CheckOptions('h', 'help');
    if ErrorMsg <> '' then
    begin
      ShowException(Exception.Create(ErrorMsg));
      Terminate;
      Exit;
    end;

    // parse parameters
    if HasOption('h', 'help') then
    begin
      WriteHelp;
      Terminate;
      Exit;
    end;

    if ParamCount = 0 then
    begin
      Writeln('Параметр по умолчанию, содержащий путь к исполняемому файлу программы:');
      Writeln(ParamStr(0));
      Writeln('Дополнительные параметры в программу не переданы.');
    end
    else
    begin
      Writeln('Параметр по умолчанию, содержащий путь к исполняемому файлу программы:');
      Writeln(ParamStr(0));
      Writeln('Дополнительные параметры, переданные в программу:');
      resFile := ExtractFileDir(ParamStr(0)) + PathDelim + 'result.xls';
      Writeln('Результирующий файл ' + resFile);
      for i := 1 to ParamCount do
      begin
        if FileExists(ParamStr(i)) and (ParamCount > 1) then
        begin
          appendFile(ParamStr(i));
        end;
      end;
    end;

    // stop program loop
    Terminate;
  end;

  constructor TPriceCompare.Create(TheOwner: TComponent);
  begin
    inherited Create(TheOwner);
    StopOnException := True;
  end;

  destructor TPriceCompare.Destroy;
  begin
    inherited Destroy;
  end;

  procedure TPriceCompare.WriteHelp;
  begin
    { add your help code here }
    writeln('Usage: ', ExeName, ' -h');
  end;

  procedure TPriceCompare.appendFile(fileName: string);
  var
    i: integer;
    InWorkbook, OutWorkbook: TsWorkbook;
    InWorksheet, OutWorksheet: TsWorksheet;
    productName, productManufact, productUnit: string;
    productCost: Single;
  begin
    Writeln('Файл-' + fileName);
    {
    OutWorkbook := TsWorkbook.Create;
    OutWorkbook.FileName := resFile;
    OutWorksheet := MyWorkbook.AddWorksheet('Compare result');
    }
    InWorkbook := TsWorkbook.Create;
    InWorkbook.ReadFromFile(fileName);
    InWorksheet := InWorkbook.GetFirstWorksheet;
    WriteLn(InWorksheet.GetLastRowIndex());
    for i := 0 to InWorksheet.GetLastRowIndex() do
    begin
      productName := InWorksheet.ReadAsText(i, 0);
      productManufact := InWorksheet.ReadAsText(i, 1);
      productUnit := InWorksheet.ReadAsText(i, 2);
      productCost := InWorksheet.ReadAsNumber(i, 3);
      WriteLn(IntToStr(i) + ' ' + productName + ' ' + productManufact + ' ' +
        productUnit + ' ' + FloatToStr(productCost));
    end;
  end;


begin
  Application := TPriceCompare.Create(nil);
  Application.Title := 'pricecompare';
  Application.Run;
  Application.Free;
end.
