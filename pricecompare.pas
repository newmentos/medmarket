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

  product = record
    productName, productManufact, productUnit: string;
    productCost: single;
  end;

  { TPriceCompare }

  TPriceCompare = class(TCustomApplication)
  protected
    procedure DoRun; override;
  public
    constructor Create(TheOwner: TComponent); override;
    destructor Destroy; override;
    procedure WriteHelp; virtual;
    procedure parseFile(fileName: string);
    procedure appendProduct(p: product);
  end;

  { TPriceCompare }
var
  Application: TPriceCompare;
  resFile: string;

  procedure printProduct(const p: product; const i: integer);
  begin
    WriteLn('запись номер ' + IntToStr(i) + ' ' + p.productName +
      ' ' + p.productManufact + ' ' + p.productUnit + ' ' + FloatToStr(p.productCost));
  end;

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
      if FileExists(resFile) then
        DeleteFile(resFile);
      Writeln('Результирующий файл ' + resFile);
      for i := 1 to ParamCount do
      begin
        if FileExists(ParamStr(i)) and (ParamCount > 1) then
        begin
          parseFile(ParamStr(i));
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

  procedure TPriceCompare.parseFile(fileName: string);
  var
    i: integer;
    InWorkbook: TsWorkbook;
    InWorksheet: TsWorksheet;
    p: product;
  begin
    Writeln('Читаем файл-' + fileName);
    InWorkbook := TsWorkbook.Create;
    InWorkbook.ReadFromFile(fileName);
    InWorksheet := InWorkbook.GetFirstWorksheet;
    WriteLn(InWorksheet.GetLastRowIndex());
    for i := 0 to InWorksheet.GetLastRowIndex() do
    begin
      with p do
      begin
        productName := InWorksheet.ReadAsText(i, 0);
        productManufact := InWorksheet.ReadAsText(i, 1);
        productUnit := InWorksheet.ReadAsText(i, 2);
        productCost := InWorksheet.ReadAsNumber(i, 3);
      end;
      appendProduct(p);
      printProduct(p, i);
    end;
  end;

  procedure TPriceCompare.appendProduct(p: product);
  var
    OutWorkbook: TsWorkbook;
    OutWorksheet: TsWorksheet;
    i: integer;
  begin
    OutWorkbook := TsWorkbook.Create;
    OutWorksheet := OutWorkbook.AddWorksheet('Compare result');
//    OutWorkbook.WriteToFile(resFile);
  end;

begin
  Application := TPriceCompare.Create(nil);
  Application.Title := 'pricecompare';
  Application.Run;
  Application.Free;
end.
