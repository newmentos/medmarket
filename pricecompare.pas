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
    function searchEquality(p: product): integer;
    function compare(p, pTmp: product): boolean;
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
  OutWorkbook: TsWorkbook;
  OutWorksheet: TsWorksheet;
  curRow, curColumn: integer;



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
      OutWorkbook := TsWorkbook.Create;
      OutWorksheet := OutWorkbook.AddWorksheet('Compare result');
      curRow := 0;
      curColumn := 3;
      Writeln('Результирующий файл ' + resFile);
      for i := 1 to ParamCount do
      begin
        if FileExists(ParamStr(i)) and (ParamCount > 1) then
        begin
          parseFile(ParamStr(i)); // Send file to parsing
          curColumn += 1;
        end;
      end;
      OutWorkbook.WriteToFile(resFile);
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
    keyRow: integer;
  begin
    keyRow := searchEquality(p);
    if (keyRow = 0) then
    begin
      curRow += 1;
      with p do
      begin
        OutWorksheet.WriteText(curRow, 0, productName);
        OutWorksheet.WriteText(curRow, 1, productManufact);
        OutWorksheet.WriteText(curRow, 2, productUnit);
        OutWorksheet.WriteCurrency(curRow, curColumn, FloattoCurr(productCost));
      end;
    end
    else
    begin
      OutWorksheet.WriteCurrency(keyRow, curColumn, FloattoCurr(p.productCost));
    end;
  end;


  function TPriceCompare.searchEquality(p: product): integer;
  var
    j: integer;
    pTmp: product;
  begin
    searchEquality := 0;
    for j := 1 to OutWorksheet.GetLastRowIndex() do
    begin
      with pTmp do
      begin
        productName := OutWorksheet.ReadAsText(j, 0);
        productManufact := OutWorksheet.ReadAsText(j, 1);
        productUnit := OutWorksheet.ReadAsText(j, 2);
        productCost := OutWorksheet.ReadAsNumber(j, 3);
      end;
      if compare(p, pTmp) then
      begin
        searchEquality := j;
        Break;
      end;
    end;
  end;

  function TPriceCompare.compare(p, pTmp: product): boolean;
  begin
    compare := True;
  end;

begin
  Application := TPriceCompare.Create(nil);
  Application.Title := 'pricecompare';
  Application.Run;
  Application.Free;
end.
