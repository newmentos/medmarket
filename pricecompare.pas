program pricecompare;

{$mode objfpc}{$H+}

uses {$IFDEF UNIX} {$IFDEF UseCThreads}
  cthreads, {$ENDIF} {$ENDIF}
  Classes,
  SysUtils,
  CustApp,
  fpspreadsheet,
  laz_fpspreadsheet,
  Math { you can add units after this };

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
    function CompareStrings_Levenshtein(const A, B: string;
      CaseSensitive: boolean = False): integer;
    function CompareStrings_Ratcliff(const A, B: string;
      CaseSensitive: boolean = False): double;
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

  procedure printProduct(const p: product);
  begin
    Write(p.productName + ' ' + p.productManufact + ' ' + p.productUnit +
      ' ' + FloatToStr(p.productCost));
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
      OutWorksheet.WriteText(0, 0, 'Наименование');
      OutWorksheet.WriteText(0, 1, 'Производитель');
      OutWorksheet.WriteText(0, 2, 'еденица измерения');
      curRow := 0;
      curColumn := 3;
      Writeln('Результирующий файл ' + resFile);
      for i := 1 to ParamCount do
      begin
        if FileExists(ParamStr(i)) then
        begin
          parseFile(ParamStr(i)); // Send file to parsing
          curColumn += 1;
          OutWorksheet.WriteText(0, 2 + i, ParamStr(i));
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
    WriteLn('Run', ExeName, ' file1.xls file2.xls ... filen.xls');
    WriteLn('program generate file result.xls where compare cost of products');
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
      //      printProduct(p);
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
    Result := 0;
    for j := 1 to curRow do
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
        Result := j;
        Break;
      end;
    end;
  end;

  function TPriceCompare.compare(p, pTmp: product): boolean;
  var
    s1, s1Tmp, s2, s2Tmp: string;
    lev1, lev2: integer;
    ratcl1, ratcl2: double;
  begin
    s1 := Trim(UpperCase(p.productName));
    s1Tmp := Trim(UpperCase(pTmp.productName));
    s2 := Trim(UpperCase(p.productManufact));
    s2Tmp := Trim(UpperCase(p.productManufact));
    Result := False;
    lev1 := CompareStrings_Levenshtein(s1, s1Tmp, False);
    lev2 := CompareStrings_Levenshtein(s2, s2Tmp, False);
    ratcl1 := CompareStrings_Ratcliff(s1, s1Tmp, False);
    ratcl2 := CompareStrings_Ratcliff(s2, s2Tmp, False);
    if (s1 = s1Tmp) and (s2 = s2Tmp) then
    begin
      {
      WriteLn(s1 + ' ' + s1Tmp + ' Levenshtein=' + IntToStr(lev1) +
        ' Ratcliff=' + FloatToStr(ratcl1));
      WriteLn(s2 + ' ' + s2Tmp + ' Levenshtein=' + IntToStr(lev2) +
        ' Ratcliff=' + FloatToStr(ratcl2));
      WriteLn();
      Write('Сравнивали ');
      printProduct(p);
      Write(' и ');
      printProduct(pTmp);
      WriteLn();
            }
      Result := True;
    end
    else
      Result := False;
  end;


  // функции взяты отсюда http://www.delphigroups.info/2/25/422444.html  (John Leavey )
  function TPriceCompare.CompareStrings_Levenshtein(const A, B: string;
    CaseSensitive: boolean = False): integer;

    function Minimum3(x, y, z: integer): integer;
    begin
      Result := Min(x, y);
      Result := Min(Result, z);
    end;

  var
    D: array of array of integer;
    n, m, i, j, Cost: integer;
    AI, BJ: char;
    A1, B1: string;
  begin
    n := Length(A);
    m := Length(B);
    if (n = 0) then
      Result := m
    else if m = 0 then
      Result := n
    else
    begin
      if CaseSensitive then
        A1 := A
      else
        A1 := UpperCase(A);
      if CaseSensitive then
        B1 := B
      else
        B1 := UpperCase(B);
      Setlength(D, n + 1, m + 1);
      for i := 0 to n do
        D[i, 0] := i;
      for j := 0 to m do
        D[0, j] := j;
      for i := 1 to n do
      begin
        AI := A1[i];
        for j := 1 to m do
        begin
          BJ := B1[j];
          //  Cost := iff(AI = BJ, 0, 1);
          if (AI = BJ) then
            Cost := 0
          else
            Cost := 1;
          D[i, j] := Minimum3(D[i - 1][j] + 1, D[i][j - 1] + 1,
            D[i - 1][j - 1] + Cost);
        end;
      end;
      Result := D[n, m];
    end;
  end;

  function TPriceCompare.CompareStrings_Ratcliff(const A, B: string;
    CaseSensitive: boolean = False): double;
  var
    A1, B1: string;
    LenA, LenB: integer;

    function CSRSub(StartA, EndA, StartB, EndB: integer): integer;
    var
      ai, bi, i, Matches, NewStartA, NewStartB: integer;
    begin
      Result := 0;
      NewStartA := 0;
      NewStartB := 0;
      if (StartA > EndA) or (StartB > EndB) or (StartA <= 0) or (StartB <= 0) then
        Exit;
      for ai := StartA to EndA do
      begin
        for bi := StartB to EndB do
        begin
          Matches := 0;
          i := 0;
          while (ai + i <= EndA) and (bi + i <= EndB) and
            (A1[ai + i] = B1[bi + i]) do
          begin
            Inc(Matches);
            if Matches > Result then
            begin
              NewStartA := ai;
              NewStartB := bi;
              Result := Matches;
            end;
            Inc(i);
          end;
        end;
      end;
      if Result > 0 then
      begin
        Inc(Result, CSRSub(NewStartA + Result, EndA, NewStartB + Result, EndB));
        Inc(Result, CSRSub(StartA, NewStartA - 1, StartB, NewStartB - 1));
      end;
    end;

  begin
    if CaseSensitive then
      A1 := A
    else
      A1 := UpperCase(A);
    if CaseSensitive then
      B1 := B
    else
      B1 := UpperCase(B);
    LenA := Length(A1);
    LenB := Length(B1);
    if A1 = B1 then
      Result := 100
    else if (LenA = 0) or (LenB = 0) then
      Result := 0
    else
      Result := CSRSub(1, LenA, 1, LenB) * 200 / (LenA + LenB);
  end;

begin
  Application := TPriceCompare.Create(nil);
  Application.Title := 'pricecompare';
  Application.Run;
  Application.Free;
end.
