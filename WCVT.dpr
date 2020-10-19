program WCVT;

{$APPTYPE CONSOLE}
{$R *.res}

uses
  System.SysUtils,
  System.Win.ComObj,
  System.Variants,
  System.IOUtils,
  System.Classes;

function AbrirPlanilha(const AName: String; AHandleCol: TProc<String, String>): Boolean;
var
  lArquivo: TStreamReader;
  lLinha: string;
  lLine: Integer;
  lColunas: TArray<string>;
begin
  lArquivo := TFile.OpenText(AName);
//  ANewLine := False;
  lLinha := EmptyStr;
  lLine := 0;
  while not lArquivo.EndOfStream do
  begin
    lLinha := lArquivo.ReadLine;
    if (lLine = 0) then begin
      lColunas := lLinha.Split([';']);
    end;
    Inc(lLine);
  end;
  lArquivo.Close;
  lArquivo.Free;
end;

var
  lWordApp: Variant;
  lNewDoc: Variant;
  lPlanilha: string;
begin
  try
    lPlanilha := TPath.Combine(ExtractFilePath(ParamStr(0)), 'Planilha.csv');
    AbrirPlanilha(lPlanilha,
      procedure(AColuna, AValor: String)
      begin
      end
    );
    { Creates a Microsoft Word application. }
    lWordApp := CreateOleObject('Word.Application');
    { Creates a new Microsoft Word document. }
    lNewDoc := lWordApp.Documents.Add;

    { Inserts the text 'Hello World!' in the document. }
    lWordApp.Selection.TypeText('Hello World!');
    { Saves the document on the disk. }
    lNewDoc.SaveAs('my_new_document.doc');
    { Closes Microsoft Word. }
    lWordApp.Quit;
    { Releases the interface by assigning the Unassigned constant to the Variant variables. }
    lNewDoc := Unassigned;
    lWordApp := Unassigned;
  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;

end.
