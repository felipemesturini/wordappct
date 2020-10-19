program WCVT;

{$APPTYPE CONSOLE}
{$R *.res}

uses
  System.SysUtils,
  System.Win.ComObj,
  System.Variants,
  System.IOUtils,
  System.Classes;

function AbrirPlanilha(const AName: String; AHandleCol: TProc<String, String>; out ANewLine: Boolean): Boolean;
var
  lArquivo: TFileStream;
begin
  lArquivo := TFile.OpenRead(AName);
  ANewLine := False;


end;

var
  lWordApp: Variant;
  lNewDoc: Variant;

begin
  try
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
