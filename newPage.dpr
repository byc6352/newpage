program newPage;

uses
  Vcl.Forms,
  uMain in 'uMain.pas' {fMain},
  uFuncs in 'uFuncs.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfMain, fMain);
  Application.Run;
end.
