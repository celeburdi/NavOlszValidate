program NavOlszValidate;

uses
  Vcl.Forms,
  mainfrm in 'mainfrm.pas' {frmMain};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
