program WebServiceEmail;

uses
  Vcl.Forms,
  uSendEmail in 'uSendEmail.pas' {fmSendEmail};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfmSendEmail, fmSendEmail);
  Application.Run;
end.
