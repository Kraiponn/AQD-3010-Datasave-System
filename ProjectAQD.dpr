program ProjectAQD;

uses
  Forms,
  OK_Dat in 'OK_Dat.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
