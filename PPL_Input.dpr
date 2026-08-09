program PPL_Input;
//library Sistsim;

uses
  Forms,
  IWMain,
  UntEntrada in 'UntEntrada.pas' {IWForm1: TIWAppForm},
  ServerController in 'ServerController.pas' {IWServerController: TIWServerControllerBase},
  UserSessionUnit in 'UserSessionUnit.pas' {IWUserSession: TIWUserSessionBase},
  lpsolve51 in 'lpsolve51.pas',
  uAPI_SolveLP in 'uAPI_SolveLP.pas';

{$R *.res}

{Início ISAPI}
//exports GetExtensionVersion, HttpExtensionProc, TerminateExtension;
{Fim ISAPI}

begin
  Application.Initialize;
  Application.CreateForm(TformIWMain, formIWMain);
  Application.Run;
  //IwRun; {ISAPI}
end.
