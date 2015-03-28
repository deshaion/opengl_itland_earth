library ITLandProj1;

uses
  ComServ,
  ITLandProj1_TLB in 'ITLandProj1_TLB.pas',
  ITLandImpl1 in 'ITLandImpl1.pas' {ITLand: TActiveForm} {ITLand: CoClass};

{$E ocx}

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}

{$R *.RES}

begin
end.
