library ItlandWProj1;

uses
  ComServ,
  ItlandWProj1_TLB in 'ItlandWProj1_TLB.pas',
  ItlandWImpl1 in 'ItlandWImpl1.pas' {ItlandW: TActiveForm} {ItlandW: CoClass};

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
