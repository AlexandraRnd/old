program ArhivBD;

uses
  Forms,
  Windows,
  Dialogs,
  SysUtils,
  registry,Messages,
  main in '..\Units\main.pas' {FmSaveBD},
  WorkProc in '..\Units\WorkProc.pas';

  const
     AtStr='ArhivBD';

{$R *.res}

function CheckThis : boolean;
var
Atom: THandle;
begin
Atom:= GlobalFindAtom(AtStr);
Result:= Atom <> 0;
if not result then GlobalAddAtom(AtStr);
end;

begin
if not CheckThis  then begin // Запуск программмы
 Application.Initialize;
  if ReadFromRegistry('ArhivBD') then
   begin
    FirstStart:=False;
    Application.ShowMainForm :=False;
   end;
   Application.CreateForm(TFmSaveBD, FmSaveBD);
   Application.Run;
   GlobalDeleteAtom(GlobalFindAtom(AtStr)); // !!!
end
else
 MessageBox(0,'Нельзя запустить две копии','',0);

end.


