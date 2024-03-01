unit WorkProc;

interface

uses Windows,DB,DBTables,Forms,Controls,SysUtils,Classes,IniFiles,Dialogs,
     Registry,ADODB;

const
    Hour = 3600000/MSecsPerDay;
    Minute = 60000/MSecsPerDay;
    Second = 1000/MSecsPerDay;

  
type

        { Для архивации  БД  }
 TRecArhiv = record
                      NameBD      : string;
                     // ODBCAlias   : string;
                      pathBackup  : string;
                      TabName     : string;       // имя таблицы формуляров
                      FldName     : string;      // поле даты по которому очищают TabName
                      DayFArh     : TDate;       // дата когда архивировать
                      InterArh    : word;        // через сколько дней
                      TimeArh     : TTime;       // во сколько
                      DateArh     : TDateTime;   // с какого числа архив.
                end;




    var
         Inifile          : TIniFile;
         TableName        : string;
         Arh              : TRecArhiv;
         ClearTab         : boolean;
         FirstStart       : boolean;
         fmClose          : boolean; 

   {=*Создание резервной копии БД*=}
 procedure BackUpBD(Qu :TADOQuery;Arh :TRecArhiv);
   {= Удаление записей из таблицы TableName  =}
 function  DeleteRecordFromTable( qTemp  : TADOQuery; const FieldTime : String; const TableName : String):boolean;
    {= Создание соединения с БД =}
 function   GetConnectionToBd(Ado : TADOConnection;Arh : TRecArhiv): boolean;
      {Запись в INI}
 procedure SaveParamIniFile(Arh : TRecArhiv;WriteALL: boolean);
      {Чтение из INI}
 procedure ReadFromIniFile(var Arh : TRecArhiv; BD : bool);

implementation




{Считываем параметры из INI}
procedure ReadFromIniFile(var Arh : TRecArhiv; BD : bool);
begin
 DateSeparator   := '.';
 ShortDateFormat := 'dd/mm/yyyy';
 TimeSeparator   := ':';

  Inifile:=TIniFile.Create(ChangeFileExt(Application.ExeName,'.INI'));
 try
   ClearTab       := Inifile.ReadBool('BDClear','ClearTable',True);
   Arh.pathBackup := Inifile.ReadString('BDCopyPath','pathBakup','D:\MSSQL\BACKUP\');
   ARH.DayFArh     :=StrToDate(Inifile.ReadString('BDTime','Dayf', DateTostr(date)));
   Arh.DateArh    := StrToDate(Inifile.ReadString('BDTime','DayNext', DateTostr(date)));
   ARH.TimeArh    := StrToTime(Inifile.ReadString('BDTime','TimeArh',TimeToStR(now)));

   Arh.InterArh   := Inifile.ReadInteger('BDTime','InterArh',7);
  if not BD then
    begin
     Arh.NameBD     := Inifile.ReadString('BD','NameBD','Aist');
     Arh.TabName    := Inifile.ReadString('BD','TableName','ArhivIRI');
     Arh.FldName    := Inifile.ReadString('BD','FldName','SysTime');
    end;
 finally
    Inifile.Free;
 end;
 
end;

{Запись в INI}
procedure SaveParamIniFile(Arh : TRecArhiv;WriteALL: boolean); // Для записи в ini
var Inifile : Tinifile;
begin
 DateSeparator   := '.';
 ShortDateFormat := 'dd/mm/yyyy';
 TimeSeparator   := ':';
  ShortTimeFormat :='hh:nn';
Inifile:=TIniFile.Create(ChangeFileExt(Application.ExeName,'.INI'));
  try
   if WriteALL then  begin
  
    Inifile.WriteString('BD','NameBD', Arh.NameBD);
    Inifile.WriteString('BD','TableName',Arh.TabName);
    Inifile.WriteString('BD','Fldname',Arh.FldName);

    Inifile.WriteString('BDCopyPath','pathBakup',Arh.pathBackup);

    Inifile.WriteString('BDTime','Dayf',DateToStr(ARH.DayFArh));
    Inifile.WriteString('BDTime','TimeArh',TimeTostr(ARH.TimeArh));
    Inifile.WriteInteger('BDTime','InterArh',Arh.InterArh);

    Inifile.WriteBool('BD','ClearTable',ClearTab);
   end;
    Inifile.WriteString('BDTime','DayNext',DateToStr(Arh.DateArh));

   finally
       Inifile.Free;
  end;
end;

function GetConnectionToBd(Ado : TADOConnection;Arh : TRecArhiv): boolean;
begin
Result := False;

 if arh.NameBD<>'' then
   Ado.DefaultDatabase :=Arh.NameBD;
   try
    Ado.Connected := True;
     Result       := True;
   except
    on E : Exception do
     if StrPos(PChar(LowerCase(E.Message)),
             PChar(LowerCase('[Microsoft][Диспетчер драйверов ODBC]'+
           ' Источник данных не найден и не указан драйвер, используемый по умолчанию')))<>nil then
             MessageDlg('Ошибка при подключении к БД '+Arh.NameBD +' !' + #13#10#13+
              'Возможно такой БД не существует', mtError, [mbOk], 0);

   end;
end;




function DeleteRecordFromTable( qTemp: TADOQuery; const FieldTime: String; const TableName: String):boolean;
begin
Result := True;
/// Удаление записей
 DateSeparator :='-';
    qTemp.Close;
    qTemp.Sql.Clear;
    qTemp.SQL.Text := 'use ' + arh.NameBD;
    qTemp.SQL.Text := ' set datefirst 1 set dateformat dmy '+
     ' delete from ' + TableName + ' where ' + FieldTime + '< '''+ DateToStr(Now-1)+' 00:00:00''';
    try
      qTemp.ExecSQL;
    except
      on E : Exception do
      begin
       Result := False;
       If StrPos(PChar(LowerCase(E.Message)),PChar(LowerCase('DELETE statement conflicted with COLUMN REFERENCE constraint')))<>nil then
            MessageDlg('Ошибка при удалении записи! Удалите привязку к записи в таблице Формуляров. '+#13#10+#13#10
       + E.Message, mtError, [mbOk], 0)
       else  MessageDlg('Ошибка при удалении записи! '+#13#10+#13#10 + E.Message, mtError, [mbOk], 0);
      end;
    end;
end;


{=*Создание резервной копии БД*=}
procedure BackUpBD(Qu : TADOQuery;Arh  : TRecArhiv);
begin
 DateSeparator := '_';
 With qu do
  begin
   Close;
   SQL.Clear;
   SQL.text :='BACKUP DATABASE ' +Arh.NameBD +
   ' To DISK=N'''+ Arh.pathBackup + Arh.NameBD+DateToStr(Now)+
     ''' WITH INIT, NAME =N'''+ Arh.NameBD  + ' backup'',NOSKIP ,NOFORMAT';
    Try
     ExecSQL;
    except on  E:EDatabaseError do
     begin
      If Pos(LowerCase('status = 112'),LowerCase(E.Message))>0 then
       Application.MessageBox('На диске нет места для создания копии !',
        'Ошибка', MB_OK + MB_ICONError)
      else
       Application.MessageBox('Резервная копия БД не сформирована. Повторите !',
        'Ошибка', MB_OK + MB_ICONError) ;
       Exit;
     end;
    end;
 end;

end;




end.
