unit OK_Dat;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, CPort,IniFiles, ExtCtrls, Menus, DB, ADODB, jpeg, GIFImg,
  MPlayer, Grids, DBGrids, Buttons;

type
  TForm1 = class(TForm)
    Memo1: TMemo;
    ComPort1: TComPort;
    Button3: TButton;
    Button4: TButton;
    Panel1: TPanel;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    edtBoxNo: TEdit;
    edtPos6: TEdit;
    edtPos_2: TEdit;
    edtPos13: TEdit;
    edtPos19: TEdit;
    edtPos27: TEdit;
    edtPos33: TEdit;
    EdtPos34: TEdit;
    edtPos35: TEdit;
    edtPos36: TEdit;
    EdtPos37: TEdit;
    EdtPos38: TEdit;
    edtPos42: TEdit;
    edtPos_22: TEdit;
    edtPos62: TEdit;
    edtPos132: TEdit;
    edtPos192: TEdit;
    edtPos332: TEdit;
    edtPos342: TEdit;
    edtPos272: TEdit;
    edtPos352: TEdit;
    edtPos362: TEdit;
    EdtPos372: TEdit;
    edtPos382: TEdit;
    edtPos422: TEdit;
    Panel2: TPanel;
    GroupBox2: TGroupBox;
    Label17: TLabel;
    GroupBox3: TGroupBox;
    Memo2: TMemo;
    Label18: TLabel;
    Label19: TLabel;
    lblLed: TLabel;
    Panel3: TPanel;
    btnSetCom: TButton;
    Button5: TButton;
    Label20: TLabel;
    Label21: TLabel;
    edtPosition: TEdit;
    edtHis: TEdit;
    Timer1: TTimer;
    MainMenu1: TMainMenu;
    PopupMenu1: TPopupMenu;
    File1: TMenuItem;
    SettingComPort1: TMenuItem;
    AutoScan1: TMenuItem;
    PauseAutoScan1: TMenuItem;
    N1: TMenuItem;
    Exit1: TMenuItem;
    SettingComPort2: TMenuItem;
    AboutProgram1: TMenuItem;
    AboutProgram2: TMenuItem;
    Autoscan2: TMenuItem;
    StopProgram1: TMenuItem;
    N2: TMenuItem;
    Exit2: TMenuItem;
    Button6: TButton;
    Button7: TButton;
    Button8: TButton;
    Button9: TButton;
    Button10: TButton;
    Button11: TButton;
    Button12: TButton;
    Button13: TButton;
    lblCount: TLabel;
    Image1: TImage;
    Label16: TLabel;
    Label15: TLabel;
    btnFirstJump: TButton;
    Timer2: TTimer;
    Image2: TImage;
    Edit1: TEdit;
    Label22: TLabel;
    Label23: TLabel;
    btnCount: TButton;
    btnSecondJump: TButton;
    CloseComPort1: TMenuItem;
    lblClock: TLabel;
    Timer3: TTimer;
    Timer4: TTimer;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    ADOQuery2: TADOQuery;
    ADOQuery3: TADOQuery;
    DataSource1: TDataSource;
    Button1: TButton;
    ADOQuery4: TADOQuery;
    Panel4: TPanel;
    EdtModel: TEdit;
    DBGrid1: TDBGrid;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Edit2: TEdit;
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure PauseAutoScan1Click(Sender: TObject);
    procedure StopProgram1Click(Sender: TObject);
    procedure Exit2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure Button12Click(Sender: TObject);
    procedure Button13Click(Sender: TObject);
    procedure SettingComPort1Click(Sender: TObject);
    procedure AutoScan1Click(Sender: TObject);
    procedure SettingComPort2Click(Sender: TObject);
    procedure Autoscan2Click(Sender: TObject);
    procedure btnSetComClick(Sender: TObject);
    procedure ComPort1RxChar(Sender: TObject; Count: Integer);
    procedure AboutProgram2Click(Sender: TObject);
    procedure btnFirstJumpClick(Sender: TObject);
    procedure Edit1Change(Sender: TObject);
    procedure btnSecondJumpClick(Sender: TObject);
    procedure CloseComPort1Click(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure Timer4Timer(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);

  private
    { Private declarations }

  public
    { Public declarations }
     procedure Clear_Data();
  end;
  procedure ConnectNetwork() ;
  procedure InsertToDb(DataIndex : integer;m_code : integer) ;
  procedure Initail_dataSave();
  //procedure InitMax();

var
  Form1: TForm1;
  data_save : TextFile;
  I,cnt_led,col,End_Scn  : Integer;
  DataPoint : array[1..27] of String;
  str : AnsiString;
  Tstr,Tdata1,AllData,Str_Init,BoxNo,RcvData,TTime : String;
  G_count_index : integer ;
  G_m_code : integer ;

implementation

{$R *.dfm}

procedure TForm1.AboutProgram2Click(Sender: TObject);
var
    StrHelp : String;
begin
     StrHelp := ' Program data save in model'+#10;
     StrHelp := StrHelp+'  AQD 3010/3026  ';
     ShowMessage(StrHelp);
end;

procedure TForm1.AutoScan1Click(Sender: TObject);
begin
      ComPort1.Open;
      Timer1.Enabled := true;
end;

procedure TForm1.Autoscan2Click(Sender: TObject);
begin
       Timer1.Enabled := True;
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
begin
        Close;
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
begin
      Timer1.Enabled := True;
end;

procedure TForm1.btnSecondJumpClick(Sender: TObject);
begin
       //------------    Check ACK Data      -------------//
       if Tstr[3] = '/' then
       begin
              I := 4;
              RcvData := '';
              BoxNo := Tstr[2];
              if Tstr[4] = '3' then
              begin
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[1] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> '/' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[2] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[3] := RcvData;
                    RcvData := '';
                    Button3.Click;
              end;
       end
       else
        //begin
               ComPort1.WriteStr(IntToStr(cnt_led));
        //end;
end;

procedure TForm1.btnFirstJumpClick(Sender: TObject);
var
  Chk_again : integer;
begin
        edtBoxNo.Text := '';
        edtHis.Text := '';
        edtPosition.Text := '';
        edtPos_2.Text := '';
        edtPos6.Text := '';
        edtPos13.Text := '';
        edtPos19.Text := '';
        edtPos27.Text := '';
        edtPos33.Text := '';
        edtPos34.Text := '';
        edtPos35.Text := '';
        edtPos36.Text := '';
        edtPos37.Text := '';
        edtPos38.Text := '';
        edtPos42.Text := '';
        edtPos_22.Text := '';
        edtPos62.Text := '';
        edtPos132.Text := '';
        edtPos192.Text := '';
        edtPos272.Text := '';
        edtPos332.Text := '';
        edtPos342.Text := '';
        edtPos352.Text := '';
        edtPos362.Text := '';
        edtPos372.Text := '';
        edtPos382.Text := '';
        edtPos422.Text := '';
        Memo1.Text := '';
         Chk_again := pos('#',Tstr);

         if Chk_again <> 0 then
          begin
                  if Tstr[1] = '0' then
                      btnSecondJump.Click;
          end
         else
         // begin
                  ComPort1.WriteStr(IntToStr(cnt_led));
          //end;
end;

procedure TForm1.btnSetComClick(Sender: TObject);
begin
       ComPort1.Open;
       ComPort1.ShowSetupDialog;
end;

procedure TForm1.Button10Click(Sender: TObject);
begin
                 I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[16] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[17] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[18] := RcvData;
                    RcvData := '';
                    Button11.Click;
end;

procedure TForm1.Button11Click(Sender: TObject);
begin               I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[19] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[20] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[21] := RcvData;
                    RcvData := '';
                    Button12.Click;

end;

procedure TForm1.Button12Click(Sender: TObject);
begin
                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[22] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[23] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[24] := RcvData;
                    RcvData := '';
                    Button13.Click;
end;

procedure TForm1.Button13Click(Sender: TObject);
begin
                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                    end;
                    DataPoint[25] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          Application.ProcessMessages;
                          if Tstr[I] = '/' then
                          begin
                                  Break;
                          end
                          else
                          begin
                                RcvData := RcvData+Tstr[I];
                                I := I+1;
                          end;
                    end;
                    DataPoint[26] := RcvData;
                    RcvData := '';
                    I := 0;
                    Memo1.Text := '';
                    Button6.Click;
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
       try
            ADOQuery3.Active := false ;
            ADOQuery3.Parameters.ParamValues['a'] := EdtModel.Text ;
            ADOQuery3.Open ;

        if ADOQuery3.FieldValues['max'] <> null then
        begin
              G_count_index := ADOQuery3.FieldValues['max'] ;
        end
        else
        begin
              G_count_index := 0 ;
        end;

            ADOQuery4.Active := false ;
            ADOQuery4.Parameters.ParamValues['a'] := EdtModel.Text ;
            ADOQuery4.Open ;

        if ADOQuery4.RecordCount <> 0 then
        begin
              G_m_code := ADOQuery4.FieldValues['m_code'] ;
        end
        else
        begin
              G_m_code := 0 ;
              messagedlg('ไม่มี Product model นี้ในระบบ',mterror,[mbok],0);
        end;

    except

    end;
    Edit2.Text := inttostr(G_m_code);
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[4] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[5] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[6] := RcvData;
                    RcvData := '';
                    Button7.Click;

end;

procedure TForm1.Button4Click(Sender: TObject);
var
  k : Integer;
  StrTime  : string;
begin
    inc(G_count_index) ;
    InsertToDb(G_count_index,G_m_code);

 /////////////////////////////////////////////////////////////////////////
     StrTime := FormatdateTime('hh:mm:ss',now);
     try
        assignfile(data_save,FormatDateTime('yymmdd',now)+'.csv');
        AllData := StrTime+','+BoxNo+','+Tdata1;
        append(data_save);
        writeln(data_save,AllData);
        flush(data_save);
        CloseFile(data_save);
        Clear_Data;
        AllData := '';
        Tdata1 := '';
        Tstr := '';
        BoxNo := '';

     except
        AllData := StrTime+','+BoxNo+','+Tdata1;
        assignfile(data_save,FormatDateTime('yymmdd',now)+'.csv');
        rewrite(data_save);
        writeln(data_save,AllData);
        flush(data_save);
        CloseFile(data_save);
        Clear_Data;
        AllData := '';
        Tdata1 := '';
        Tstr := '';
        BoxNo := '';
     end;

   for k:= 0 to 27 do
     begin
          DataPoint[k] := '';
     end;
   End_Scn := 1;
end;

procedure TForm1.Clear_Data();
var
   clr_Client : string;
begin
      clr_Client := '';
      case Tstr[2] of
          '1':  clr_Client := 'A';
          '2':  clr_Client := 'B';
          '3':  clr_Client := 'C';
          '4':  clr_Client := 'D';
          '5':  clr_Client := 'E';
          '6':  clr_Client := 'F';
          '7':  clr_Client := 'G';
      end;
      ComPort1.WriteStr(clr_Client);
end;

procedure TForm1.Button5Click(Sender: TObject);
begin
      Str_Init := '';
      Str_Init := 'Timer,BoxNum,Position,Data,pos-2,pos6,pos13,pos19,pos27,pos33,pos34';
      Str_Init := Str_Init+',pos35,pos36,pos37,pos38,pos42';
      Str_Init := Str_Init+',pos-2,pos6,pos13,pos19,pos27,pos33,pos34';
      Str_Init := Str_Init+',pos35,pos36,pos37,pos38,pos42';

     try
        assignfile(data_save,FormatDateTime('yymmdd',now)+'.CSV');
        append(data_save);
        writeln(data_save,Str_Init);
        flush(data_save);
        CloseFile(data_save);
     except
        assignfile(data_save,FormatDateTime('yymmdd',now)+'.CSV');
        rewrite(data_save);
        append(data_save);
        writeln(data_save,Str_Init);
        flush(data_save);
        CloseFile(data_save);
     end;
end;

procedure TForm1.Button6Click(Sender: TObject);
begin
                 edtBoxNo.Text := BoxNo;
                 edtPosition.Text := DataPoint[1];
                 edtHis.Text := DataPoint[2];
                 edtPos_2.Text := DataPoint[3];
                 edtPos6.Text := DataPoint[4];
                 edtPos13.Text := DataPoint[5];
                 edtPos19.Text := DataPoint[6];
                 edtPos27.Text := DataPoint[7];
                 edtPos33.Text := DataPoint[8];
                 edtPos34.Text := DataPoint[9];
                 edtPos35.Text := DataPoint[10];
                 edtPos36.Text := DataPoint[11];
                 edtPos37.Text := DataPoint[12];
                 edtPos38.Text := DataPoint[13];
                 edtPos42.Text := DataPoint[14];
                 edtPos_22.Text := DataPoint[15];
                 edtPos62.Text := DataPoint[16];
                 edtPos132.Text := DataPoint[17];
                 edtPos192.Text := DataPoint[18];
                 edtPos272.Text := DataPoint[19];
                 edtPos332.Text := DataPoint[20];
                 edtPos342.Text := DataPoint[21];
                 edtPos352.Text := DataPoint[22];
                 edtPos362.Text := DataPoint[23];
                 edtPos372.Text := DataPoint[24];
                 edtPos382.Text := DataPoint[25];
                 edtPos422.Text := DataPoint[26];

           Tdata1 := edtPosition.Text+','+edtHis.Text+','+edtPos_2.Text+','+
                     edtPos6.Text+','+edtPos13.Text+','+edtPos19.Text+','+
                     edtPos27.Text+','+edtPos33.Text+','+edtPos34.Text+','+
                     edtPos35.Text+','+edtPos36.Text+','+edtPos37.Text+','+
                     edtPos38.Text+','+edtPos42.Text+','+edtPos_22.Text+','+
                     edtPos62.Text+','+edtPos132.Text+','+edtPos192.Text+','+
                     edtPos272.Text+','+edtPos332.Text+','+edtPos342.Text+','+
                     edtPos352.Text+','+edtPos362.Text+','+edtPos372.Text+','+
                     edtPos38.Text+','+edtPos42.Text;

          Button4.Click;
end;

procedure TForm1.Button7Click(Sender: TObject);
begin
                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[7] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[8] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[9] := RcvData;
                    RcvData := '';
                    Button8.Click;
end;

procedure TForm1.Button8Click(Sender: TObject);
begin
                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[10] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[11] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[12] := RcvData;
                    RcvData := '';
                    Button9.Click;
end;

procedure TForm1.Button9Click(Sender: TObject);
begin
                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[13] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[14] := RcvData;
                    RcvData := '';

                    I := I+1;
                    while Tstr[I] <> ',' do
                    begin
                          RcvData := RcvData+Tstr[I];
                          I := I+1;
                          Application.ProcessMessages;
                    end;
                    DataPoint[15] := RcvData;
                    RcvData := '';
                    Button10.Click;
end;

procedure TForm1.CloseComPort1Click(Sender: TObject);
begin
       ComPort1.Close;
end;

procedure TForm1.ComPort1RxChar(Sender: TObject; Count: Integer);
var ChkEnd : integer;
begin
      ComPort1.ReadStr(str,Count);
      Memo1.Text := Memo1.Text+str;
    //--------  CHECK LAST OF STRING DATA
            try
              ChkEnd := pos('#',Memo1.Text);
              if ChkEnd <> 0 then
                 begin
                       if Memo1.Text[1] =  'N' then
                                Memo2.Text := Memo1.Text[1]
                       else
                          begin
                                Tstr := Memo1.Text;
                                btnFirstJump.Click;
                                Memo2.Text := '';
                          end;
                 end;
            except

            end;
end;

procedure TForm1.Edit1Change(Sender: TObject);
begin
      Timer1.Interval := StrToIntdef(edit1.Text,2000);
end;

procedure TForm1.Exit1Click(Sender: TObject);
begin
        Close;
end;

procedure TForm1.Exit2Click(Sender: TObject);
begin
        Close;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
       cnt_led := 0;
       col := 0;
       End_Scn := 1;
       Timer1.Enabled := False;
       ComPort1.Close;
       ConnectNetwork();
       Initail_dataSave;

end;
procedure TForm1.FormShow(Sender: TObject);
begin
    Button1.Click ;
end;

procedure Initail_dataSave();
begin
      Str_Init := '';
      Str_Init := 'Timer,BoxNum,Position,Data,pos-2,pos6,pos13,pos19,pos27,pos33,pos34';
      Str_Init := Str_Init+',pos35,pos36,pos37,pos38,pos42';
      Str_Init := Str_Init+',pos-2,pos6,pos13,pos19,pos27,pos33,pos34';
      Str_Init := Str_Init+',pos35,pos36,pos37,pos38,pos42';

     try
        assignfile(data_save,FormatDateTime('yymmdd',now)+'.CSV');
        append(data_save);
        writeln(data_save,Str_Init);
        flush(data_save);
        CloseFile(data_save);
     except
        assignfile(data_save,FormatDateTime('yymmdd',now)+'.CSV');
        rewrite(data_save);
        append(data_save);
        writeln(data_save,Str_Init);
        flush(data_save);
        CloseFile(data_save);
     end;
end;

procedure TForm1.PauseAutoScan1Click(Sender: TObject);
begin
      Timer1.Enabled := false;
end;

procedure TForm1.SettingComPort1Click(Sender: TObject);
begin
       ComPort1.ShowSetupDialog;
end;

procedure TForm1.SettingComPort2Click(Sender: TObject);
begin
      ComPort1.ShowSetupDialog;
      ComPort1.Open;
end;

procedure TForm1.StopProgram1Click(Sender: TObject);
begin
      Timer1.Enabled := false;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
     try
       if End_Scn = 1 then
          begin
                lblCount.Caption := '';
                inc(cnt_led);
                inc(col);
                if cnt_led > 7 then
                    cnt_led := 1;
                case col of
                    1: lblLed.Color := clRed;
                    2: begin
                      col := 0;
                      lblLed.Color := clBlack;
                end;
          end;

            Memo1.Text := '';
            ComPort1.WriteStr(IntToStr(cnt_led));
            lblCount.Caption := lblCount.Caption+IntToStr((cnt_led));
            End_Scn := 0;
      end;
     except

     end;
end;

procedure TForm1.Timer2Timer(Sender: TObject);
begin
   TTime := FormatDateTime('HH:MM:SS',now);
   lblClock.Caption := TTime;
end;

procedure TForm1.Timer4Timer(Sender: TObject);
begin
        End_Scn := 1;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////                      //
//                                                    //                      //
//        ConnectNetwork.                             //                      //
//                                                    //                      //
////////////////////////////////////////////////////////                      //
////////////////////////////////////////////////////////////////////////////////
procedure ConnectNetwork() ;
begin

        try
            Form1.ADOConnection1.Open('user_setup','user_setup');
            //Gb_FlagNetwork := true  ;
        except
                MessageDlg('เชื่อมต่อกับฐานข้อมูลไม่ได้'+#13#10+
                           'Can not connect to database.',mtError,[mbOk],0);
        end;

end;

procedure InsertToDb(DataIndex : integer;m_code : integer) ;
var i ,j : integer ;
begin
    with Form1 do
    begin
          ADOQuery1.Active := false ;
          ADOQuery1.Parameters.ParamValues['a'] := m_code ; // m_code
          ADOQuery1.Parameters.ParamValues['b'] := edtBoxNo.Text ; // mc_no
          ADOQuery1.Parameters.ParamValues['c'] := edtPosition.Text ; // pos
          ADOQuery1.Parameters.ParamValues['d'] := '' ; // err_code
          ADOQuery1.Parameters.ParamValues['e'] := edtHis.Text ; // hysr
          //ADOQuery1.Parameters.ParamValues['f'] := Edit2.Text+' '+edttime.Text ; // date
          ADOQuery1.Parameters.ParamValues['g'] := DataIndex ; // data_index
          ADOQuery1.ExecSQL ;

        for i := 5 to 28 do
        begin
            j := i - 4 ;
            ADOQuery2.Active := false ;
            ADOQuery2.Parameters.ParamValues['a'] := m_code ;
            ADOQuery2.Parameters.ParamValues['b'] := DataIndex ;
            ADOQuery2.Parameters.ParamValues['c'] := j ;
            ADOQuery2.Parameters.ParamValues['d'] := DataPoint[i-2];     //Gb_Capture[i] ;
            ADOQuery2.ExecSQL ;
        end;
    end;
end;
end.
