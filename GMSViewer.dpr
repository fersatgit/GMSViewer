//Writed in Delphi 7
program GMSViewer;
{$R Resources.res}

uses
  Windows,Messages,ActiveX,CommCtrl;

const
  DelimiterWidth         =4; //Width of vertical delimiter in pixels
  ID_PROJECTSYSKIND      =1;
  ID_PROJECTCOMPATVERSION=$4A;
  ID_PROJECTLCID         =2;
  ID_PROJECTLCIDINVOKE   =$14;
  ID_PROJECTCODEPAGE     =3;
  ID_PROJECTNAME         =4;
  ID_PROJECTDOCSTRING    =5;
  ID_PROJECTHELPFILEPATH =6;
  ID_PROJECTHELPCONTEXT  =7;
  ID_PROJECTLIBFLAGS     =8;
  ID_PROJECTVERSION      =9;
  ID_PROJECTCONSTANTS    =$C;
  ID_REFERENCENAME       =$16;
  ID_REFERENCECONTROL    =$2F;
  ID_REFERENCEORIGINAL   =$33;
  ID_REFERENCEREGISTERED =$D;
  ID_REFERENCEPROJECT    =$E;
  ID_PROJECTMODULES      =$F;
  ID_PROJECTCOOKIE       =$13;
  ID_MODULENAME          =$19;
  ID_MODULENAMEUNICODE   =$47;
  ID_MODULESTREAMNAME    =$1A;
  ID_MODULEDOCSTRING     =$1C;
  ID_MODULEOFFSET        =$31;
  ID_MODULEHELPCONTEXT   =$1E;
  ID_MODULECOOKIE        =$2C;
  ID_MODULETYPE          =$21; //or $22
  ID_MODULETYPE1         =$21;
  ID_MODULETYPE2         =$22;
  ID_MODULEREADONLY      =$25;
  ID_MODULEPRIVATE       =$28;
  ID_MODULETERMINATOR    =$2B;
  MODULE_FLAG_PROCEDURAL =1;
  MODULE_FLAG_READONLY   =2;
  MODULE_FLAG_PRIVATE    =4;

var
  VdelimPos:            dword=350;
  StdCursor,HSizeCursor:dword;
  Streams,data:         array of byte;
  StreamsLen,datalen:   dword;
  InsertStruct:         tagTVINSERTSTRUCTW=(hInsertAfter:TVI_LAST;item:(mask:TVIF_TEXT+TVIF_PARAM));
  TreeView,Edit,Dialog: dword;
  ModulesCount:         dword;
  Modules:              array of packed record
                                 Name:       PWideChar;
                                 StreamName: PWideChar;
                                 StreamData: PAnsiChar;
                                 DataSize:   dword;
                                 Flags:      dword;
                                 TextOffset: dword;
                                 end;
  Dir:                  packed record
                        ProjectName:   PAnsiChar;
                        Constants:     PWideChar;
                        StreamData:    PAnsiChar;
                        DataSize:      dword;
                        SysKind:       PWideChar;
                        CompatVersion: dword;
                        VersionMajor:  dword;
                        VersionMinor:  dword;
                        CodePage:      dword;
                        end;
  rect:                 TRECT;

procedure error(text: PWideChar);
begin
  MessageBoxW(Dialog,text,0,0);
  ExitProcess(0);
end;

function UnpackStream(CompressedData, DecompressedData: pointer; compdatasize: dword): dword;
label
  finish;
var
 header:                                     word;
 i,j,k,token,FlagByte:                       integer;
 bitcount,Length,Offset,chunkstart,chunkend: integer;
 compdata: array of byte absolute CompressedData;
 data:     array of byte absolute DecompressedData;
const
  LengthMask: array[0..8] of dword=(15,31,63,127,255,511,1023,2047,4095);

  function calcbitcount(a: dword): dword;
  asm
    dec   eax
    bsr   eax,eax
    movzx eax,byte[@bitcount+eax+1]
    ret
    @bitcount: db 12,12,12,12,12,11,10,9,8,7,6,5,4
  end;

begin
  if compdata[0]=1 then
  begin
    chunkstart:=0;
    i:=1;
    repeat
      header:=pword(@compdata[i])^;
      if header and $3000<>$3000 then
        error('Данные повреждены');
      chunkend:=i+(header and $FFF)+3;
      if chunkend>compdatasize then
        chunkend:=compdatasize;
      if header and $8000=0 then
      begin
        MessageBoxW(0,'Обнаружены несжатые данные (этот код не тестировался)',0,0);
        move(compdata[i+2],data[chunkstart],4096);
        i:=chunkend;
        inc(chunkstart,4096);
      end
      else
      begin
        inc(i,2);
        j:=0;
        repeat
          FlagByte:=compdata[i];
          inc(i);
          for k:=7 downto 0 do
          begin
            if i>=chunkend then
              goto finish;
            if FlagByte and 1=0 then
            begin
              data[chunkstart+j]:=compdata[i];
              inc(i);
              inc(j);
            end
            else
            begin
              Token:=pword(@compdata[i])^;
              bitcount:=calcbitcount(j);
              Length:=(Token and LengthMask[BitCount-4])+3;
              Offset:=Token shr BitCount;
              repeat
                data[chunkstart+j]:=data[chunkstart+j-Offset-1];
                inc(j);
                dec(Length);
              until Length=0;
              inc(i,2);
            end;
            FlagByte:=FlagByte shr 1;
          end;
        until false;
        finish:
        inc(chunkstart,j);
      end;
    until i>=compdatasize;
  end
  else
    error('Данные повреждены');
  result:=chunkstart;
  pointer(compdata):=0;
  pointer(data):=0;
end;

procedure ReadStorage(const Storage: IStorage; parent: HTREEITEM);
const
  SysKind: array[0..4] of PWideChar=('Win16','Win32','MAC','Win64','Unknown');
var
  statstg:     tagSTATSTG;
  EnumSTATSTG: IEnumSTATSTG;
  Storage2:    IStorage;
  itemHandle:  HTREEITEM;
  Stream:      IStream;
  StreamLen:   int64;
  i,j,size:    integer;
begin
  Storage.EnumElements(0,0,0,EnumSTATSTG);
  while EnumSTATSTG.Next(1,statstg,@StreamLen)=0 do
  begin
    InsertStruct.item.pszText:=statstg.pwcsName;
    InsertStruct.hParent:=parent;
    case statstg.dwType of
    STGTY_STORAGE:begin
                  InsertStruct.item.lParam:=-1;
                  itemHandle:=HTREEITEM(SendMessageW(TreeView,TVM_INSERTITEMW,0,LongInt(@InsertStruct)));
                  Storage.OpenStorage(statstg.pwcsName,nil,STGM_READ+STGM_SHARE_EXCLUSIVE,0,0,Storage2);
                  ReadStorage(Storage2,itemHandle);
                  Storage2._Release;
                  pointer(Storage2):=0;
                  end;
     STGTY_STREAM:begin
                  Storage.OpenStream(statstg.pwcsName,0,STGM_READ+STGM_SHARE_EXCLUSIVE,0,Stream);
                  Stream.Seek(0,STREAM_SEEK_END,StreamLen);
                  pointer(InsertStruct.item.lParam):=@Streams[StreamsLen];
                  Pdword(InsertStruct.item.lParam)^:=StreamLen;
                  Stream.Seek(0,STREAM_SEEK_SET,pint64(0)^);
                  Stream.Read(pointer(InsertStruct.item.lParam+4),StreamLen,@StreamLen);
                  inc(StreamsLen,(StreamLen+5) and -2); //align to 2 for unicode strings

                  for i:=ModulesCount-1 downto 0 do
                   if lstrcmpW(Modules[i].StreamName,statstg.pwcsName)=0 then
                     with Modules[i] do
                     begin
                       StreamData:=pointer(InsertStruct.item.lParam+4);
                       DataSize:=StreamLen;
                       InsertStruct.item.lParam:=i;
                       break;
                     end;

                  if (lstrcmpW(InsertStruct.item.pszText,'PROJECT')=0)or(lstrcmpW(InsertStruct.item.pszText,#3'VBFrame')=0) then
                    pdword(InsertStruct.item.lParam)^:=pdword(InsertStruct.item.lParam)^ or $80000000
                  else if (lstrcmpW(statstg.pwcsName,'dir')=0)and(Dir.ProjectName=nil) then
                  begin
                    StreamLen:=UnpackStream(pointer(InsertStruct.item.lParam+4),data,StreamLen);
                    i:=0;
                    j:=0;
                    repeat
                      size:=pdword(@data[i+2])^;
                      case pword(@data[i])^ of
                              ID_PROJECTSYSKIND:begin
                                                size:=pdword(@data[i+6])^;
                                                if size>3 then
                                                  size:=4;
                                                Dir.SysKind:=SysKind[size];
                                                inc(i,10);
                                                end;
                        ID_PROJECTCOMPATVERSION:begin
                                                Dir.CompatVersion:=pdword(@data[i+6])^;
                                                inc(i,10);
                                                end;
                                 ID_PROJECTLCID,
                          ID_PROJECTHELPCONTEXT,
                             ID_PROJECTLIBFLAGS,
                           ID_MODULEHELPCONTEXT,
                           ID_PROJECTLCIDINVOKE:inc(i,10);
                             ID_PROJECTCODEPAGE:begin
                                                Dir.CodePage:=pword(@data[i+6])^;
                                                inc(i,8);
                                                end;
                                 ID_PROJECTNAME:begin
                                                move(data[i+6],Streams[StreamsLen],size);
                                                Dir.ProjectName:=@Streams[StreamsLen];
                                                SendMessageA(Dialog,WM_SETTEXT,0,LongInt(Dir.ProjectName));
                                                inc(StreamsLen,(size+2) and -2); //align to 2 for unicode strings
                                                Streams[StreamsLen-1]:=0;
                                                inc(i,size+6);
                                                end;
                            ID_PROJECTDOCSTRING,
                               ID_REFERENCENAME,
                             ID_MODULEDOCSTRING,
                         ID_PROJECTHELPFILEPATH:begin
                                                inc(i,size+8);              //skip ansi
                                                inc(i,pdword(@data[i])^+4); //skip unicode
                                                end;
                              ID_PROJECTVERSION:begin
                                                Dir.VersionMajor:=pdword(@data[i+6])^;
                                                Dir.VersionMajor:=pdword(@data[i+10])^;
                                                inc(i,12);
                                                end;
                            ID_PROJECTCONSTANTS:begin
                                                inc(i,size+8);              //skip ansi
                                                size:=pdword(@data[i])^;
                                                move(data[i+4],Streams[StreamsLen],size);
                                                Dir.Constants:=@Streams[StreamsLen];
                                                inc(StreamsLen,size+2);
                                                pword(@Streams[StreamsLen-2])^:=0;
                                                inc(i,size+4);
                                                end;
                            ID_REFERENCECONTROL:begin
                                                inc(i,size+6);             //skip to NameRecordExtended
                                                size:=pdword(@data[i+2])^;
                                                inc(i,size+8);             //skip NameRecordExtended.Name
                                                inc(i,pdword(@data[i])^+6);//skip to SizeExtended
                                                inc(i,pdword(@data[i])^+4);
                                                end;
                           ID_REFERENCEORIGINAL:begin
                                                inc(i,size+6);             //skip to ReferenceRecord
                                                size:=pdword(@data[i+2])^;
                                                inc(i,size+6);             //skip to NameRecordExtended
                                                size:=pdword(@data[i+2])^;
                                                inc(i,size+8);             //skip NameRecordExtended.Name
                                                inc(i,pdword(@data[i])^+6);//skip to SizeExtended
                                                inc(i,pdword(@data[i])^+4);
                                                end;
                         ID_REFERENCEREGISTERED,
                                  ID_MODULENAME,
                            ID_REFERENCEPROJECT:inc(i,size+6);
                              ID_PROJECTMODULES:begin
                                                ModulesCount:=pword(@data[i+6])^;
                                                pointer(Modules):=VirtualAlloc(0,ModulesCount*sizeof(Modules[0]),MEM_COMMIT,PAGE_READWRITE);
                                                inc(i,16)
                                                end;
                                ID_MODULECOOKIE,
                               ID_PROJECTCOOKIE:inc(i,8);
                           ID_MODULENAMEUNICODE:begin
                                                if ModulesCount=0 then
                                                begin
                                                  ModulesCount:=65535;
                                                  pointer(Modules):=VirtualAlloc(0,65535*sizeof(Modules[0]),MEM_COMMIT,PAGE_READWRITE);
                                                end;
                                                move(data[i+6],Streams[StreamsLen],size);
                                                Modules[j].Name:=@Streams[StreamsLen];
                                                inc(StreamsLen,size+2);
                                                pword(@Streams[StreamsLen-2])^:=0;
                                                inc(i,size+6);
                                                end;
                            ID_MODULESTREAMNAME:begin
                                                inc(i,size+8);              //skip ansi
                                                size:=pdword(@data[i])^;
                                                move(data[i+4],Streams[StreamsLen],size);
                                                Modules[j].StreamName:=@Streams[StreamsLen];
                                                inc(StreamsLen,size+2);
                                                pword(@Streams[StreamsLen-2])^:=0;
                                                inc(i,size+4);
                                                end;
                                ID_MODULEOFFSET:begin
                                                Modules[j].TextOffset:=pdword(@data[i+6])^;
                                                inc(i,10);
                                                end;
                                 ID_MODULETYPE1:begin
                                                Modules[j].Flags:=Modules[j].Flags or MODULE_FLAG_PROCEDURAL;
                                                inc(i,6);
                                                end;
                                 ID_MODULETYPE2:inc(i,6);
                              ID_MODULEREADONLY:begin
                                                Modules[j].Flags:=Modules[j].Flags or MODULE_FLAG_READONLY;
                                                inc(i,6);
                                                end;
                               ID_MODULEPRIVATE:begin
                                                Modules[j].Flags:=Modules[j].Flags or MODULE_FLAG_PRIVATE;
                                                inc(i,6);
                                                end;
                            ID_MODULETERMINATOR:begin
                                                inc(i,6);
                                                inc(j);
                                                if j=ModulesCount then
                                                  break;
                                                end;
                       else if ModulesCount>0 then
                       begin
                         ModulesCount:=j;
                         break;
                       end
                       else
                         error('Неожиданный id внутри потока директорий');
                       end;
                    until i>=StreamLen;
                    
                    Dir.StreamData:=pointer(InsertStruct.item.lParam+4);
                    Dir.DataSize:=StreamLen;
                    pointer(InsertStruct.item.lParam):=@Dir;
                  end;

                  SendMessageW(TreeView,TVM_INSERTITEMW,0,LongInt(@InsertStruct));
                  Stream._Release;
                  pointer(Stream):=0;
                  end;
    end;
    CoTaskMemFree(statstg.pwcsName);
  end;
  EnumSTATSTG._Release;
  pointer(EnumSTATSTG):=0;
end;

function DlgFunc(wnd,msg,wParam,lParam: dword):dword;stdcall;
const
  hex:    array[0..15] of AnsiChar='0123456789ABCDEF';
  DirFmt: PWideChar='Constants:%n%2!s!%nSysKind: %5!s!%nCompatVersion: %6!u!%nVersion: %7!u!.%8!u!%nCodePage: %9!u!'#0;
var
  i,j,k:     integer;
  LockBytes: ILockBytes;
  Storage:   IStorage;
  tmp:       PWideChar;
  hGlobal,f: dword;
  endChar:   WideChar;
begin
  result:=0;
  case msg of
      WM_NOTIFY:with PNMTREEVIEWW(lParam)^ do
                  if (hdr.code=TVN_SELCHANGEDW)and(itemNew.lParam<>-1) then //if not a storage
                    if itemNew.lParam=LongInt(@Dir) then                    //dir stream
                    begin
                      FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER+FORMAT_MESSAGE_ARGUMENT_ARRAY+FORMAT_MESSAGE_FROM_STRING,DirFmt,0,0,@tmp,0,@Dir);
                      SendMessageW(Edit,WM_SETTEXT,0,LongInt(tmp));
                      LocalFree(LongInt(tmp));
                    end
                    else if dword(itemNew.lParam)<ModulesCount then         //module stream
                      with Modules[itemNew.lParam] do
                      begin
                        i:=UnpackStream(@StreamData[TextOffset],data,DataSize-TextOffset);
                        data[i]:=0;
                        tmp:=VirtualAlloc(0,i*2+2,MEM_COMMIT,PAGE_READWRITE);
                        MultiByteToWideChar(Dir.CodePage,0,@data[0],i,tmp,i+1);
                        SendMessageW(Edit,WM_SETTEXT,0,LongInt(tmp));
                        VirtualFree(tmp,0,MEM_RELEASE);
                      end
                    else if pdword(itemNew.lParam)^ and $80000000>0 then    //Project or VBFrame streams
                    begin
                      i:=pdword(itemNew.lParam)^ and $7FFFFFFF;
                      tmp:=VirtualAlloc(0,i*2+2,MEM_COMMIT,PAGE_READWRITE);
                      MultiByteToWideChar(Dir.CodePage,0,pointer(itemNew.lParam+4),i,tmp,i+1);
                      SendMessageW(Edit,WM_SETTEXT,0,LongInt(tmp));
                      VirtualFree(tmp,0,MEM_RELEASE);
                    end
                    else                                                    //Other stream (HEX)
                    begin
                      i:=pdword(itemNew.lParam)^;
                      tmp:=VirtualAlloc(0,i*6,MEM_COMMIT,PAGE_READWRITE);
                      for j:=0 to i-1 do
                      begin
                        k:=pbyte(itemNew.lParam+j+4)^;
                        tmp[j*3+0]:=WideChar(hex[k shr 4]);
                        tmp[j*3+1]:=WideChar(hex[k and 15]);
                        tmp[j*3+2]:=' ';
                      end;
                      tmp[i*3-1]:=#0;
                      SendMessageW(Edit,WM_SETTEXT,0,LongInt(tmp));
                      VirtualFree(tmp,0,MEM_RELEASE);
                    end;
 WM_LBUTTONDOWN:SetCapture(wnd);
   WM_LBUTTONUP:ReleaseCapture;
   WM_MOUSEMOVE:if (wParam and MK_LBUTTON>0)and(dword(Rect.Right-DelimiterWidth-loword(lParam))<rect.Right) then
                begin
                  VdelimPos:=loword(lParam)-(DelimiterWidth shr 1);
                  DlgFunc(wnd,WM_SIZE,0,0);
                end
                else if dword(loword(lParam)-VdelimPos)<DelimiterWidth then
                  SetCursor(HSizeCursor)
                else
                  SetCursor(StdCursor);
        WM_SIZE:begin
                GetClientRect(wnd,rect);
                dec(rect.Bottom,6);
                MoveWindow(TreeView,3,3,VDelimPos-3,rect.Bottom,true);
                MoveWindow(Edit,VDelimPos+DelimiterWidth,3,rect.Right-(VDelimPos+DelimiterWidth+3),rect.Bottom,true);
                end;
  WM_INITDIALOG:begin
                Dialog:=wnd;
                TreeView:=GetDlgItem(wnd,1);
                Edit:=GetDlgItem(wnd,2);

                tmp:=GetCommandLineW;
                i:=0;
                repeat
                  inc(i);
                until ((tmp[i]='"')and(tmp[i+1]=' '))or(tmp[i]=#0);
                if (tmp[i]=#0)or(tmp[i+2]=#0) then
                  error('Командная строка пуста')
                else
                begin
                  inc(i,2);
                  endChar:=' ';
                  if tmp[i]='"' then
                  begin
                    endChar:='"';
                    inc(i);
                  end;
                  j:=i;
                  repeat
                    inc(j);
                  until (tmp[j]=EndChar)or(tmp[j]=#0);
                  tmp[j]:=#0;
                end;

                f:=CreateFileW(@tmp[i],GENERIC_READ,0,0,OPEN_EXISTING,0,0);
                if f=INVALID_HANDLE_VALUE then
                  error('Не удалось открыть файл');
                datalen:=GetFileSize(f,0);
                pointer(data):=VirtualAlloc(0,datalen,MEM_COMMIT,PAGE_READWRITE);
                if data=nil then
                  error('Не достаточно памяти');
                ReadFile(f,data[0],datalen,datalen,0);
                CloseHandle(f);

                if pdword(data)^and $FFFFFF=$534D47 then //GMS signature
                begin
                  f:=pdword(@data[4])^+8; //skip GMS header
                  dec(datalen,f);
                  hGlobal:=GlobalAlloc(GMEM_MOVEABLE,datalen);
                  if hGlobal=0 then
                    error('Не достаточно памяти');
                  move(data[f],GlobalLock(hGlobal)^,datalen);
                  GlobalUnlock(hGlobal);
                  CreateILockBytesOnHGlobal(hGlobal,true,LockBytes);
                  StgOpenStorageOnILockBytes(LockBytes,nil,STGM_SHARE_DENY_WRITE,0,0,Storage);
                  if Storage=nil then
                    error('Не удалось открыть хранилище IStorage, возможно данные повреждены');
                  pointer(Streams):=VirtualAlloc(0,datalen,MEM_COMMIT,PAGE_READWRITE);
                  if Streams=nil then
                    error('Не достаточно памяти');
                  ReadStorage(Storage,TVI_ROOT);
                  Storage._Release;
                  pointer(Storage):=0;
                  LockBytes._Release;
                  pointer(LockBytes):=0;
                  GlobalFree(hGlobal);
                end
                else
                  error('Отсутствует сигнатура GMS');

                SetClassLongW(wnd,GCL_HICON,LoadIconW($400000,PWideChar(1)));
                StdCursor:=LoadCursorW(0,pointer(IDC_ARROW));
                HSizeCursor:=LoadCursorW(0,pointer(IDC_SIZEWE));
                DlgFunc(wnd,WM_SIZE,0,0);
                result:=1;
                end;
       WM_CLOSE:begin
                //System will automaticaly freee memory when exit process
                //VirtualFree(pointer(Streams),0,MEM_RELEASE);
                //VirtualFree(pointer(data),0,MEM_RELEASE);
                pointer(Streams):=0; //Zeroing pointers to avoid delphi garbage collector glitches
                pointer(data):=0;
                ExitProcess(0);
                end;
  end;
end;

begin
  InitCommonControls;
  DialogBoxParamW($400000,PWideChar(1),0,@DlgFunc,0);
end.
