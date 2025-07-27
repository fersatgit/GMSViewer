//Writed in Delphi 7
//To compile from IDE press Ctrl+F9
//This is not a unit but a program. This hack reduses output file size.
unit GMSViewer;
{$IMAGEBASE $400000}{$E .exe}{$G-}{$R-}{$I-}{$M-}{$Y-}{$D-}{$C-}{$L-}{$Q-}{$O+}
interface
implementation

uses
  Windows,Messages,ActiveX,CommCtrl,CommDlg;

{$I OVBA.inc}
{$I Processing.inc}
{$R Resources.res}

type
  PNMTVKEYDOWN=^tagTVKEYDOWN;
  _TBBUTTON = packed record
              iBitmap:   Integer;
              idCommand: Integer;
              fsState:   Byte;
              fsStyle:   Byte;
              bReserved: array[1..2] of Byte;
              dwData:    Longint;
              iString:   PWideChar;
              end;

const
  DelimiterWidth=4; //Width of vertical delimiter in pixels
  CMD_SAVE      =0;
  CMD_PROCESS   =1;

var
  VdelimPos:            integer=350;
  StdCursor,HSizeCursor:THandle;
  savefilename:         array[0..1023] of WideChar;
  filename:             PWideChar;
  ofn:                  OPENFILENAMEW=(lStructSize:sizeof(OPENFILENAMEW);lpstrFilter:#0#0#0;lpstrFile: @savefilename;nMaxFile:length(savefilename));
  rect:                 TRECT;

procedure ReadStorage(const Storage: IStorage; parent: dword);
var
  statstg:     tagSTATSTG;
  EnumSTATSTG: IEnumSTATSTG;
  Storage2:    IStorage;
  Stream:      IStream;
  StreamLen:   int64;
begin
  Storage.EnumElements(0,0,0,EnumSTATSTG);
  while EnumSTATSTG.Next(1,statstg,@StreamLen)=0 do
  begin
    InsertStruct.item.pszText:=statstg.pwcsName;
    InsertStruct.hParent:=Elements[parent].TreeViewHandle;
    InsertStruct.item.lParam:=ElementsCount;
    Elements[ElementsCount].Parent:=parent;
    Elements[ElementsCount].Name:=statstg.pwcsName;
    Elements[ElementsCount].StreamName:=statstg.pwcsName;    
    Elements[ElementsCount].TreeViewHandle:=HTREEITEM(SendMessageW(TreeView,TVM_INSERTITEMW,0,LongInt(@InsertStruct)));
    case statstg.dwType of
    STGTY_STORAGE:begin
                  Elements[ElementsCount].StreamType:=STREAM_TYPE_STORAGE;
                  Storage.OpenStorage(statstg.pwcsName,nil,STGM_READ+STGM_SHARE_EXCLUSIVE,0,0,Storage2);
                  inc(ElementsCount);
                  ReadStorage(Storage2,ElementsCount-1);
                  Storage2._Release;
                  pointer(Storage2):=0;
                  end;
     STGTY_STREAM:begin
                  Elements[ElementsCount].StreamType:=STREAM_TYPE_UNKNOWN;
                  Storage.OpenStream(statstg.pwcsName,0,STGM_READ+STGM_SHARE_EXCLUSIVE,0,Stream);
                  Stream.Seek(0,STREAM_SEEK_END,StreamLen);
                  Stream.Seek(0,STREAM_SEEK_SET,pint64(0)^);
                  Elements[ElementsCount].StreamData:=HeapAlloc(Heap,0,StreamLen);
                  Stream.Read(Elements[ElementsCount].StreamData,StreamLen,@StreamLen);
                  Elements[ElementsCount].DataSize:=StreamLen;
                  if lstrcmpW(InsertStruct.item.pszText,'PROJECT')=0 then
                  begin
                    Elements[ElementsCount].StreamType:=STREAM_TYPE_TEXT;
                    Project:=ElementsCount;
                  end
                  else if lstrcmpW(InsertStruct.item.pszText,#3'VBFrame')=0 then
                    Elements[ElementsCount].StreamType:=STREAM_TYPE_TEXT
                  else if lstrcmpW(statstg.pwcsName,'dir')=0 then
                  begin
                    Elements[ElementsCount].StreamType:=STREAM_TYPE_DIR;
                    Dir.ElementIndex:=ElementsCount;
                    Dir.StreamData  :=Elements[ElementsCount].StreamData;
                    Dir.DataSize    :=Elements[ElementsCount].DataSize;
                  end;
                  Stream._Release;
                  pointer(Stream):=0;
                  inc(ElementsCount);
                  end;
    end;
  end;
  EnumSTATSTG._Release;
  pointer(EnumSTATSTG):=0;
end;

function WriteStorage(const Storage: IStorage; i: LongInt): dword;
var
  Storage2: IStorage;
  Stream:   IStream;
  j:        LongInt;
  parent:   dword;
begin
  parent:=Elements[i].Parent;
  while (i<ElementsCount)and(Elements[i].Parent=parent) do
    with Elements[i] do
    begin
      if StreamType=STREAM_TYPE_STORAGE then
      begin
        Storage.CreateStorage(Name,STGM_CREATE+STGM_WRITE+STGM_SHARE_EXCLUSIVE,0,0,Storage2);
        i:=WriteStorage(Storage2,i+1);
        Storage2._Release;
        pointer(Storage2):=0;
      end
      else
      begin
        Storage.CreateStream(StreamName,STGM_CREATE+STGM_WRITE+STGM_SHARE_EXCLUSIVE,0,0,Stream);
        Stream.Write(StreamData,DataSize,@j);
        Stream._Release;
        pointer(Stream):=0;
        inc(i);
      end;
    end;
  result:=i;
end;

//Transfers data from edit (if modified) to tree item by index
procedure EditToItem(index: dword);
var
  tmp: pointer;
begin
  if SendMessageW(Edit,EM_GETMODIFY,0,0)>0 then
    with Elements[index] do
    begin
      DataSize:=SendMessageW(Edit,WM_GETTEXTLENGTH,0,0);
      SendMessageW(Edit,WM_GETTEXT,DataSize,LongInt(data));
      StreamData:=HeapRealloc(Heap,0,StreamData,TextOffset+DataSize*2+3);
      tmp:=pointer(SendMessageW(edit,EM_GETHANDLE,0,0));
      WideCharToMultiByte(Dir.CodePage,0,LocalLock(LongInt(tmp)),DataSize,@data[0],DataSize,0,0);
      LocalUnlock(LongInt(tmp));
      if StreamType=STREAM_TYPE_PACKED_TEXT then
        DataSize:=Compress(data,@StreamData[TextOffset],DataSize)+TextOffset
      else
        move(data[0],StreamData[0],DataSize);
      StreamData:=HeapRealloc(Heap,0,StreamData,DataSize);
    end;
end;

function DlgFunc(wnd,msg,wParam,lParam: dword):dword;stdcall;
const
  hex:          array[0..15] of AnsiChar='0123456789ABCDEF';
  DirFmt:       PWideChar='Conditional Compilation Arguments: %1!s!%nSysKind: %3!u! (%2!s!)%nVersion: %4!u!.%5!u!%nCodePage: %6!u!%nLCID: %7!u!%nReferences:%n%8!s!'#0;
  SysKind:      array[0..4] of PWideChar=('Win16','Win32','MAC','Win64','Unknown');
  GMSHeader:    array[0..17] of byte=($47,$4D,$53,$01,$0A,$00,$00,$00,$02,$00,$00,$00,$00,$00,$01,$00,$00,$00);
  tbbuttons:    array[0..1] of _TBBUTTON=((iBitmap: 0; idCommand: CMD_SAVE;    fsState: TBSTATE_ENABLED; fsStyle: BTNS_SHOWTEXT; iString: 'Сохранить'),
                                          (iBitmap: 1; idCommand: CMD_PROCESS; fsState: TBSTATE_ENABLED; fsStyle: BTNS_BUTTON; iString: 'Обработка'));
var
  LockBytes:          ILockBytes;
  Storage:            IStorage;
  tmp:                PWideChar;
  hGlobal,f,i,j,k,n:  LongInt;
  size:               dword;
  endChar:            WideChar;
  statstg:            tagSTATSTG;
  ModuleNameLen:      dword;
  ModuleName:         TModuleName;
  hittestinfo:        tagTVHITTESTINFO;
begin
  result:=0;
  case msg of
     WM_COMMAND:case wParam of
                     CMD_SAVE:begin
                              EditToItem(SelectedItem);
                              StgCreateDocfile(0,STGM_CREATE+STGM_WRITE+STGM_SHARE_EXCLUSIVE,0,Storage);
                              WriteStorage(Storage,0);
                              Storage.Stat(statstg,0);
                              Storage._Release;
                              pointer(Storage):=0;
                              f:=CreateFileW(statstg.pwcsName,GENERIC_READ,0,0,OPEN_EXISTING,0,0);
                              size:=GetFileSize(f,0);
                              tmp:=VirtualAlloc(0,size,MEM_COMMIT,PAGE_READWRITE);
                              ReadFile(f,tmp^,size,size,0);
                              CloseHandle(f);
                              DeleteFileW(statstg.pwcsName);
                              CoTaskMemFree(statstg.pwcsName);
                              i:=strcpyW(savefilename,filename);
                              pdword(@savefilename[i])^  :=$62002E;//'.b'
                              pdword(@savefilename[i+2])^:=$6B0061;//'ak'
                              MoveFileW(filename,savefilename);
                              f:=CreateFileW(filename,GENERIC_WRITE,0,0,CREATE_ALWAYS,0,0);
                              WriteFile(f,GMSHeader,sizeof(GMSHeader),pdword(@j)^,0);
                              WriteFile(f,tmp^,size,pdword(@j)^,0);
                              CloseHandle(f);
                              VirtualFree(tmp,0,MEM_RELEASE);
                              end;
                  CMD_PROCESS:DialogBoxParamW($400000,PWideChar(2),wnd,@ProcessingDlgFunc,0);
                end;
      WM_NOTIFY:if PNMHDR(lParam)^.hwndFrom=TreeView then
                ///////////////////////////////////////////////////////////////
                //Context menu
                ///////////////////////////////////////////////////////////////
                  if PNMHDR(lParam)^.code=NM_RCLICK then
                  begin
                    GetCursorPos(hittestinfo.pt);
                    ScreenToClient(TreeView,hittestinfo.pt);
                    SendMessageW(TreeView,TVM_HITTEST,0,LongInt(@hittestinfo));
                    if LongInt(hittestinfo.hItem)>0 then
                    begin
                      SendMessageW(TreeView,TVM_SELECTITEM,TVGN_CARET,LongInt(hittestinfo.hItem));
                      with Elements[SelectedItem] do
                      begin
                        GetCursorPos(hittestinfo.pt);
                        TreeMenu:=CreatePopupMenu();
                        AppendMenu(TreeMenu, MF_STRING,1,'Удалить');
                        if StreamType<>STREAM_TYPE_STORAGE then
                        begin
                          AppendMenu(TreeMenu, MF_STRING,2,'Сохранить');
                          if (StreamData[TextOffset]=#1)and(pbyte(@StreamData[TextOffset+2])^ and $F0=$B0) then
                            AppendMenu(TreeMenu, MF_STRING,3,'Распаковать');
                        end;
                        case dword(TrackPopupMenu(TreeMenu,TPM_RETURNCMD+TPM_NOANIMATION,hittestinfo.pt.x,hittestinfo.pt.y,0,TreeView,0)) of
                        1:DeleteItem(SelectedItem);
                        2:if GetSaveFileNameW(ofn) then
                          begin
                            f:=CreateFileW(@savefilename,GENERIC_WRITE,0,0,CREATE_ALWAYS,0,0);
                            WriteFile(f,StreamData^,DataSize,size,0);
                            CloseHandle(f);
                          end;
                        3:if GetSaveFileNameW(ofn) then
                          begin
                            f:=CreateFileW(@savefilename,GENERIC_WRITE,0,0,CREATE_ALWAYS,0,0);
                            size:=Decompress(@StreamData[TextOffset],data,DataSize-TextOffset);
                            WriteFile(f,data[0],size,size,0);
                            CloseHandle(f);
                          end;
                        end;
                        DestroyMenu(TreeMenu);
                      end;
                    end;
                  end
                ///////////////////////////////////////////////////////////////
                //Delete key
                ///////////////////////////////////////////////////////////////
                  else if (PNMHDR(lParam)^.code=TVN_KEYDOWN)and(PNMTVKEYDOWN(lParam)^.wVKey=VK_DELETE) then
                    DeleteItem(SelectedItem)
                ///////////////////////////////////////////////////////////////
                //TreeView selection change
                ///////////////////////////////////////////////////////////////
                  else if PNMHDR(lParam)^.code=TVN_SELCHANGEDW then
                  begin
                    EditToItem(PNMTREEVIEWW(lParam)^.itemOld.lParam);
                    SelectedItem:=PNMTREEVIEWW(lParam)^.itemNew.lParam;
                    SendMessageW(Edit,EM_SETREADONLY,1,0);
                    with Elements[SelectedItem] do
                      case StreamType of
                              STREAM_TYPE_DIR:begin
                                              FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER+FORMAT_MESSAGE_ARGUMENT_ARRAY+FORMAT_MESSAGE_FROM_STRING,DirFmt,0,0,@tmp,0,@Dir);
                                              SendMessageW(Edit,WM_SETTEXT,0,LongInt(tmp));
                                              LocalFree(LongInt(tmp));
                                              end;
                      STREAM_TYPE_PACKED_TEXT:begin
                                              i:=Decompress(@StreamData[TextOffset],data,DataSize-TextOffset);
                                              data[i]:=0;
                                              tmp:=VirtualAlloc(0,i*2+2,MEM_COMMIT,PAGE_READWRITE);
                                              MultiByteToWideChar(Dir.CodePage,0,@data[0],i,tmp,i+1);
                                              SendMessageW(Edit,WM_SETTEXT,0,LongInt(tmp));
                                              VirtualFree(tmp,0,MEM_RELEASE);
                                              SendMessageW(Edit,EM_SETREADONLY,0,0);
                                              end;
                             STREAM_TYPE_TEXT:begin
                                              tmp:=VirtualAlloc(0,DataSize*2+2,MEM_COMMIT,PAGE_READWRITE);
                                              MultiByteToWideChar(Dir.CodePage,0,StreamData,DataSize,tmp,DataSize+1);
                                              SendMessageW(Edit,WM_SETTEXT,0,LongInt(tmp));
                                              VirtualFree(tmp,0,MEM_RELEASE);
                                              SendMessageW(Edit,EM_SETREADONLY,0,0);
                                              end;
                          STREAM_TYPE_UNKNOWN:begin
                                              tmp:=VirtualAlloc(0,DataSize*6,MEM_COMMIT,PAGE_READWRITE);
                                              for j:=0 to DataSize-1 do
                                              begin
                                                k:=pbyte(LongInt(StreamData)+j)^;
                                                tmp[j*3+0]:=WideChar(hex[k shr 4]);
                                                tmp[j*3+1]:=WideChar(hex[k and 15]);
                                                tmp[j*3+2]:=' ';
                                              end;
                                              tmp[DataSize*3-1]:=#0;
                                              SendMessageW(Edit,WM_SETTEXT,0,LongInt(tmp));
                                              VirtualFree(tmp,0,MEM_RELEASE);
                                              end;
                      else
                        SendMessageW(Edit,WM_SETTEXT,0,0);
                      end;
                      SendMessageW(Edit,EM_SETMODIFY,0,0);
                  end;
 WM_LBUTTONDOWN:SetCapture(wnd);
   WM_LBUTTONUP:ReleaseCapture;
   WM_MOUSEMOVE:if (wParam and MK_LBUTTON>0)and(dword(Rect.Right-DelimiterWidth-loword(lParam))<dword(rect.Right)) then
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
                dec(rect.Bottom,36);
                MoveWindow(Toolbar,0,0,rect.Right,30,true);
                MoveWindow(TreeView,3,30,VDelimPos-3,rect.Bottom,true);
                MoveWindow(Edit,VDelimPos+DelimiterWidth,30,rect.Right-(VDelimPos+DelimiterWidth+3),rect.Bottom,true);
                end;
  WM_INITDIALOG:begin
                Dialog:=wnd;
                ofn.hWndOwner:=wnd;
                TreeView:=GetDlgItem(wnd,1);
                Edit:=GetDlgItem(wnd,2);
                Toolbar:=GetDlgItem(wnd,3);

                tmp:=GetCommandLineW;
                i:=0;
                endChar:=' ';
                if tmp[0]='"' then
                  endChar:='"';
                repeat
                  inc(i);
                until (tmp[i]=endChar)or(tmp[i]=#0);
                if (tmp[i]=#0)or(tmp[i+1]=#0) then
                  error('Командная строка пуста')
                else
                begin
                  repeat
                    inc(i);
                  until tmp[i]<>' ';
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
                filename:=@tmp[i];

                f:=CreateFileW(filename,GENERIC_READ,0,0,OPEN_EXISTING,0,0);
                if f=-1 then
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
                  pointer(Elements):=VirtualAlloc(0,65536*sizeof(Elements[0]),MEM_COMMIT,PAGE_READWRITE);
                  Heap:=HeapCreate(0,datalen+datalen,0);
                  if (Elements=nil)or(hGlobal=0)or(Heap=0) then
                    error('Не достаточно памяти');
                  move(data[f],GlobalLock(hGlobal)^,datalen);
                  GlobalUnlock(hGlobal);
                  CreateILockBytesOnHGlobal(hGlobal,true,LockBytes);
                  StgOpenStorageOnILockBytes(LockBytes,nil,STGM_SHARE_DENY_WRITE,0,0,Storage);
                  if Storage=nil then
                    error('Не удалось открыть хранилище IStorage, возможно данные повреждены');
                  Elements[0].TreeViewHandle:=TVI_ROOT;
                  ReadStorage(Storage,0);
                  Storage._Release;
                  pointer(Storage):=0;
                  LockBytes._Release;
                  pointer(LockBytes):=0;
                  GlobalFree(hGlobal);

                ///////////////////////////////////////////////////////////////
                //dir stream parsing
                ///////////////////////////////////////////////////////////////
                  k:=Decompress(Dir.StreamData,data,Dir.DataSize);
                  Dir.References:=HeapAlloc(Heap,0,k);
                  n:=0;
                  i:=0;
                  repeat
                    size:=pdword(@data[i+2])^;
                    case pword(@data[i])^ of
                            ID_PROJECTSYSKIND:begin
                                              Dir.SysKind:=pdword(@data[i+6])^;
                                              size:=Dir.SysKind;
                                              if size>3 then
                                                size:=4;
                                              Dir.strSysKind:=SysKind[size];
                                              inc(i,10);
                                              end;
                               ID_PROJECTLCID:begin
                                              Dir.LCID:=pdword(@data[i+6])^;
                                              inc(i,10);
                                              end;
                      ID_PROJECTCOMPATVERSION,
                        ID_PROJECTHELPCONTEXT,
                           ID_PROJECTLIBFLAGS,
                         ID_MODULEHELPCONTEXT,
                         ID_PROJECTLCIDINVOKE:inc(i,10);
                           ID_PROJECTCODEPAGE:begin
                                              Dir.CodePage:=pword(@data[i+6])^;
                                              inc(i,8);
                                              end;
                               ID_PROJECTNAME:begin
                                              Dir.ProjectNameLen:=size;
                                              Dir.ProjectName[MultiByteToWideChar(Dir.CodePage,0,@data[i+6],size,Dir.ProjectName,length(Dir.ProjectName))]:=#0;
                                              SendMessageW(Dialog,WM_SETTEXT,0,LongInt(@Dir.ProjectName));
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
                                              Dir.VersionMinor:=pdword(@data[i+10])^;
                                              inc(i,12);
                                              end;
                          ID_PROJECTCONSTANTS:begin
                                              inc(i,size+8);              //skip ansi
                                              size:=pdword(@data[i])^;
                                              move(data[i+4],Dir.Constants,size);
                                              Dir.Constants[size shr 1]:=#0;
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
                                              inc(size,pdword(@data[i+size+8])^+14);  //SizeTwiddled
                                              inc(size,pdword(@data[i+size])^+6);     //SizeOfName
                                              inc(size,pdword(@data[i+size])^+6);     //SizeOfNameUnicode
                                              inc(size,pdword(@data[i+size])^+4);     //SizeExtended    
                                              Dir.Originals[Dir.OriginalsCount]:=HeapAlloc(Heap,0,size);
                                              move(data[i],Dir.Originals[Dir.OriginalsCount]^,size);
                                              inc(Dir.OriginalsCount);
                                              inc(i,size);
                                              end;
                       ID_REFERENCEREGISTERED:begin
                                              size:=pdword(@data[i+6])^;
                                              MultiByteToWideChar(Dir.CodePage,0,@data[i+10],size,@Dir.References[n],size);
                                              inc(n,size);
                                              pdword(@Dir.References[n])^:=$A000D;
                                              inc(n,2);
                                              inc(i,size+16);
                                              end;
                            ID_PROJECTMODULES,
                              ID_MODULECOOKIE,
                             ID_PROJECTCOOKIE,
                          ID_REFERENCEPROJECT:inc(i,size+6);
                                ID_MODULENAME:begin
                                              ModuleNameLen:=size+size;
                                              ModuleName[MultiByteToWideChar(Dir.CodePage,0,@data[i+6],size,@ModuleName,size+1)]:=#0;
                                              inc(i,size+6);
                                              end;
                         ID_MODULENAMEUNICODE:begin
                                              move(data[i+6],ModuleName,size);
                                              ModuleNameLen:=size;
                                              ModuleName[size shr 1]:=#0;
                                              inc(i,size+6);
                                              end;
                          ID_MODULESTREAMNAME:begin
                                              inc(i,size+8);              //skip ansi
                                              size:=pdword(@data[i])^ shr 1;
                                              j:=0;
                                              repeat
                                                inc(j);
                                                if j>=ElementsCount then
                                                  error('Не найден поток для модуля');
                                              until (Elements[j].StreamType<>STREAM_TYPE_STORAGE) and strcmpnW(Elements[j].StreamName,@data[i+4],size);
                                              Elements[j].ModuleName:=ModuleName;
                                              Elements[j].Name:=@Elements[j].ModuleName;
                                              Elements[j].ModuleNameLen:=ModuleNameLen;
                                              Elements[j].StreamType:=STREAM_TYPE_PACKED_TEXT;
                                              inc(i,size+size+4);
                                              end;
                              ID_MODULEOFFSET:begin
                                              Elements[j].TextOffset:=pdword(@data[i+6])^;
                                              inc(i,10);
                                              end;
                               ID_MODULETYPE1,
                               ID_MODULETYPE2:begin
                                              Elements[j].ModuleType:=data[i];
                                              inc(i,6);
                                              end;
                            ID_MODULEREADONLY:begin
                                              Elements[j].Flags:=Elements[j].Flags or MODULE_FLAG_READONLY;
                                              inc(i,6);
                                              end;
                             ID_MODULEPRIVATE:begin
                                              Elements[j].Flags:=Elements[j].Flags or MODULE_FLAG_PRIVATE;
                                              inc(i,6);
                                              end;
                          ID_MODULETERMINATOR,
                         ID_MODULESTERMINATOR:inc(i,6);
                     else
                       error('Неожиданный id внутри dir потока');
                     end;
                  until i>=k;
                  Dir.References[n]:=#0;
                  HeapRealloc(Heap,HEAP_REALLOC_IN_PLACE_ONLY,Dir.References,n+n+2);
                end
                else
                  error('Отсутствует сигнатура GMS');

                SendMessageW(Edit,EM_SETLIMITTEXT,high(integer),0);
                i:=ImageList_LoadImageW($400000,PWideChar(1),16,0,CLR_DEFAULT,IMAGE_BITMAP,0);
                SendMessageW(Toolbar,TB_SETIMAGELIST,0,i);
                SendMessageW(Toolbar,TB_BUTTONSTRUCTSIZE,sizeof(_TBBUTTON),0);
                SendMessageW(Toolbar,TB_ADDBUTTONSW,length(tbbuttons),LongInt(@tbbuttons));
                SendMessageW(Toolbar,TB_SETMAXTEXTROWS ,0,0);
                SetClassLongW(wnd,GCL_HICON,LoadIconW($400000,PWideChar(1)));
                StdCursor:=LoadCursorW(0,pointer(IDC_ARROW));
                HSizeCursor:=LoadCursorW(0,pointer(IDC_SIZEWE));
                DlgFunc(wnd,WM_SIZE,0,0);
                result:=1;
                end;
       WM_CLOSE:begin
                //System will automaticaly freee memory when exit process
                //HeapDestroy(Heap);
                //VirtualFree(pointer(data),0,MEM_RELEASE);
                pointer(data):=0;
                ExitProcess(0);
                end;
  end;
end;

begin
  Dir.PConstants:=@Dir.Constants;
  InitCommonControls;
  DialogBoxParamW($400000,PWideChar(1),0,@DlgFunc,0);
end.
