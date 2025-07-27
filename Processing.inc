type
  TModuleName =array[0..31] of WideChar;

const
  HEAP_REALLOC_IN_PLACE_ONLY=16;
  MODULE_FLAG_PROCEDURAL =1;
  MODULE_FLAG_READONLY   =2;
  MODULE_FLAG_PRIVATE    =4;
  MODULE_TYPE_UNKNOWN    =0;
  MODULE_TYPE_MODULE     =1;
  MODULE_TYPE_CLASS      =2;
  STREAM_TYPE_STORAGE    =0;
  STREAM_TYPE_DIR        =1;
  STREAM_TYPE_TEXT       =2;
  STREAM_TYPE_PACKED_TEXT=3;
  STREAM_TYPE_UNKNOWN    =4;
  _VBA_PROJECT:   array[0..31] of byte=($CC,$61,$FF,$FF,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00);
  VBAProjectDAta: array [0..131] of byte=($90,$01,$04,$00,$00,$00,$01,$00,$00,$00,$9A,$01,$04,$00,$00,$00,$01,$00,$00,$00,$A4,$01,$04,$00,$00,$00,$26,$00,$00,$00,$AE,$01,$4C,$00,$00,$00,$7B,$00,$46,$00,$42,$00,$46,$00,$34,$00,$33,$00,$30,$00,$30,$00,$46,$00,$2D,$00,$44,$00,$39,$00,$32,$00,$31,$00,$2D,$00,$31,$00,$31,$00,$44,$00,$31,$00,$2D,$00,$42,$00,$38,$00,$30,$00,$36,$00,$2D,$00,$30,$00,$30,$00,$41,$00,$30,$00,$43,$00,$39,$00,$30,$00,$36,$00,$34,$00,$36,$00,$41,$00,$39,$00,$7D,$00,$B8,$01,$04,$00,$00,$00,$00,$00,$00,$00,$CC,$01,$04,$00,$00,$00,$E3,$32,$06,$E5);

var
  data:                                array of byte;
  wdata:                               array of WideChar absolute data;
  TreeView,Edit,Toolbar,TreeMenu,Heap: THandle;
  datalen,SelectedItem,Project:        dword;
  ElementsCount:                       integer;
  InsertStruct:                        tagTVINSERTSTRUCTW=(hInsertAfter:TVI_LAST;item:(mask:TVIF_TEXT+TVIF_PARAM));
  Elements:                            array of packed record
                                                Name:           PWideChar;
                                                StreamName:     PWideChar;
                                                StreamData:     PAnsiChar;
                                                DataSize:       dword;
                                                Flags:          dword;
                                                TextOffset:     dword;
                                                TreeViewHandle: HTREEITEM;
                                                StreamType:     byte;
                                                ModuleType:     byte;
                                                Parent:         word;
                                                ModuleNameLen:  dword;
                                                ModuleName:     TModuleName;
                                                end;
  Dir:                                 packed record
                                       PConstants:    PWideChar;
                                       strSysKind:    PWideChar;
                                       SysKind:       dword;                                       
                                       VersionMajor:  dword;
                                       VersionMinor:  dword;
                                       CodePage:      dword;
                                       LCID:          dword;
                                       References:    PWideChar;
                                       StreamData:    PAnsiChar;
                                       DataSize:      dword;
                                       ElementIndex:  dword;
                                       ProjectNameLen:dword;
                                       OriginalsCount:dword;
                                       ProjectName:   array[0..128] of WideChar;
                                       Constants:     array[0..1015] of WideChar;
                                       Originals:     array[0..63] of pointer;
                                       end;

function strcmpnW(str1,str2: PWideChar; len: integer): boolean;
begin
  result:=false;
  if (str1=nil)or(len=0) then
    exit;
  repeat
    dec(len);
    if str1[len]<>str2[len] then
      exit;
  until len=0;
  result:=true;
end;

function strcpyW(OutStr,InStr: PWideChar): dword;
begin
  result:=0;
  repeat
    Outstr[result]:=Instr[result];
    inc(result);
  until Instr[result-1]=#0;
  dec(result);
end;

procedure DeleteItem(index: integer);
var
  i,j,k,shift: integer;
  itemhandle:  THandle;
begin
  j:=index;
  itemhandle:=THandle(Elements[j].TreeViewHandle);
  SendMessageW(Edit,EM_SETMODIFY,0,0);
  repeat
   inc(j);
  until (Elements[j].Parent<=Elements[index].parent)or(j>=ElementsCount);
  for i:=index to j-1 do
  begin
    CoTaskMemFree(Elements[i].StreamName);
    HeapFree(Heap,0,Elements[i].StreamData);
  end;
  i:=index;
  shift:=j-index;
  if Dir.ElementIndex>i then
    dec(Dir.ElementIndex,shift);
  if Project>i then
    dec(Project,shift);
  for k:=j to ElementsCount-1 do
  begin
    if Elements[k].Parent>i then
      dec(Elements[k].Parent,shift);
    Elements[k-shift]:=Elements[k];
    if Elements[k].Name=@Elements[k].ModuleName then
      Elements[k-shift].Name:=@Elements[k-shift].ModuleName;
  end;
  dec(ElementsCount,shift);
  InsertStruct.item.mask:=TVIF_PARAM+TVIF_HANDLE;
  for i:=ElementsCount-1 downto 0 do
  begin
    InsertStruct.item.hItem:=Elements[i].TreeViewHandle;
    InsertStruct.item.lParam:=i;
    SendMessageW(TreeView,TVM_SETITEMW,0,LongInt(@InsertStruct.item));
  end;
  InsertStruct.item.mask:=TVIF_PARAM+TVIF_TEXT;
  SendMessageW(TreeView,TVM_DELETEITEM,0,itemhandle);
end;

function ProcessingDlgFunc(wnd,msg,wParam,lParam: dword):dword;stdcall;
label
  NoModule;
var
  i,j,k,size: LongInt;
  tmp:        PWideChar;
  IsBase:     boolean;
begin
  result:=0;
  case msg of
     WM_COMMAND:if loword(wParam)=2 then
                begin
                  i:=SendMessageW(lParam,BM_GETCHECK,0,0);
                  EnableWindow(GetDlgItem(wnd,3),LongBool(i));
                  EnableWindow(GetDlgItem(wnd,4),LongBool(i));
                end
                else if wParam=0 then
                begin
////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Rebuilding PROJECT stream
////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if SendDlgItemMessageW(wnd,1,BM_GETSTATE,0,0) and BST_CHECKED>0 then
                begin
                  k:=strcpyW(@wdata[0],'ID="{00000000-0000-0000-0000-000000000000}"'#13#10);
                  tmp:=VirtualAlloc(0,datalen*3,MEM_COMMIT,PAGE_READWRITE);
                  for i:=0 to ElementsCount-1 do
                  begin
                    IsBase:=false;
                    with Elements[i] do
                      if StreamType=STREAM_TYPE_PACKED_TEXT then
                      begin
                        size:=Decompress(@StreamData[TextOffset],@tmp[datalen],DataSize-TextOffset);
                        MultiByteToWideChar(Dir.CodePage,0,@tmp[datalen],size,tmp,size);
                        j:=0;
                        ModuleType:=ID_MODULETYPE2;
                        while j<size-28 do
                        begin
                          if strcmpnW(@tmp[j],'Attribute VB_Base = "0{',23) then
                          begin
                            inc(j,23);
                            repeat
                              inc(j);
                            until (tmp[j]='}')or(tmp[j]<=#13);
                            if pdword(@tmp[j])^=$7B007D then //'}{'
                            begin
                              inc(k,strcpyW(@wdata[k],'BaseClass='));
                              inc(k,strcpyW(@wdata[k],Name));
                              pdword(@wdata[k])^:=$000A000D;
                              inc(k,2);
                              goto NoModule;
                            end;
                            IsBase:=true;
                          end
                          else if IsBase and (strcmpnW(@tmp[j],'Attribute VB_Customizable = "0{',28)) then
                          begin
                            inc(j,28);
                            if (tmp[j]='t')or(tmp[j]='T') then
                            begin
                              inc(k,strcpyW(@wdata[k],'Document='));
                              inc(k,strcpyW(@wdata[k],Name));
                              inc(k,strcpyW(@wdata[k],'/&H00000000'#13#10));
                              goto NoModule;
                            end
                            else
                            begin
                              inc(k,strcpyW(@wdata[k],'Class='));
                              inc(k,strcpyW(@wdata[k],Name));
                              pdword(@wdata[k])^:=$000A000D;
                              inc(k,2);
                              goto NoModule;
                            end
                          end;
                          while j<size-28 do
                          begin
                            if pdword(@tmp[j])^=$000A000D then
                              break;
                            inc(j);
                          end;
                          inc(j,2);
                        end;
                        ModuleType:=ID_MODULETYPE1;
                        inc(k,strcpyW(@wdata[k],'Module='));
                        inc(k,strcpyW(@wdata[k],Name));
                        pdword(@wdata[k])^:=$000A000D;
                        inc(k,2);
                        NoModule:
                      end;
                  end;    
                  VirtualFree(tmp,0,MEM_RELEASE);
                  inc(k,strcpyW(@wdata[k],'Name="'));
                  inc(k,strcpyW(@wdata[k],@Dir.ProjectName));
                  inc(k,strcpyW(@wdata[k],'"'#13#10'HelpContextID="0"'#13#10'VersionCompatible32="393222000"'#13#10#13#10));
                  HeapFree(Heap,0,Elements[Project].StreamData);
                  Elements[Project].StreamData:=HeapAlloc(Heap,0,k);
                  Elements[Project].DataSize:=k;
                  WideCharToMultiByte(Dir.CodePage,0,@wdata[0],k,Elements[Project].StreamData,k+1,0,0);
                  if SelectedItem=Project then
                  begin
                    SendMessageW(Edit,WM_SETTEXT,0,LongInt(data));
                    SendMessageW(Edit,EM_SETMODIFY,0,0);
                  end;
                end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Rebuilding dir stream
////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if SendDlgItemMessageW(wnd,2,BM_GETSTATE,0,0) and BST_CHECKED>0 then
                begin
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //Deleting P-code
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////
                  if SendDlgItemMessageW(wnd,3,BM_GETSTATE,0,0) and BST_CHECKED>0 then
                  begin
                    i:=0;
                    while i<ElementsCount do
                    begin
                      with Elements[i] do
                        if StreamType=STREAM_TYPE_PACKED_TEXT then
                        begin
                          dec(DataSize,TextOffset);
                          move(StreamData[TextOffset],StreamData[0],DataSize);
                          TextOffset:=0;
                          StreamData:=HeapRealloc(Heap,HEAP_REALLOC_IN_PLACE_ONLY,StreamData,DataSize);
                        end
                        else if lstrcmpW(StreamName,'_VBA_PROJECT')=0 then
                        begin
                          DataSize:=sizeof(_VBA_PROJECT);
                          StreamData:=HeapRealloc(Heap,0,StreamData,sizeof(_VBA_PROJECT));
                          move(_VBA_PROJECT,StreamData[0],sizeof(_VBA_PROJECT));
                        end
                        else if strcmpnW(StreamName,'__SRP_',6) then
                        begin
                          DeleteItem(i);
                          dec(i);
                        end;
                        inc(i);
                    end;
                  end;
                  pword(@data[0])^  :=ID_PROJECTCODEPAGE;
                  pdword(@data[2])^ :=4;
                  pword(@data[6])^  :=Dir.CodePage;
                  pword(@data[8])^  :=ID_PROJECTNAME;
                  pdword(@data[10])^:=Dir.ProjectNameLen;
                  k:=WideCharToMultiByte(Dir.CodePage,0,@Dir.ProjectName,Dir.ProjectNameLen,@data[14],Dir.ProjectNameLen+1,0,0)+14;
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //References
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////
                  if SendDlgItemMessageW(wnd,4,BM_GETSTATE,0,0) and BST_CHECKED=0 then
                  begin
                    for i:=Dir.OriginalsCount-1 downto 0 do
                    begin
                      j:=HeapSize(Heap,0,Dir.Originals[i]);
                      move(Dir.Originals[i]^,data[k],j);
                      inc(k,j);
                    end;
                    i:=0;
                    while Dir.References[i]<>#0 do
                    begin
                      pword(@data[k])^:=ID_REFERENCEREGISTERED;
                      j:=0;
                      while Dir.References[i+j]>#13 do
                        inc(j);
                      WideCharToMultiByte(Dir.CodePage,0,@Dir.References[i],j,@data[k+10],j,0,0);
                      pdword(@data[k+2])^:=j+10;
                      pdword(@data[k+6])^:=j;
                      inc(k,j+16);
                      pdword(@data[k-6])^:=0;
                      pword(@data[k-2])^ :=0;
                      if Dir.References[i]=#0 then
                        break;
                      inc(i,j+2);
                    end;
                  end;
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //Modules
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////
                  for i:=0 to ElementsCount-1 do
                    with Elements[i] do
                      if StreamType=STREAM_TYPE_PACKED_TEXT then
                      begin
                        pword(@data[k])^    :=ID_MODULENAME;
                        j:=ModuleNameLen shr 1;
                        pdword(@data[k+2])^ :=j;
                        inc(k,WideCharToMultiByte(Dir.CodePage,0,@Name,j,@data[k+6],j+1,0,0)+6);
                        pword(@data[k])^    :=ID_MODULENAMEUNICODE;
                        pdword(@data[k+2])^ :=ModuleNameLen;
                        move(Name^,data[k+6],ModuleNameLen);
                        inc(k,ModuleNameLen+6);
                        pword(@data[k])^    :=ID_MODULESTREAMNAME;
                        j:=lstrlenW(StreamName);
                        pdword(@data[k+2])^ :=j;
                        inc(k,WideCharToMultiByte(Dir.CodePage,0,@StreamName,j,@data[k+6],j+1,0,0)+6);
                        inc(j,j);
                        pword(@data[k])^    :=ID_MODULESTREAMNAMEUNICODE;
                        pdword(@data[k+2])^ :=j;
                        move(StreamName^,data[k+6],j);
                        inc(k,j+6);
                        pword(@data[k])^    :=ID_MODULEOFFSET;
                        pdword(@data[k+2])^ :=4;
                        pdword(@data[k+6])^ :=TextOffset;
                        pword(@data[k+10])^ :=ModuleType;
                        pdword(@data[k+12])^:=0;
                        pword(@data[k+16])^ :=ID_MODULETERMINATOR;
                        pdword(@data[k+18])^:=0;
                        inc(k,22);
                      end;
                  pword(@data[k])^   :=ID_MODULESTERMINATOR;
                  pdword(@data[k+2])^:=0;
                  inc(k,6);
                  HeapFree(Heap,0,Dir.StreamData);
                  Dir.StreamData:=HeapAlloc(Heap,0,datalen);
                  k:=Compress(data,Dir.StreamData,k);
                  Dir.StreamData:=HeapRealloc(Heap,HEAP_REALLOC_IN_PLACE_ONLY,Dir.StreamData,k);
                  Dir.DataSize:=k;
                  Elements[Dir.ElementIndex].StreamData:=Dir.StreamData;
                  Elements[Dir.ElementIndex].DataSize:=k;
                end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Downgrade typelib version
////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if SendDlgItemMessageW(wnd,5,BM_GETSTATE,0,0) and BST_CHECKED>0 then
                begin
                  for i:=0 to ElementsCount-1 do
                    with Elements[i] do
                      if lstrcmpW(StreamName,'VBA Project Data')=0 then
                      begin
                        DataSize:=sizeof(VBAProjectData);
                        StreamData:=HeapRealloc(Heap,0,StreamData,sizeof(VBAProjectData));
                        move(VBAProjectData,StreamData[0],sizeof(VBAProjectData));
                      end;
                end;
                EndDialog(wnd,0);
                end;
  WM_INITDIALOG:result:=1;
       WM_CLOSE:EndDialog(wnd,0);
  end;
end;
