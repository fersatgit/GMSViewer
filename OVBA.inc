var
  Dialog: THandle;

const
  ID_PROJECTSYSKIND         =1;
  ID_PROJECTCOMPATVERSION   =$4A;
  ID_PROJECTLCID            =2;
  ID_PROJECTLCIDINVOKE      =$14;
  ID_PROJECTCODEPAGE        =3;
  ID_PROJECTNAME            =4;
  ID_PROJECTDOCSTRING       =5;
  ID_PROJECTDOCSTRINGUNICODE=$40;
  ID_PROJECTHELPFILEPATH    =6;
  ID_PROJECTHELPFILEPATH2   =$3D;
  ID_PROJECTHELPCONTEXT     =7;
  ID_PROJECTLIBFLAGS        =8;
  ID_PROJECTVERSION         =9;
  ID_PROJECTCONSTANTS       =$C;
  ID_REFERENCENAME          =$16;
  ID_REFERENCECONTROL       =$2F;
  ID_REFERENCEORIGINAL      =$33;
  ID_REFERENCEREGISTERED    =$D;
  ID_REFERENCEPROJECT       =$E;
  ID_PROJECTMODULES         =$F;
  ID_PROJECTCOOKIE          =$13;
  ID_MODULENAME             =$19;
  ID_MODULENAMEUNICODE      =$47;
  ID_MODULESTREAMNAME       =$1A;
  ID_MODULESTREAMNAMEUNICODE=$32;
  ID_MODULEDOCSTRING        =$1C;
  ID_MODULEDOCSTRINGUNICODE =$48;
  ID_MODULEOFFSET           =$31;
  ID_MODULEHELPCONTEXT      =$1E;
  ID_MODULECOOKIE           =$2C;
  ID_MODULETYPE             =$21; //or $22
  ID_MODULETYPE1            =$21;
  ID_MODULETYPE2            =$22;
  ID_MODULEREADONLY         =$25;
  ID_MODULEPRIVATE          =$28;
  ID_MODULETERMINATOR       =$2B;
  ID_MODULESTERMINATOR      =$10;
  LengthMask: array[0..8] of dword=(15,31,63,127,255,511,1023,2047,4095);

procedure error(text: PWideChar);
begin
  MessageBoxW(Dialog,text,0,0);
  ExitProcess(0);
end;

function calcbitcount(a: dword): dword;
asm
  dec   eax
  bsr   eax,eax
  movzx eax,byte[@bitcount+eax+1]
  ret
  @bitcount: db 12,12,12,12,12,11,10,9,8,7,6,5,4
end;

function Compress(DecompressedData,CompressedData: pointer; datasize: integer): dword;
var
  compressed_chunk_start,decompressed_end,decompressed_chunk_start: integer;
  candidate,best,L,i,j,k:    integer;
  length,bitcount,maxlength: integer;
  flagindex,FlagByte,bit:    dword;
  compdata: array of byte absolute CompressedData;
  data:     array of byte absolute DecompressedData;
begin
  i:=0;
  j:=1;
  compdata[0]:=1;
  repeat
    compressed_chunk_start  :=j;
    decompressed_chunk_start:=i;
    decompressed_end        :=i+4096;
    inc(j,2);
		if decompressed_end>datasize then
			decompressed_end:=datasize;
		while i<decompressed_end do
    begin
      FlagIndex:=j;
      FlagByte :=0;
      bit      :=1;
      inc(j);
      for k:=7 downto 0 do
      begin
        if i>=decompressed_end then
          break;
        candidate:=i-1;
        length   :=0;
        while candidate>=decompressed_chunk_start do
        begin
          L:=0;
          while (i+L<decompressed_end)and(data[i+L]=data[candidate+L]) do
            inc(L);
          if L>length then
          begin
            length:=L;
            best  :=candidate;
          end;
          dec(candidate);
        end;
        dec(length,3);
        if length>=0 then
        begin
          bitcount:=calcbitcount(i-decompressed_chunk_start);
          maxlength:=LengthMask[bitcount-4];
          if maxlength<length then
            length:=maxlength;
          pword(@compdata[j])^:=((i-best-1) shl bitcount)+length;
          inc(FlagByte,bit);
          inc(j,2);
          inc(i,length+3);
        end
        else
        begin
          compdata[j]:=data[i];
          inc(j);
          inc(i);
        end;
        inc(bit,bit);
      end;
      CompData[FlagIndex]:=FlagByte;
    end;
    k:=j-compressed_chunk_start-3;
    if k<0 then
      k:=0;
    pword(@compdata[compressed_chunk_start])^:=$B000+k;
  until i>=datasize;
  result:=j;
end;

function Decompress(CompressedData, DecompressedData: pointer; compdatasize: integer): dword;
label
  finish;
var
 header:                            word;
 i,j,k,token,FlagByte,bitcount:     integer;
 Length,Offset,chunkstart,chunkend: integer;
 compdata: array of byte absolute CompressedData;
 data:     array of byte absolute DecompressedData;
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
        move(compdata[i+2],data[chunkstart],4096);
        i:=chunkend;
        inc(chunkstart,4096);
      end
      else
      begin
        inc(i,2);
        j:=0;
        while i<chunkend do
        begin
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
        end;
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
