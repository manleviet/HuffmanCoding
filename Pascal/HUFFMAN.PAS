{Chuong trinh lap ma Compact cho bo chu tieng Viet
 Tac gia: Le Viet Man
 Lop: Tin K24B
 Truong: DHKH - Hue
 Email: manleviet@yahoo.com}
Program HuffmanCoding;
Const FI='HUFFMAN.INP';
      FO='HUFFMAN.OUT';
Type
    RHuffman=Record
                   ID: Integer;
                   Char:String[3];
                   HCode:String[40];
                   R,L,B:Integer;
                   p:Real;
             End;
    THuffman=Array[1..32,1..33] of RHuffman;
Var Huffman:THuffman;

Procedure SwapH(Var a, b:RHuffman);
Var t:RHuffman;
Begin
     t := a;
     a := b;
     b := t;
End;

Procedure SortByp(Var Huffman:THuffman; t:Integer);
Var i,j:Integer;
Begin
     For i:=1 to 33-t do
         For j:=i+1 to 34-t do
             If Huffman[t,i].p < Huffman[t,j].p then
                SwapH(Huffman[t,i],Huffman[t,j]);
End;

Procedure FixSource(Var Huffman:THuffman);
Var i,j:Integer;
Begin
     SortByp(Huffman,1);
     For i := 1 to 31 do
         Begin
              For j:=1 to 32-i do
                  Begin
                       Huffman[i+1,j].ID:=j;
                       Huffman[i+1,j].B:=j;
                       Huffman[i+1,j].p:=Huffman[i,j].p;
                  End;
              Huffman[i+1,j+1].ID:=j+1;
              Huffman[i+1,j+1].R:=j+1;
              Huffman[i+1,j+1].L:=j+2;
              Huffman[i+1,j+1].p:=Huffman[i,j+1].p + Huffman[i,j+2].p;
              SortByp(Huffman,i+1);
         End;
End;

Procedure HCoding(Var Huffman:THuffman);
Var i,j:Integer;
Begin
     Huffman[32,1].HCode:='0';
     Huffman[32,2].HCode:='1';
     For i:=32 downto 2 do
         For j:=1 to 34-i do
             If Huffman[i,j].B <> 0 then
                Huffman[i-1,Huffman[i,j].B].HCode := Huffman[i,j].HCode
             Else
                 Begin
                      Huffman[i-1,Huffman[i,j].R].HCode := Huffman[i,j].HCode + '0';
                      Huffman[i-1,Huffman[i,j].L].HCode := Huffman[i,j].HCode + '1';
                 End;
End;

Procedure ReadInfor(Var Huffman:THuffman);
{Doc du lieu tu tap tin Huffman.INP va khoi tao mang Huffman}
Var F:Text;
    i:Integer;
Begin
     Assign(F,FI);
     Reset(F);
     For i:=1 to 33 do
         Begin
              Read(F,Huffman[1,i].ID);
              Read(F,Huffman[1,i].Char);
              Read(F,Huffman[1,i].p);
              Readln(F);
         End;
     Close(F);
End;

Procedure SortByID(Var Huffman:THuffman; t:Integer);
Var i,j:Integer;
Begin
     For i:=1 to 33-t do
         For j:=i+1 to 34-t do
             If Huffman[t,i].ID > Huffman[t,j].ID then
                SwapH(Huffman[t,i],Huffman[t,j]);
End;

{Procedure WriteHuffman(Huffman:THuffman);
Var F:Text;
    i,j:Integer;
Begin
     Assign(F,FO);
     Rewrite(F);
     For i:=1 to 32 do
         Begin
              Writeln(F,'Huffman ',i);
              For j:=1 to 34-i do
                  Begin
                       Write(F,Huffman[i,j].ID,' ');
                       Write(F,Huffman[i,j].Char,' ');
                       Write(F,Huffman[i,j].R,' ');
                       Write(F,Huffman[i,j].L,' ');
                       Write(F,Huffman[i,j].B,' ');
                       Write(F,Huffman[i,j].HCode,' ');
                       Write(F,Huffman[i,j].p:0:2);
                       Writeln(F);
                  End;
         End;
     Close(F);
End;}

Procedure WriteInfor(Huffman:THuffman);
Var F:Text;
    i:Integer;
    t:Real;
Begin
     Assign(F,FO);
     Rewrite(F);
     SortByID(Huffman,1);
     t:=0;
     For i:=1 to 33 do
         Begin
              t:= t + (Huffman[1,i].p * Length(Huffman[1,i].HCode));
              Write(Huffman[1,i].ID:5);
              Write(Huffman[1,i].Char:4);
              Write(Huffman[1,i].HCode);
              Writeln;
              Write(F,Huffman[1,i].ID:5);
              Write(F,Huffman[1,i].Char:4);
              Write(F,Huffman[1,i].HCode);
              Writeln(F);
         End;
     Write('L(a)=',t:0:6);
     Write(F,'L(a)=',t:0:6);
     Close(F);
End;

Begin
     ReadInfor(Huffman);
     FixSource(Huffman);
     HCoding(Huffman);
     WriteInfor(Huffman);
{     WriteHuffman(Huffman);}
     Readln;
End.