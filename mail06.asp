<%
Option Explicit
	Server.ScriptTimeout = 3000
%>
<!-- #include file="../adovbs.inc" -->
<%

Dim strSQL
Dim rsFechas, rsJugadores
'Dim rsRanking1, rsRanking2, rsRanking3, rsRanking5, rsRanking6, rsRanking7,  rsRanking9, rsRanking10, rsRanking11, rsRanking13, rsRanking14, rsRanking15, rsRanking17, rsRanking18, rsRanking19
Dim rsFecha, rsJugador, rsScores, rsScores2
Dim rsMiercoles, rsJueves, rsViernes, rsDomingos
Dim idCancha
Dim Fecha,idScore,eMails
Dim idRanking
Dim cmdDC, rsTarjeta, rsMisc, rsPromedio, rsResumen , rsScore
Dim Item
Dim i,j
Dim Yardas,Par,Score
Dim YardasIda,ParIda,ScoreIda
Dim YardasVta,ParVta,ScoreVta
Dim Anterior,Longitud,Linea
Dim Mail
Dim tx
Dim rsLinea
Dim G,R,Q,K,P,X,Y
Dim objFile


        ' AQUI TENES QUE DEFINIR LAS CATEGORIAS
        
      ' Declarando un array para 3 Categorias (SI SE PONE 2 PARA DEFINIR 3).
      Dim categoriaHCPDesde(2) ' As Integer
      Dim categoriaHCPHasta(2) ' As Integer
      Dim categoriaDescripcion(2) ' As String

      'Categoria 1
      categoriaHCPDesde(0)=0
      categoriaHCPHasta(0)=12
      categoriaDescripcion(0)= "Categoria Hasta 12"
      'Categoria 2
      categoriaHCPDesde(1)=13
      categoriaHCPHasta(1)=19
      categoriaDescripcion(1)= "Categoria 13 a 19"
      'Categoria 3
      categoriaHCPDesde(2)=20
      categoriaHCPHasta(2)=54
      categoriaDescripcion(2)= "Categoria 20 a 54"



Set objFile = Server.CreateObject("Scripting.FileSystemObject")

Fecha = Trim(Request.Form("Fecha"))
eMails = Trim(Request.Form("eMails"))
G = Trim(Request.Form("G"))
R = Trim(Request.Form("R"))
Q = Trim(Request.Form("Q"))
K = Trim(Request.Form("K"))
X = Trim(Request.Form("X"))
Y = Trim(Request.Form("Y"))
P = Trim(Request.Form("P"))
%>

<!--#include file="../connopen.asp"-->
<!--#include file="sec.asp"-->
<html>
<body>

<% if Fecha <> "" and eMails<>"" Then


  Set cmdDC = Server.CreateObject("ADODB.Command")
  cmdDC.ActiveConnection = DataConn


cmdDC.CommandText = "select * from fechas " &_
"where fecha='" & Fecha & "'"
cmdDC.CommandType = 1
Set rsFecha = Server.CreateObject("ADODB.Recordset")
rsFecha.Open cmdDC, , 0, 1

if G="ON" then
   cmdDC.CommandText = "sp_scores3 '" & Fecha & "'"
   cmdDC.CommandType = 1

   Set rsScores = Server.CreateObject("ADODB.Recordset")
   rsScores.Open cmdDC, , 0, 1
   Set rsScores2 = Server.CreateObject("ADODB.Recordset")
   rsScores2.Open cmdDC, , 0, 1
end if


if P="ON" then
   cmdDC.CommandText = "select Top 2 Fecha,Nombre from Fechas,canchas "&_
   "where Fechas.idCancha=Canchas.idCancha "&_
   "and left(Fecha,1)='m' "&_
   "and right(Fecha,5) > '" & right(Fecha,5) & "' "&_
   "order by Fecha asc"
   cmdDC.CommandType = 1
   Set rsMiercoles = Server.CreateObject("ADODB.Recordset")
   rsMiercoles.Open cmdDC, , 0, 1

   cmdDC.CommandText = "select Top 2 Fecha,Nombre from Fechas,canchas "&_
   "where Fechas.idCancha=Canchas.idCancha "&_
   "and left(Fecha,1)='j' "&_
   "and right(Fecha,5) > '" & right(Fecha,5) & "' "&_
   "order by Fecha asc"
   cmdDC.CommandType = 1
   Set rsJueves = Server.CreateObject("ADODB.Recordset")
   rsJueves.Open cmdDC, , 0, 1

   cmdDC.CommandText = "select Top 2 Fecha,Nombre from Fechas,canchas "&_
   "where Fechas.idCancha=Canchas.idCancha "&_
   "and left(Fecha,1)='v' "&_
   "and right(Fecha,5) > '" & right(Fecha,5) & "' "&_
   "order by Fecha asc"
   cmdDC.CommandType = 1
   Set rsViernes = Server.CreateObject("ADODB.Recordset")
   rsViernes.Open cmdDC, , 0, 1

   cmdDC.CommandText = "select Top 2 Fecha,Nombre from Fechas,canchas "&_
   "where Fechas.idCancha=Canchas.idCancha "&_
   "and left(Fecha,1)='d' "&_
   "and right(Fecha,5) > '" & right(Fecha,5) & "' "&_
   "order by Fecha asc"
   cmdDC.CommandType = 1
   Set rsDomingos = Server.CreateObject("ADODB.Recordset")
   rsDomingos.Open cmdDC, , 0, 1
end if

  Anterior=1
  if asc(right(eMails,1))<>10 then
    eMails=eMails+chr(13)+chr(10)
  end if
  Longitud=len(eMails)
  for j= 1 to Longitud-1
    if asc(mid(eMails,j,1))=13 and asc(mid(eMails,j+1,1))=10  then
      Linea=mid(eMails,Anterior,j-Anterior)
      Anterior=j+2 
      enviarmail(Linea) %>
<p><font face="Verdana">***<%= left(Linea,instr(Linea,",")-1) %>***<%= Right(Linea,Len(Linea)-instr(Linea,",")) %>&nbsp;</font></p>
    <% end if
  next
else %>
	<form action="mail06.asp" method="post" name="formulario">
		<p align="center"><b><u>Mail Tarjeta</u></b></p>
		<% strSQL = "SELECT Fecha,Nombre FROM fechas,canchas where fechas.idcancha=canchas.idcancha ORDER BY dia desc"
		'right(fecha,5) 
		Set rsFechas = Server.CreateObject("ADODB.Recordset")
		rsFechas.Open strSQL, DataConn, adOpenForwardOnly, adLockOptimistic, adCmdText

		If Not rsFechas.EOF Then
			rsFechas.MoveFirst	%>
			<b>Fecha:</b> <select name="Fecha">
				<option></option>
			<% Do While Not rsFechas.EOF
				Response.Write "<option value="""
				Response.Write rsFechas.Fields("Fecha")
				Response.Write """"
				Response.Write ">"
				Response.Write rsFechas.Fields("Fecha") & "  -  " & Trim(rsFechas.Fields("Nombre"))
				Response.Write "</option>" & vbCrLf

				rsFechas.MoveNext
			Loop %>
			</select>
		<% End If
		rsFechas.Close
		Set rsFechas =  Nothing%>
  		<p><input type="checkbox" name="G" value="ON" checked>Ganadores</p>
	  <p><input type="checkbox" name="R" value="ON" checked>Ranking 2023</p>
      <p>
        <input type="checkbox" name="Q" value="ON" checked>
        Ranking Toso J-A</p>
      <p>
        <input type="checkbox" name="K" value="ON" checked>
        Ranking Toso S-O</p>
      <!-- <p><input type="checkbox" name="Q" value="ON" checked>Ranking2 MIERCOLES y VIERNES</p>
		  <p><input type="checkbox" name="K" value="ON" checked>Ranking3 FIN DE SEMANA y FERIADO</p> -->
      
      <!-- 	<p><input type="checkbox" name="X" value="ON" checked>Ranking4 CORDOBA</p>
		  <p><input type="checkbox" name="Y" value="ON" checked>Ranking5 PARAGUAY</p> -->
		<p></p>
        <textarea rows="14" name="emails" cols="109"></textarea><p></p>
		<input type="submit" value="Submit" /> </p>
	</form>
<% end if
DataConn.Close
Set DataConn = Nothing
%>
</body>
</html>

<%


'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------


sub tarjeta()
if not rsTarjeta.eof then
  tx = tx & "<table width=""430"" height=""470"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
  tx = tx & "  <tr>"
  tx = tx & "    <td valign=""top"">"
  TablaTitulos
  TablaMitad(1)
  YardasIda=Yardas
  ParIda=Par
  ScoreIda=Score
  TablaMitad(10)
  YardasVta=Yardas
  ParVta=Par
  ScoreVta=Score
  TablaInferior
  tx = tx & "    </td>"
  tx = tx & "  </tr>"
  tx = tx & "</table>"
end if
end sub

'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub TablaTitulos()
tx = tx & "  <table width=""100%"" height=""64"" border=""0"" cellpadding=""0"" cellspacing=""6"" background=""http://www.golfguide.com.ar/db/images/" & trim(rsMisc.fields("NombreC")) & "_top.jpg"">"
tx = tx & "    <tr>"
tx = tx & "      <td valign=""top"">"
                   TablaTitulos2
tx = tx & "      </td>"
tx = tx & "    </tr>"
tx = tx & "  </table>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub TablaInferior()
tx = tx & "  <table width=""100%"" height=""205"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""FONT-SIZE: 10px; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; BORDER-COLLAPSE: collapse"">"
tx = tx & "    <tr valign=""top"">"
tx = tx & "      <td width=""205"">"
                   TablaIzquierda
tx = tx & "      </td>"
tx = tx & "      <td width=""20"">&nbsp;</td>"
tx = tx & "      <td  align=""center"" width=""205"">"
                   TablaDerecha 
tx = tx & "      </td>"
tx = tx & "    </tr>"
tx = tx & "  </table>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub TablaDerecha()
tx = tx & "  <table width=""100%"" height=""206"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""FONT-SIZE: 14px; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; BORDER-COLLAPSE: collapse"">"
tx = tx & "    <tr>"
tx = tx & "      <td align=""center"">"
             TablaPromedios 
tx = tx & "      </td>"
tx = tx & "    </tr>"
tx = tx & "    <tr>"
tx = tx & "      <td>"
             TablaResumen 
tx = tx & "      </td>"
tx = tx & "    </tr>"
tx = tx & "    <tr>"
             if idRanking = "" then
tx = tx & "      <td align=""center"">"
tx = tx & "        &nbsp;"
tx = tx & "        <a href=""http://www.golfguide.com.ar/db/stats.asp?Matricula=" & rsMisc.fields("Matricula") & """><b>CLICK AQUI PARA VER MAS ESTADISTICAS SUYAS EN GOLF GUIDE</b></a>"
tx = tx & "      </td>"
'             else
'tx = tx & "      <td align=""center"">"
'tx = tx & "        <a href=""http://www.golfguide.com.ar/db/stats.asp?Matricula=" & rsMisc.fields("Matricula") & "&idRanking=" & idRanking & """><b>ESTADISTICA RANKING</b></a>"
'tx = tx & "      </td>"
             end if
tx = tx & "    </tr>"
tx = tx & "  </table>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub TablaPromedios()
tx = tx & "  <TABLE id=""Promedios"" style=""FONT-SIZE: 12px; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; BORDER-COLLAPSE: collapse"" borderColor=""#c0c0c0"" cellSpacing=""0"" width=""205"" border=""1"">"
tx = tx & "    <TR>"
tx = tx & "      <TD align=""center"" bgColor=""#000000"" colSpan=""2"">"
tx = tx & "        <FONT color=""#ffffff""><B>Promedios</B></FONT>"
tx = tx & "      </TD>"
tx = tx & "    </TR>"
             Do While Not rsPromedio.EOF
tx = tx & "      <tr>"
tx = tx & "        <td align=""center"" width=""80%""><b>Par" & rsPromedio.Fields(0) & "s&nbsp;</b></td>"
tx = tx & "        <td align=""center"" width=""20%"">" & int(rsPromedio.Fields(1)*100)/100 & "&nbsp;</td>"
tx = tx & "      </tr>"
               rsPromedio.MoveNext
             Loop
tx = tx & "  </TABLE>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub TablaResumen()
tx = tx & "  <table style=""FONT-SIZE: 12px; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; BORDER-COLLAPSE: collapse"" bordercolor=""#c0c0c0"" cellspacing=""0"" width=""205"" border=""1"">"
tx = tx & "    <tr>"
tx = tx & "      <td align=""center"" bgcolor=""#000000"" colspan=""2"">"
tx = tx & "        <font color=""#ffffff""><b>Resumen</b></font>"
tx = tx & "      </td>"
tx = tx & "    </tr>"
             Do While Not rsResumen.EOF
               if rsResumen.Fields(2) <> 0 then
tx = tx & "    <tr>"
tx = tx & "      <td align=""center"" width=""80%""><b>" & rsResumen.Fields(1) & "&nbsp;</b></td>"
tx = tx & "      <td align=""center"" width=""20%"">" & rsResumen.Fields(2) & "&nbsp;</td>"
tx = tx & "    </tr>"
               end if
               rsResumen.MoveNext
             Loop
tx = tx & "  </table>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub TablaIzquierda()
tx = tx & "  <table width=""100%"" height=""206"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""FONT-SIZE: 10px; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; BORDER-COLLAPSE: collapse"">"
tx = tx & "    <tr>"
tx = tx & "      <td colspan=""2"" align=""center"">"
             tabla_xx3
tx = tx & "      </td>"
tx = tx & "    </tr>"
tx = tx & "    <tr>"
             if objFile.FileExists(server.mappath("/db/images/jugadores/" & trim(rsMisc.fields("Matricula")) & ".jpg" )) then     
tx = tx & "      <td width=""50%"" align=""center""><img src=""http://www.golfguide.com.ar/db/images/jugadores/" & trim(rsMisc.fields("Matricula")) & ".jpg"" width=""102"" height=""90""></td>"
             else
tx = tx & "      <td width=""50%"" align=""center""><img src=""http://www.golfguide.com.ar/db/images/jugadores/00000.jpg"" width=""102"" height=""90""></td>"
             end if
tx = tx & "      <td align=""center"">"
             TablaEtiquetas
tx = tx & "      </td>"
tx = tx & "    </tr>"
tx = tx & "    <tr>"
tx = tx & "      <td colspan=""2"" align=""center"">"
             tabla_xx4  
tx = tx & "      </td>"
tx = tx & "    </tr>"
tx = tx & "  </table>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub TablaEtiquetas()
tx = tx & "  <TABLE style=""FONT-SIZE: 12px; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; BORDER-COLLAPSE: collapse"" borderColor=""#c0c0c0"" cellSpacing=""0"" width=""50"" border=""1"">"
tx = tx & "    <TR>"
tx = tx & "      <TD align=""center"" bgColor=""#3399FF""><B><font color=""#66FFFF"">Birdie/-</font></B></TD>"
tx = tx & "    </TR>"
tx = tx & "    <TR>"
tx = tx & "      <TD align=""center"" bgColor=""#339900""><B><font color=""#33FF99"">Par</font></B></TD>"
tx = tx & "    </TR>"
tx = tx & "    <TR>"
tx = tx & "      <TD align=""center"" bgColor=""#FF6600""><B><font color=""#FFCC66"">Bogie</font></B></TD>"
tx = tx & "    </TR>"
tx = tx & "    <TR>"
tx = tx & "    <TR>"
tx = tx & "      <TD align=""center"" bgColor=""#990000""><B><FONT color=""#ffe0e0"">D.Bogie</FONT></B></TD>"
tx = tx & "    </TR>"
tx = tx & "    <TR>"
tx = tx & "      <TD align=""center"" bgColor=""#990066""><B><FONT color=""#FFCCFF"">T.Bogie/+</FONT></B></TD>"
tx = tx & "    </TR>"
tx = tx & "  </TABLE>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub tabla_xx3()
tx = tx & "  <TABLE style=""FONT-SIZE: 12px; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; BORDER-COLLAPSE: collapse"" borderColor=""#c0c0c0"" cellSpacing=""0"" width=""100%"" border=""1"">"
tx = tx & "    <TR bgcolor=""#DFDFDF"">"
tx = tx & "      <TD align=""center"" width=""33%""><b>Yardas:</b></TD>"
tx = tx & "      <TD align=""center"" width=""33%""><b>PAR:</b></TD>"
tx = tx & "      <TD align=""center"" width=""33%""><b>Gross:</b></TD>"
tx = tx & "    </TR>"
tx = tx & "    <TR bgColor=""#808080"">"
tx = tx & "      <TD align=""center""><B><FONT color=""#ffffff"">" & (YardasIda + YardasVta) & "</FONT></B></TD>"
tx = tx & "      <TD align=""center""><B><FONT color=""#ffffff"">" & (ParIda + ParVta) & "</FONT></B></TD>"
tx = tx & "      <TD align=""center""><B><FONT color=""#ffffff"">" & (ScoreIda + ScoreVta) & "</FONT></B></TD>"
tx = tx & "    </TR>"
tx = tx & "  </TABLE>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------


sub tabla_xx4()
tx = tx & "  <TABLE style=""FONT-SIZE: 12px; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; BORDER-COLLAPSE: collapse"" borderColor=""#c0c0c0"" cellSpacing=""0"" width=""100%"" border=""1"">"
tx = tx & "    <TR bgColor=""#000000"">"
tx = tx & "      <TD align=""center"" width=""30%""><B><FONT color=""#ffffff"">Gross:</FONT></B></TD>"
tx = tx & "      <TD align=""center"" width=""40%""><B><FONT color=""#ffffff"">Handicap:</FONT></B></TD>"
tx = tx & "      <TD align=""center"" width=""30%""><B><FONT color=""#ffffff"">Neto:</FONT></B></TD>"
tx = tx & "    </TR>"
tx = tx & "    <TR>"
tx = tx & "      <TD align=""center""><B>" & ScoreIda + ScoreVta & "</B></TD>"
tx = tx & "      <TD align=""center""><B>" & rsMisc.fields("Handicap") & "</B></TD>"
tx = tx & "      <TD align=""center"" style=""font-size:18px""><B>" & (ScoreIda + ScoreVta - rsMisc.fields("Handicap")) & "</B></TD>"
tx = tx & "    </TR>"
tx = tx & "  </TABLE>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------


sub TablaMitad(Desde) 
             Yardas=0
             Par=0
             Score=0
tx = tx & "  <table width=""100%"" border=""1"" cellpadding=""0"" cellspacing=""0"" bordercolor=""#cccccc"" style=""font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px; border-collapse: collapse"">"
             FHoyo(Desde)
             FYardas(Desde)
             FPar(Desde)
             FGolpes(Desde )
tx = tx & "    <tr>"
             if Desde=1 then
tx = tx & "      <td height=""60"" colspan=""11""><img src=""http://www.golfguide.com.ar/db/images/" & trim(rsMisc.fields("NombreC")) & "_01_09.jpg"" width=""430"" height=""60""></td>"
             else
tx = tx & "      <td height=""60"" colspan=""11""><img src=""http://www.golfguide.com.ar/db/images/" & trim(rsMisc.fields("NombreC")) & "_10_18.jpg"" width=""430"" height=""60""></td>"
             end if
tx = tx & "    </tr>"
tx = tx & "  </table>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub FHoyo(Desde) 
tx = tx & "    <tr>"
tx = tx & "      <td  align=""center"" width=""13%"" bgcolor=""#000000""><font color=""#FFFFFF""><b>Hoyo</b></font>"
tx = tx & "      </td>"
             for i = Desde to Desde + 8
tx = tx & "      <td align=""center"" width=""8%"" bgcolor=""#DFDFDF"">"
tx = tx & "        <b>" & i & "</b>"
tx = tx & "      </td>"
             next
             if Desde = 1 then
tx = tx & "      <td align=""center"" width=""13%"" bgcolor=""#999999""><b>Ida</b>"
             else
tx = tx & "      <td align=""center"" width=""13%"" bgcolor=""#999999""><b>Vuelta</b>"
             end if
tx = tx & "      </td>"
tx = tx & "    </tr>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub FYardas(Desde) 
tx = tx & "    <tr>"
tx = tx & "      <td align=""center"" bgcolor=""#000000""><font color=""#FFFFFF""><b>Yardas</b></font>"
tx = tx & "      </td>"
             rsTarjeta.MoveFirst
             for i = 1 to Desde - 1
               rsTarjeta.MoveNext
             Next
             for i = 1 to 9
               Select Case trim(LCase(rsMisc.Fields(9)))
                 Case "amarillo" 
tx = tx & "      <td align=""center"" bgcolor=""#FFFFC0""><font color=""#000000"">" & rsTarjeta.fields("Yardas") & "</font>"
                 Case "azul"
tx = tx & "      <td align=""center"" bgcolor=""#B0C4DE""><font color=""#000000"">" & rsTarjeta.fields("Yardas") & "</font>"
                 Case "blanco" 
tx = tx & "      <td align=""center"" bgcolor=""#FFFFFF""><font color=""#000000"">" & rsTarjeta.fields("Yardas") & "</font>"
                 Case "negro" 
tx = tx & "      <td align=""center"" bgcolor=""#696969""><font color=""#FFFFFF"">" & rsTarjeta.fields("Yardas") & "</font>"
                 Case "rojo" 
tx = tx & "      <td align=""center"" bgcolor=""#FFC0C0""><font color=""#000000"">" & rsTarjeta.fields("Yardas") & "</font>"
                 Case Else 
tx = tx & "      <td align=""center"">" & rsTarjeta.fields("Yardas")
               End Select 
tx = tx & "      </td>"
               Yardas = Yardas + rsTarjeta.fields("Yardas")
               rsTarjeta.MoveNext
             next
tx = tx & "      <td align=""center"" bgcolor=""#999999""><b>" & Yardas & "</b>"
tx = tx & "      </td>"
tx = tx & "    </tr>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub FPar(Desde) 
tx = tx & "    <tr>"
tx = tx & "      <td align=""center"" bgcolor=""#000000""><font color=""#FFFFFF""><b>PAR</b></font>"
tx = tx & "      </td>"
             rsTarjeta.MoveFirst
             for i = 1 to Desde - 1
               rsTarjeta.MoveNext
             Next
             for i = 1 to 9
tx = tx & "      <td align=""center"" bgcolor=""#DFDFDF"">" & rsTarjeta.fields("Par") 
tx = tx & "      </td>"
               Par = Par + rsTarjeta.fields("Par")
               rsTarjeta.MoveNext
             next
tx = tx & "      <td align=""center"" bgcolor=""#999999""><b>" & Par & "</b>"
tx = tx & "      </td>"
tx = tx & "    </tr>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub FGolpes(Desde) 
tx = tx & "    <tr>"
tx = tx & "      <td align=""center"" bgcolor=""#000000"" style=""font-size:14px""><font color=""#FFFFFF""><b>SCORE</b></font>"
tx = tx & "      </td>"
             rsTarjeta.MoveFirst
             for i = 1 to Desde - 1
               rsTarjeta.MoveNext
             Next
             for i = 1 to 9
               if rsTarjeta.Fields("Golpes") <= rsTarjeta.Fields("Par") - 1 then
                 'birdie o menos (Celeste)
tx = tx & "      <td align=""center"" bgColor=""#3399FF"" style=""font-size:14px""><font color=""#66FFFF""><b>" & rsTarjeta.fields("Golpes") & "</b></font>"
tx = tx & "      </td>"
               else
                 if rsTarjeta.Fields("Golpes") = rsTarjeta.Fields("Par") then
                   'par (verde)
tx = tx & "      <td align=""center"" bgColor=""#339900"" style=""font-size:14px""><font color=""#33FF99""><b>" & rsTarjeta.fields("Golpes") & "</b></font>"
tx = tx & "      </td>"
                 else
                   if rsTarjeta.Fields("Golpes") = rsTarjeta.Fields("Par") + 1 then
                     'bogie (Naranja)
tx = tx & "      <td align=""center"" bgcolor=""#FF6600"" style=""font-size:14px""><font color=""#FFCC66""><b>" & rsTarjeta.fields("Golpes") & "</b></font>"
tx = tx & "      </td>"
                   else
                     if rsTarjeta.Fields("Golpes") = rsTarjeta.Fields("Par") + 2 then
                       'doble bogie (Rojo)
tx = tx & "      <td align=""center"" bgcolor=""#990000"" style=""font-size:14px""><font color=""#ffe0e0""><b>" & rsTarjeta.fields("Golpes") & "</b></font>"
tx = tx & "      </td>"
                     else 
                       'triple bogie (violeta)
tx = tx & "      <td align=""center"" bgcolor=""#990066"" style=""font-size:14px""><font color=""#ffccff""><b>" & rsTarjeta.fields("Golpes") & "</b></font>"
tx = tx & "      </td>"
                     end if
                   end if
                 end if
               end if
               Score = Score + rsTarjeta.fields("Golpes")
               rsTarjeta.MoveNext
             next
tx = tx & "      <td align=""center"" bgcolor=""#999999"" style=""font-size:14px""><b>" & Score & "</b>"
tx = tx & "      </td>"
tx = tx & "    </tr>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub TablaTitulos2()
tx = tx & "  <TABLE cellSpacing=0 cellPadding=2 width=""100%""  border=0 style=""FONT-SIZE: 10px; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; BORDER-COLLAPSE: collapse"">"
tx = tx & "    <TR>"
             if rsMIsc.Fields("Pos")<>99 then
tx = tx & "      <TD colspan=""2"" style=""font-size:14px""><B><FONT color=""#FF0000""><span style=""text-transform: capitalize"">" & rsMisc.fields("Jugador") & "</span></FONT>&nbsp;&nbsp;&nbsp;&nbsp;(" & rsMisc.fields("Pos") & ")</B></TD>"
             else
tx = tx & "      <TD colspan=""2"" style=""font-size:14px""><B><FONT color=""#FF0000""><span style=""text-transform: capitalize"">" & rsMisc.fields("Jugador") & "</span></FONT>&nbsp;&nbsp;&nbsp;&nbsp;(--)</B></TD>"
             end if
tx = tx & "    </TR><TR>"
tx = tx & "      <TD colspan=""2""><B>" & rsMisc.fields("Cancha") & "&nbsp;&nbsp;&nbsp;&nbsp;" & toFecha(rsMisc.Fields("Fecha"))
'  Select Case mid(rsMisc.Fields("Fecha"),3,2) 
'    Case "01" tx = tx & " de Enero de 202" & mid(rsMisc.Fields("Fecha"),2,1)
'    Case "02" tx = tx & " de Febrero de 202" & mid(rsMisc.Fields("Fecha"),2,1)
'    Case "03" tx = tx & " de Marzo de 202" & mid(rsMisc.Fields("Fecha"),2,1)
'    Case "04" tx = tx & " de Abril de 202" & mid(rsMisc.Fields("Fecha"),2,1)
'    Case "05" tx = tx & " de Mayo de 202" & mid(rsMisc.Fields("Fecha"),2,1)
'    Case "06" tx = tx & " de Junio de 202" & mid(rsMisc.Fields("Fecha"),2,1)
'    Case "07" tx = tx & " de Julio de 202" & mid(rsMisc.Fields("Fecha"),2,1)
'    Case "08" tx = tx & " de Agosto de 202" & mid(rsMisc.Fields("Fecha"),2,1)
'    Case "09" tx = tx & " de Septiembre de 202" & mid(rsMisc.Fields("Fecha"),2,1)
'    Case "10" tx = tx & " de Octubre de 202" & mid(rsMisc.Fields("Fecha"),2,1)
'    Case "11" tx = tx & " de Noviembre de 202" & mid(rsMisc.Fields("Fecha"),2,1)
'    Case "12" tx = tx & " de Diciembre de 202" & mid(rsMisc.Fields("Fecha"),2,1)
'  End Select
tx = tx & "        </B></TD>"
tx = tx & "    </TR>"
tx = tx & "    <TR>"
tx = tx & "      <TD width=""45%""><B>"
             if rsMisc.Fields("D") then
tx = tx & "        Damas"
             end if
             if rsMisc.Fields("C") then
                if rsMisc.Fields("D") then
tx = tx & "        - Caballeros"
                else
tx = tx & "        Caballeros"
                end if
             end if
tx = tx & "        &nbsp;&nbsp;&nbsp;&nbsp;"
             if rsMisc.Fields("Desde")=-3 then
tx = tx & "        Hasta " & rsMisc.Fields("Hasta")
             else
tx = tx & rsMisc.Fields("Desde") & " a " & rsMisc.Fields("Hasta")
             end if
tx = tx & "        </B></TD>"
tx = tx & "      <TD><B><font color=""#FF0000"">" & rsMisc.Fields("Tee") & "(" & rsMisc.Fields("Calificacion") & ")</font></B></TD>"
tx = tx & "    </TR>"
tx = tx & "  </TABLE>"
end sub


'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------
sub WriteRankingTop3v2(pIdRanking,pHcpDesde,pHcpHasta,sTituloCat)


dim rsCatN
Set rsCatN=QueryRankingTop3(cmdDC,Fecha,pIdRanking,pHcpDesde,pHcpHasta)

tx = tx & "      <tr>"
tx = tx & "        <td align=""center"" colspan=""3"">"
tx = tx & "          <b>" & sTituloCat & "</b>"
tx = tx & "        </td>"
tx = tx & "      </tr>"
rsCatN.MoveFirst
dim sBoltI
dim sBoltF
dim nPos
sBoltI="<b>"
sBoltF="</b>"
nPos=1
do while not rsCatN.eof
    tx = tx & "      <tr>"
    tx = tx & "        <td>" & sBoltI & "<font color=""#ff0000"">" & nPos & ".&nbsp;&nbsp;</font>" & sBoltF & "</td>"
    tx = tx & "        <td>" & sBoltI & "<font color=""#ff0000"">" & propercase(rsCatN("Nombre")) & "</font>" & sBoltF & "</td>"
    tx = tx & "        <td>" & sBoltI & "<font color=""#ff0000"">" & rsCatN("Puntaje") & "&nbsp;Pts.</font>" & sBoltF & "</td>"
    tx = tx & "      </tr>"
    rsCatN.MoveNext
    sBoltI=""
    sBoltF=""
    nPos=nPos+1
loop
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub WriteRankingCore(pIdRankig, pImageRanking, pTitulo,pSubTitulo,pDetalle) 

' pIdRankig: id del tanking Ej: 277
' pImageRanking: imagen del logo del ranking   Ej: "iberostar.jpg"
' pTitulo: Texto Principal  Ej: "RANKING IBEROSTAR 2018 (Mis resultados, click aqui)"
' pSubTitulo: Text del sub titulo Ej: "Cierra el 20 de Marzo 2018 | Final 54 hoyos en Abril"
' pDetalle: Observaciin / Detalle Ej: "Clasifican 36 por Categoria | Minimo 8 torneos jugados" 


   tx = tx & chr(13)&chr(10) & "        <tr>"
   tx = tx & chr(13)&chr(10) & "          <td align=center>"
   tx = tx & chr(13)&chr(10) & "            <IMG src=""http://www.golfguide.com.ar/mail/2007/images/" & pImageRanking & """>"
   tx = tx & chr(13)&chr(10) & "            <FONT size=3><b><A href=""http://www.golfguide.com.ar/db/ranking2019.asp?id=" & pIdRankig & """>" & pTitulo & "</A></FONT>"
   tx = tx & chr(13)&chr(10) & "            <BR>"
   tx = tx & chr(13)&chr(10) & "            <FONT size=2><b>" & pSubTitulo & "</A></FONT>"
   tx = tx & chr(13)&chr(10) & "            <BR>"
   tx = tx & chr(13)&chr(10) & "            <FONT size=2><font color=""#999999"">" & pDetalle & "<b></A></FONT>"
   tx = tx & chr(13)&chr(10) & "            <BR>"
   tx = tx & chr(13)&chr(10) & "            <BR>"
                          'ranking5
                          WriteRankingPosAct pIdRankig 
                          WriteRankingNv2 pIdRankig
   tx = tx & chr(13)&chr(10) & "          </td>"
   tx = tx & chr(13)&chr(10) & "        </tr>"

end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------


sub WriteRankingNv2(pIdRanking)

tx = tx & "<table cellspacing=""0"" cellpadding=""0"" width=""350"" border=""0"" style=""font-family: Verdana; font-size:8pt; border-collapse:collapse"" bordercolor=""#111111"">"

dim aPos
aPos=0
For Each item In categoriaHCPDesde
	WriteRankingTop3v2 pIdRanking, categoriaHCPDesde(aPos), categoriaHCPHasta(aPos), categoriaDescripcion(aPos)
         aPos=aPos+1
Next

'For index = 0 To categoriaHCPDesde.Length-1 'categoriaHCPDesde.GetUpperBound(0)
'         'Console.WriteLine(categoriaHCPDesde(index))
'         WriteRankingTop3v2 pIdRanking, categoriaHCPDesde(index), categoriaHCPHasta(index), categoriaDescripcion(index)
'Next


tx = tx & "      <tr>"
tx = tx & "        <td align=""center"" colspan=""3"">"
tx = tx & "      </tr>"
tx = tx & "    </table>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub WriteRankingPosAct(pIdRanking)

dim sBoltI
dim sBoltF
sBoltI=""
sBoltF=""

dim DesHcp
dim HasHcp
dim jHcp

DesHcp=0
HasHcp=0
jHcp=rsMisc.fields("Handicap")

if jHcp<0 then
    jHcp=0
end if

dim aPos
aPos=0
For Each item In categoriaHCPDesde
	if jHcp>=categoriaHCPDesde(aPos) and jHcp<=categoriaHCPHasta(aPos) then
             DesHcp=categoriaHCPDesde(aPos)
             HasHcp=categoriaHCPHasta(aPos)
         end if
         aPos=aPos+1
Next

if HasHcp>0 then
    dim rsRankingMatricula
    set rsRankingMatricula=QueryRankingMatricula(cmdDC,Fecha,pIdRanking,DesHcp,HasHcp)
    rsRankingMatricula.MoveFirst
    do while not rsRankingMatricula.eof
        'SALTOS..?
        if rsRankingMatricula("Puntaje")>0 then
            
            tx = tx & "<table cellspacing=""0"" cellpadding=""0"" width=""473"" border=""0"" style=""font-family: Verdana; font-size:12pt; border-collapse:collapse"" bordercolor=""#111111"" bgcolor=""#00FFFF"">"
            tx = tx & "      <tr>"
            
            tx = tx & "        <td>" & sBoltI & "<font color=""#00FFFF"">.<strong><font color=""#0000CC"" size=""3"">" & rsRankingMatricula("ranking_posicion") & ".&nbsp;&nbsp;</font>" & sBoltF & "</td>"
            tx = tx & "        <td>" & sBoltI & "<strong><font color=""#0000CC"" size=""3"">" & propercase(rsRankingMatricula("Nombre")) & "</font>" & sBoltF & "</td>"
            tx = tx & "        <td>" & sBoltI & "<strong><font color=""#0000CC"" size=""3"">" & rsRankingMatricula("Puntaje") & "&nbsp;Pts.</font>" & sBoltF & "</td>"

            
            tx = tx & "      </tr>"
            tx = tx & "      <tr>"
            tx = tx & "        <td align=""center"" colspan=""3"">"
            tx = tx & "      </tr>"
            tx = tx & "    </table>"
            'SALTOS..?
            Exit Do 
        end if
        rsRankingMatricula.MoveNext
    loop
        
    
end if

end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub ganadores()
rsFecha.MoveFirst
tx = tx & "<table cellspacing=""0"" cellpadding=""0"" width=""520"" border=""0"" style=""font-family: Verdana; font-size:8pt; border-collapse:collapse"">"
if rsFecha("Matricula1") <> 0 or rsFecha("Matricula2") <> 0 then		
  tx = tx & "        <tr>"
  tx = tx & "          <td colspan=""3"" align=""center"" width=""50%""><br>"
  tx = tx & "            <b>" & trim(rsFecha("Premio1")) & "</b><br>"
  cmdDC.CommandText = "select Nombre from Jugadores " &_
  "where Matricula=" & rsFecha("Matricula1")
  cmdDC.CommandType = 1
  Set rsJugador = Server.CreateObject("ADODB.Recordset")
  rsJugador.Open cmdDC, , 0, 1
  tx = tx & "           <font color=""#ff0000"">" & propercase(rsJugador("Nombre")) & "</font>"
  tx = tx & "           </td>"
  tx = tx & "           <td colspan=""3"" align=""center"" width=""50%""><br>"
  tx = tx & "             <b>" & trim(rsFecha("Premio2")) & "</b><br>"
  cmdDC.CommandText = "select Nombre from Jugadores " &_
  "where Matricula=" & rsFecha("Matricula2")
  cmdDC.CommandType = 1
  Set rsJugador = Server.CreateObject("ADODB.Recordset")
  rsJugador.Open cmdDC, , 0, 1
  tx = tx & "           <font color=""#ff0000"">" & propercase(rsJugador("Nombre")) & "</font>"
  tx = tx & "           </td>"
  tx = tx & "      </tr>"
end if  
if rsFecha("Matricula3") <> 0 or rsFecha("Matricula4") <> 0 then		
  tx = tx & "        <tr>"
  tx = tx & "          <td colspan=""3"" align=""center"" width=""50%""><br>"
  tx = tx & "            <b>" & trim(rsFecha("Premio3")) & "</b><br>"
  cmdDC.CommandText = "select Nombre from Jugadores " &_
  "where Matricula=" & rsFecha("Matricula3")
  cmdDC.CommandType = 1
  Set rsJugador = Server.CreateObject("ADODB.Recordset")
  rsJugador.Open cmdDC, , 0, 1
  tx = tx & "           <font color=""#ff0000"">" & propercase(rsJugador("Nombre")) & "</font>"
  tx = tx & "           </td>"
  tx = tx & "           <td colspan=""3"" align=""center"" width=""50%""><br>"
  tx = tx & "             <b>" & trim(rsFecha("Premio4")) & "</b><br>"
  cmdDC.CommandText = "select Nombre from Jugadores " &_
  "where Matricula=" & rsFecha("Matricula4")
  cmdDC.CommandType = 1
  Set rsJugador = Server.CreateObject("ADODB.Recordset")
  rsJugador.Open cmdDC, , 0, 1
  tx = tx & "           <font color=""#ff0000"">" & propercase(rsJugador("Nombre")) & "</font>"
  tx = tx & "           </td>"
  tx = tx & "      </tr>"
end if

rsScores.MoveFirst
rsScores2.MoveFirst

do while not rsScores.eof
  tx = tx & "      <tr>"
  tx = tx & "        <td colspan=""3"" align=""center""><br>"
  tx = tx & "          <b>"
  if rsScores.Fields("D") and not rsScores.Fields("C") then
    tx = tx & "        Damas"
  end if
  if rsScores.Fields("C") then
    if rsScores.Fields("D") then
      tx = tx & "      Dam - Cab"
    else
      tx = tx & "      Caballeros"
    end if
  end if
  if rsScores.Fields("Desde")=-3 then
    if rsScores.Fields("Hasta")<>36 then
      tx = tx & "      Hasta " & rsScores.Fields("Hasta") 
    end if
  else
    tx = tx & " " & rsScores.Fields("Desde") & " a " & rsScores.Fields("Hasta")
  end if
  
  rsScores2.MoveNext
  rsScores2.MoveNext
  
  if not rsScores2.eof then
    tx = tx & "        <td colspan=""3"" align=""center""><br>"
    tx = tx & "          <b>"
    if rsScores2.Fields("D") and not rsScores2.Fields("C") then
      tx = tx & "          Damas"
    end if
    if rsScores2.Fields("C") then
      if rsScores2.Fields("D") then
        tx = tx & "      Dam - Cab"
      else
        tx = tx & "      Caballeros"
      end if
    end if
    if rsScores2.Fields("Desde")=-3 then
      if rsScores2.Fields("Hasta")<>36 then
        tx = tx & "     Hasta " & rsScores2.Fields("Hasta") 
      end if
    else
      tx = tx & " " & rsScores2.Fields("Desde") & " a " & rsScores2.Fields("Hasta")
    end if
    tx = tx & "         </b></td>"
  else
    tx = tx & "    <td colspan=""3""></td>"
  end if
  tx = tx & "<tr>"
  
  tx = tx & "    <td align=""center"" width=""5%""><font color=""#ff0000"">1.</font></td>"
  tx = tx & "    <td width=""37%""><font color=""#ff0000"">" & propercase(rsScores("Nombre")) & "</font></td>"
  tx = tx & "    <td width=""8%""><font color=""#ff0000"">(" & trim(rsScores.Fields("Neto")) & ")</font></td>"
  if not rsScores2.eof then
    tx = tx & "    <td align=""center"" width=""5%""><font color=""#ff0000"">1.</font></td>"
    tx = tx & "    <td width=""37%""><font color=""#ff0000"">" & propercase(rsScores2("Nombre")) & "</font></td>"
    tx = tx & "    <td width=""8%""><font color=""#ff0000"">(" & trim(rsScores2.Fields("Neto")) & ")</font></td>"
  else
    tx = tx & "    <td></td>"
    tx = tx & "    <td></td>"
    tx = tx & "    <td></td>"
  end if
  tx = tx & "</tr>"
  rsScores.MoveNext
  tx = tx & "    <td align=""center""><font color=""#ff0000"">2.</font></td>"
  tx = tx & "    <td><font color=""#ff0000"">" & propercase(rsScores("Nombre")) & "</font></td>"
  tx = tx & "    <td><font color=""#ff0000"">(" & trim(rsScores.Fields("Neto")) & ")</font></td>"
  rsScores.MoveNext
  if not rsScores2.eof then
    rsScores2.MoveNext
    tx = tx & "    <td align=""center""><font color=""#ff0000"">2.</font></td>"
    tx = tx & "    <td><font color=""#ff0000"">" & propercase(rsScores2("Nombre")) & "</font></td>"
    tx = tx & "    <td><font color=""#ff0000"">(" & trim(rsScores2.Fields("Neto")) & ")</font></td>"
    rsScores2.MoveNext
    rsScores.MoveNext
    rsScores.MoveNext
  else
    tx = tx & "    <td></td>"
    tx = tx & "    <td></td>"
    tx = tx & "    <td></td>"
  end if
  tx = tx & "</tr>"
loop
tx = tx & "    </table>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub proximos()

rsMiercoles.MoveFirst
rsJueves.MoveFirst
rsViernes.MoveFirst
rsDomingos.MoveFirst

tx = tx & "<table cellspacing=""0"" cellpadding=""0"" width=""520"" border=""0"" style=""font-family: Verdana; font-size:8pt; border-collapse:collapse"">"
  tx = tx & "  <tr>"
  tx = tx & "    <td>&nbsp;</td>"
  tx = tx & "    <td>&nbsp;</td>"
  tx = tx & "    <td>&nbsp;</td>"
  tx = tx & "  </tr>"
  tx = tx & "  <tr>"
  tx = tx & "    <td align=""center"" width=""25%""><font color=""#999999"">TEMPORADA 2009</font>"
  tx = tx & "    </td>"
  tx = tx & "    <td align=""center"" width=""25%""><font color=""#999999"">TEMPORADA 2009</font>"
  tx = tx & "    </td>"
  tx = tx & "    <td align=""center"" width=""25%""><font color=""#999999"">TEMPORADA 2009</font>"
  tx = tx & "    </td>"
  tx = tx & "    <td align=""center"" width=""25%""><font color=""#999999"">TEMPORADA 2009</font>"
  tx = tx & "    </td>"
  tx = tx & "  </tr>"
  if not (rsMiercoles.eof and rsJueves.eof and rsViernes.eof and rsDomingos.eof) then 
    tx = tx & "  <tr>"
    if not rsMiercoles.eof then
      tx = tx & "    <td><font color=""#FF0000"">Miircoles </font>" & tofecha(rsMiercoles("Fecha"))
      tx = tx & "    </td>"
    else
      tx = tx & "<td></td>"
    end if
    if not rsJueves.eof then
      tx = tx & "    <td><font color=""#FF0000"">Jueves </font>" & tofecha(rsJueves("Fecha"))
      tx = tx & "    </td>"
    else
      tx = tx & "<td></td>"
    end if
    if not rsViernes.eof then
      tx = tx & "    <td><font color=""#FF0000"">Viernes </font>" & tofecha(rsViernes("Fecha"))
      tx = tx & "    </td>"
    else
      tx = tx & "<td></td>"
    end if
    if not rsDomingos.eof then
      tx = tx & "    <td><font color=""#FF0000"">Domingo </font>" & tofecha(rsDomingos("Fecha"))
      tx = tx & "    </td>"
    else
      tx = tx & "<td></td>"
    end if
    tx = tx & "  </tr>"
    tx = tx & "  <tr>"
    if not rsMiercoles.eof then
      tx = tx & "    <td>" & ucase(rsMiercoles("Nombre"))
      tx = tx & "    </td>"
    else
      tx = tx & "<td></td>"
    end if
    if not rsJueves.eof then
      tx = tx & "    <td>" & ucase(rsJueves("Nombre"))
      tx = tx & "    </td>"
    else
      tx = tx & "<td></td>"
    end if
    if not rsViernes.eof then
      tx = tx & "    <td>" & ucase(rsViernes("Nombre"))
      tx = tx & "    </td>"
    else
      tx = tx & "<td></td>"
    end if
    if not rsDomingos.eof then
      tx = tx & "    <td>" & ucase(rsDomingos("Nombre"))
      tx = tx & "    </td>"
    else
      tx = tx & "<td></td>"
    end if
    tx = tx & "  </tr>"
    tx = tx & "  <tr>"
    if not rsMiercoles.eof then
      tx = tx & "    <td>"
      tx = tx & "      <FONT color=""#666666"">Reservas - <A href=""mailto:circuito@golfguide.com.ar?subject=Reserva Miircoles " & int(right(rsMiercoles("Fecha"),2)) & " Circuito Golf Guide"">click aqui</A></FONT>"
    tx = tx & "    </td>"
      rsMiercoles.MoveNext
    else
      tx = tx & "<td></td>"
    end if
    if not rsJueves.eof then
      tx = tx & "    <td>"
      tx = tx & "      <FONT color=""#666666"">Reservas - <A href=""mailto:circuito@golfguide.com.ar?subject=Reserva Jueves " & int(right(rsJueves("Fecha"),2)) & " Circuito Golf Guide"">click aqui</A></FONT>"
    tx = tx & "    </td>"
      rsJueves.MoveNext
    else
      tx = tx & "<td></td>"
    end if
    if not rsViernes.eof then
      tx = tx & "    <td>"
      tx = tx & "      <FONT color=""#666666"">Reservas - <A href=""mailto:circuito@golfguide.com.ar?subject=Reserva Viernes " & int(right(rsViernes("Fecha"),2)) & " Circuito Golf Guide"">click aqui</A></FONT>"
      tx = tx & "    </td>"
      rsViernes.MoveNext
    else
      tx = tx & "<td></td>"
    end if
    if not rsDomingos.eof then
      tx = tx & "    <td>"
      tx = tx & "      <FONT color=""#666666"">Reservas - <A href=""mailto:circuito@golfguide.com.ar?subject=Reserva Domingo " & int(right(rsDomingos("Fecha"),2)) & " Circuito Golf Guide"">click aqui</A></FONT>"
      tx = tx & "    </td>"
      rsDomingos.MoveNext
    else
      tx = tx & "<td></td>"
    end if
    tx = tx & "  </tr>"
  end if
  if not (rsMiercoles.eof and rsJueves.eof and rsViernes.eof and rsDomingos.eof) then 
    tx = tx & "  <tr>"
    tx = tx & "    <td>&nbsp;</td>"
    tx = tx & "    <td>&nbsp;</td>"
    tx = tx & "    <td>&nbsp;</td>"
    tx = tx & "  </tr>"
    tx = tx & "  <tr>"
    if not rsMiercoles.eof then
      tx = tx & "    <td><font color=""#FF0000"">Miircoles </font>" & tofecha(rsMiercoles("Fecha"))
      tx = tx & "    </td>"
    else
      tx = tx & "<td></td>"
    end if
    if not rsJueves.eof then
      tx = tx & "    <td><font color=""#FF0000"">Jueves </font>" & tofecha(rsJueves("Fecha"))
      tx = tx & "    </td>"
    else
      tx = tx & "<td></td>"
    end if
    if not rsViernes.eof then
      tx = tx & "    <td><font color=""#FF0000"">Viernes </font>" & tofecha(rsViernes("Fecha"))
      tx = tx & "    </td>"
    else
      tx = tx & "<td></td>"
    end if
    if not rsDomingos.eof then
      tx = tx & "    <td><font color=""#FF0000"">Domingo </font>" & tofecha(rsDomingos("Fecha"))
      tx = tx & "    </td>"
    else
      tx = tx & "<td></td>"
    end if
    tx = tx & "  </tr>"
    tx = tx & "  <tr>"
    if not rsMiercoles.eof then
      tx = tx & "    <td>" & ucase(rsMiercoles("Nombre"))
      tx = tx & "    </td>"
    else
      tx = tx & "<td></td>"
    end if
    if not rsJueves.eof then
      tx = tx & "    <td>" & ucase(rsJueves("Nombre"))
      tx = tx & "    </td>"
    else
      tx = tx & "<td></td>"
    end if
    if not rsViernes.eof then
      tx = tx & "    <td>" & ucase(rsViernes("Nombre"))
      tx = tx & "    </td>"
    else
      tx = tx & "<td></td>"
    end if
    if not rsDomingos.eof then
      tx = tx & "    <td>" & ucase(rsDomingos("Nombre"))
      tx = tx & "    </td>"
    else
      tx = tx & "<td></td>"
    end if
    tx = tx & "  </tr>"
  end if
  tx = tx & "</table>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

function tofecha(entrada)
  dim parte
  parte=right(entrada,2)
  Select Case left(right(entrada,4),2) 
    Case "01" parte=parte & " de Enero 202"
    Case "02" parte=parte & " de Febrero 202"
    Case "03" parte=parte & " de Marzo 202"
    Case "04" parte=parte & " de Abril 202"
    Case "05" parte=parte & " de Mayo 202"
    Case "06" parte=parte & " de Junio 202"
    Case "07" parte=parte & " de Julio 202"
    Case "08" parte=parte & " de Agosto 202"
    Case "09" parte=parte & " de Septiembre 202"
    Case "10" parte=parte & " de Octubre 202"
    Case "11" parte=parte & " de Noviembre 202"
    Case "12" parte=parte & " de Diciembre 202"
  End Select
  parte=parte & left(right(entrada,5),1)
  tofecha=parte
end function
'-------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------

Function propercase(entrada)
  Dim sw
  Dim aux
  Dim i
  sw = False
  aux = ""
  For i = 1 To Len(entrada)
    If i = 1 Or sw Then
      aux = aux + UCase(Mid(entrada, i, 1))
    Else
      aux = aux + LCase(Mid(entrada, i, 1))
    End If
    sw = Mid(entrada, i, 1) = " "
  Next
  propercase = aux
End Function
'-------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------

sub armarmail
tx = "<html>"
tx = tx & chr(13)&chr(10) & "<body vLink=""#ff0000"" aLink=""#ff0000"" link=""#ff0000"" bgColor=""#dfdfdf"" leftMargin=""0"" marginwidth=""0"">"
tx = tx & chr(13)&chr(10) & "<TABLE borderColor=""#000000"" cellSpacing=""0"" bgColor=""#ffffff"" cellPadding=""0"" width=""541"" align=""center"" border=""1"">"
tx = tx & chr(13)&chr(10) & "  <tr>"
tx = tx & chr(13)&chr(10) & "    <td>"
tx = tx & chr(13)&chr(10) & "      <TABLE cellSpacing=""0"" cellPadding=""0"" align=""center"" border=""0"" style=""FONT-SIZE: 12; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; BORDER-COLLAPSE: collapse"">"
tx = tx & chr(13)&chr(10) & "        <tr>"
tx = tx & chr(13)&chr(10) & "          <td align=""center"">"
tx = tx & chr(13)&chr(10) & "            <A href=""http://www.golfguide.com.ar""><IMG src=""http://www.golfguide.com.ar/mail/2007/images/logo_mail0.jpg"">"
tx = tx & chr(13)&chr(10) & "            <a href=""mailto:circuito@golfguide.com.ar?subject=Reserva%20Circuito%20Golf%20Guide""><IMG src=""http://www.golfguide.com.ar/mail/2007/images/logo_mail3.jpg"">"
tx = tx & chr(13)&chr(10) & "          </td>"
tx = tx & chr(13)&chr(10) & "        </tr>"
tx = tx & chr(13)&chr(10) & "        <tr>"
tx = tx & chr(13)&chr(10) & "          <td  align=""center"">"
tx = tx & chr(13)&chr(10) & "            <BR>"
                       tarjeta
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "          </td>"
tx = tx & chr(13)&chr(10) & "        </tr>"
tx = tx & chr(13)&chr(10) & "        <tr>"
tx = tx & chr(13)&chr(10) & "          <td align=""center"">"
tx = tx & chr(13)&chr(10) & "            <FONT size=4>Mis resultados: <A href=""http://www.golfguide.com.ar"">www.golfguide.com.ar</A></FONT>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <FONT size=1>----------------------------------------------------------</A></FONT>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "          </td>"
tx = tx & chr(13)&chr(10) & "        </tr>"
if G="ON" then
   tx = tx & chr(13)&chr(10) & "            <BR>"
   tx = tx & chr(13)&chr(10) & "        <tr>"
   tx = tx & chr(13)&chr(10) & "          <td align=center>"
   tx = tx & chr(13)&chr(10) & "            <FONT size=4><b>GANADORES DE LA FECHA</b></FONT>"
                          ganadores
   tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <FONT size=1>----------------------------------------------------------</A></FONT>"
tx = tx & chr(13)&chr(10) & "            <BR>"
   tx = tx & chr(13)&chr(10) & "          </td>"
   tx = tx & chr(13)&chr(10) & "        </tr>"
end if

' MUESTRA RANKINGS


      
      

if R="ON" then
   WriteRankingCore 201, "rank_buzios.jpg", "RANKING BUZIOS 2023 (Click aqui)", "LUNES a DOMINGOS", "Clasifican 45 por Categoria | Minimo 6 torneos jugados"
end if
if Q="ON" then
   WriteRankingCore 274, "toso2.jpg", "RANKING PASCUAL TOSO ALTA (Click aqui)", "LUNES a DOMINGOS", "Clasifican 30 por Categoria | Minimo 4 torneos jugados"
end if
if K="ON" then
  WriteRankingCore 278, "toso2.jpg", "RANKING PASCUAL TOSO ALTA (Click aqui)", "LUNES a DOMINGOS", "Clasifican 30 por Categoria | Minimo 4 torneos jugados"
end if
if Y="ON" then
   WriteRankingCore 325, "paraguay.jpg", "RANKING PARAGUAY 2018 (Mis resultados, click aqui)", "Cierra el 20 de Agosto 2018 | Final 18 hoyos en Septiembre", "Clasifican 16 por Categoria | Minimo 8 torneos jugados"
end if
if X="ON" then
   WriteRankingCore 207, "cordoba.jpg", "RANKING CORDOBA 2018 (Mis resultados, click aqui)", "Cierra el 30 de Agosto 2018 | Final MATCH PLAY en Septiembre", "Clasifican 8 por Categoria | Minimo 8 torneos jugados"
end if


' (** FIN **) MUESTRA RANKINGS

tx = tx & chr(13)&chr(10) & "        <tr>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "          <td align=center>"
   tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <FONT size=1>----------------------------------------------------------</A></FONT>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <BR>"
' tx = tx & chr(13)&chr(10) & "            <FONT size=4><b><A href=""http://golfguide.com.ar/circuito/reglas/reglas2019.htm"">REGLAS RANKING GOLF GUIDE (Click aqui)</A></FONT>"
' tx = tx & chr(13)&chr(10) & "            <BR>"
'    tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <FONT size=1>----------------------------------------------------------</A></FONT>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <a href=""https://youtu.be/z3NekLKHjno""><img border=""0""  src=""http://www.golfguide.com.ar/video2017.jpg"" width=""451"" /></a><br />"
   tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <FONT size=1>----------------------------------------------------------</A></FONT>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <IMG src=""http://www.golfguide.com.ar/mail/2007/images/sponsors_mail.jpg"">"
tx = tx & chr(13)&chr(10) & "            <IMG height=23 src=""http://www.golfguide.com.ar/mail/2006/images/barra.jpg"" width=451>"
tx = tx & chr(13)&chr(10) & "            <A href=""http://www.golfguide.com.ar""><FONT color=""#ff0000"" size=5>www.golfguide.com.ar</FONT></A>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "          </td>"
tx = tx & chr(13)&chr(10) & "        </tr>"
if P="ON" then
   tx = tx & chr(13)&chr(10) & "        <tr>"
   tx = tx & chr(13)&chr(10) & "          <td align=center>"
   tx = tx & chr(13)&chr(10) & "            <FONT size=3><b>Priximos Torneos</b></FONT>"
                          proximos
   tx = tx & chr(13)&chr(10) & "            <BR>"
   tx = tx & chr(13)&chr(10) & "            <BR>"
   tx = tx & chr(13)&chr(10) & "          </td>"
   tx = tx & chr(13)&chr(10) & "        </tr>"
end if
tx = tx & chr(13)&chr(10) & chr(13)&chr(10) & "        <tr>"
tx = tx & chr(13)&chr(10) & "          <td align=""center"">"
tx = tx & chr(13)&chr(10) & "            <FONT size=3 color=""#ff0000""><b>Mas informacion o reservas: (+54911) 6676.4653</b>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <a href=""mailto:circuito@golfguide.com.ar?subject=Reserva%20Circuito%20Golf%20Guide""><b>circuito@golfguide.com.ar</b></a>"
tx = tx & chr(13)&chr(10) & "            </FONT>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <FONT color=""#ff0000"">Para tener en cuenta: </FONT>LLAMAR CON ANTICIPACION."
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <FONT color=""#ff0000"">Si reservan y no pueden ir: </FONT>LLAMAR PARA CANCELAR."
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <FONT size=""2""><b>Por favor cuidemos las canchas!!:"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            <FONT size=""3"">ARREGLAR LOS PIQUES - ALISAR LOS BUNKERS<BR>REPONER LOS DIVOTS.</FONT>"
tx = tx & chr(13)&chr(10) & "            <BR>"
tx = tx & chr(13)&chr(10) & "            Y en los Countries CONTROLEN LA VELOCIDAD</b></FONT>"
tx = tx & chr(13)&chr(10) & "            <BR><BR>"
tx = tx & chr(13)&chr(10) & "          </td>"
tx = tx & chr(13)&chr(10) & "        </tr>"
tx = tx & chr(13)&chr(10) & "      </table>"
tx = tx & chr(13)&chr(10) & "    </td>"
tx = tx & chr(13)&chr(10) & "  </tr>"
tx = tx & chr(13)&chr(10) & "</table>"
tx = tx & chr(13)&chr(10) & "</body>"
tx = tx & chr(13)&chr(10) & "</html>"
end sub
'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

sub enviarmail(Linea)

cmdDC.CommandText = "insert into temp (linea) values ('" & trim(linea) & "')"
cmdDC.CommandType = 1
Set rsLinea = Server.CreateObject("ADODB.Recordset")
rsLinea.Open cmdDC, , 0, 1

cmdDC.CommandText = "select idscore from scores,fechas where scores.idfecha=fechas.idfecha and fecha='" & fecha & "' and matricula=" & left(Linea,instr(Linea,",")-1)
cmdDC.CommandType = 1
Set rsScore = Server.CreateObject("ADODB.Recordset")
rsScore.Open cmdDC, , 0, 1
idScore=rsScore("idscore")

cmdDC.CommandText = "sp_tarjeta " & idScore
cmdDC.CommandType = 1
Set rsTarjeta = Server.CreateObject("ADODB.Recordset")
rsTarjeta.Open cmdDC, , 0, 1

cmdDC.CommandText = "sp_misc_tarjeta " & idScore
cmdDC.CommandType = 1
Set rsMisc = Server.CreateObject("ADODB.Recordset")
rsMisc.Open cmdDC, , 0, 1

cmdDC.CommandText = "sp_promedios " & idScore
cmdDC.CommandType = 1
Set rsPromedio = Server.CreateObject("ADODB.Recordset")
rsPromedio.Open cmdDC, , 0, 1

cmdDC.CommandText = "sp_resumen " & idScore
cmdDC.CommandType = 1
Set rsResumen = Server.CreateObject("ADODB.Recordset")
rsResumen.Open cmdDC, , 0, 1
YardasIda=0
ParIda=0
ScoreIda=0
YardasVta=0
ParVta=0
ScoreVta=0

Set Mail = Server.CreateObject("Persits.MailSender")    


Mail.Host = "mail.golfguide.com.ar" ' 
Mail.From =  "mail@golfguide.com.ar"  			
Mail.FromName = "Resultados Golf Guide"				
Mail.Username = "mail@golfguide.com.ar" ' Nombre de usuario de cuenta de mail
Mail.Password = "xxxxxxxx" ' Contraseia


Mail.AddAddress Right(Linea,Len(Linea)-instr(Linea,","))

Mail.Subject = 	"ESTADISTICAS y RANKING - " & rsMisc.fields("Cancha")		
	
Mail.IsHTML = True 

armarmail()

Mail.Body=tx

Mail.Send

' response.write (Mail.Subject)

 response.write(tx)

end sub
'-------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------
function QueryRankingTop3(pcmdDC,pFecha,pIdRanking,pHcpDesde,pHcpHasta)

   '***
   'SI pHcpDesde ES MENOR E IGUAL A 0 TOMA TODOS LOS HCP MENORES A pHcpHasta
   dim filtroHcp
   dim rsRankingXX
   filtroHcp=" and UltHcp<=" & pHcpHasta & " "
   if pHcpDesde>0 then
       filtroHcp=" and UltHcp>=" & pHcpDesde & " " & filtroHcp
   end if
   '***
   
   pcmdDC.CommandText = "select top 3 auxranking.matricula,Nombre,Puntaje "&_
   "from auxranking,jugadores,fechas "&_
   "where auxranking.matricula=jugadores.matricula "&_
   "and auxranking.idranking=" & pIdRanking & " "&_
   "and fecha='" & Fecha & "' " & filtroHcp & " "&_
   "order by puntaje desc, jugados desc,Nombre"
   cmdDC.CommandType = 1
   Set rsRankingXX = Server.CreateObject("ADODB.Recordset")
   rsRankingXX.Open pcmdDC, , 0, 1

    set QueryRankingTop3 = rsRankingXX

end function

'-------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------
function QueryRankingMatricula(pcmdDC,pFecha,pIdRanking,pHcpDesde,pHcpHasta)
    'DUDAS: SI CAMBIA DE CATEGORIA???
    'COMO SABER A QUE CATEGORIA PERTENECE EN EL MOMENTO DE ESTA CONSULTA, �HISTORIAL? �PODRIA SER EL HCP ACTUAL?
    ' rsMisc.fields("Matricula") EN TEORIA DEBERIA SER VISIBLE EN ESTE AMBITO...


    dim filtroHcp
    dim rsRankingPos
    filtroHcp=" and UltHcp<=" & pHcpHasta & " "
    if pHcpDesde>0 then
        filtroHcp=" and UltHcp>=" & pHcpDesde & " " & filtroHcp
    end if
    dim QueryPosAct
    QueryPosAct= "select top 1 auxranking.matricula,Nombre,Puntaje, jugados "&_
    ", (select count(*)+1 from auxranking aux2 where aux2.idranking=" & pIdRanking & " " & filtroHcp & " and aux2.puntaje>auxranking.puntaje) ranking_posicion " &_
   " from auxranking,jugadores,fechas "&_
   " where auxranking.matricula=jugadores.matricula "&_
   " and auxranking.idranking=" & pIdRanking & " "&_
   " and fecha='" & Fecha & "' " & filtroHcp & " "&_
   " and auxranking.matricula=" & rsMisc.fields("Matricula") & " "&_
   " union select 0 matricula, 'sin ranking' Nombre,0 Puntaje, 0 jugados , 0 ranking_posicion order by puntaje desc, jugados desc,Nombre"
   
   
    'tx = tx & QueryPosAct
    pcmdDC.CommandText =QueryPosAct
   
   cmdDC.CommandType = 1
   Set rsRankingPos = Server.CreateObject("ADODB.Recordset")
   rsRankingPos.Open pcmdDC , , 0, 1
   

   set QueryRankingMatricula = rsRankingPos

end function

'-------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------

 %>
