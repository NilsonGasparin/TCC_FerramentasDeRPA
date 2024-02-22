			#NoEnv  
			SendMode Input  
			SetWorkingDir %A_ScriptDir%  
				CoordMode, Pixel, Window
				CoordMode, Mouse, Window
				SetTitleMatchMode, 2
				SetTitleMatchMode, RegEx
				Ct = 1
				Dtf= 20221022
				Def = Default
				
				Test_Zone:
				
						Sleep, 175
						FormatTime, dt, %A_Now%, ddMMyyyy
						Sleep, 175
						dt1 := dt +  01000000
						dt1 = 0%dt1%
						Sleep, 175
						;~ dt2 := dt +  00050000
						;~ dt2 = "0".%dt1%
						dt2 := 05012023
				goto License
				
				
				Inicio:

				;Variaveis
					Nome:= []
					UC:= []
					Classe := []
					Rua := []
					Num := []
					CEP := []
					Bairro := []
					Cidade := []
					Email:= []
					Tel := []
					Cell := []
					CPF := []
					Pot:= []
					Tensao:= []
					Tipo:= []
					Ramal:= []
					Remoto:= []
					DJ:= []
					Fuso:= []
					Longitude:= []
					Latitude:= []
					Ligacao:= []
					Zona:= []
					PG:= []
					Ncabos:= []
					Fase:= []
					Neutro:= []
					Terra:= []
					Nplaca:= []
					Placa:= []
					Area:= []
					PotMod:= []
					Ninv:= []
					Inv:= []
					PInv:= []
					DJCA:= []
					MiniouMic:= []
					Atendimento_Trafo:= []
					Acoplado_Trafo:= []
					Marca_Placa:= []
					Marca_Inv:= []
					Nome_Comp:= []
					UC_Comp:= []
					End_Comp:= []
					Perc_Comp:= []
					CPF_Comp:= []
					Linha_Inv:= []
					Linha_Placa:= []
					Linha_Marca_Inv:= []
					Linha_Marca_Placa:= []
					Index_Cidade:=[]
					Index_Tensão:=[]
					Index_Fuso:=[]
					Index_Trafo_Atendimento:=[]
					;xProgress := A_ScreenWidth - 350
					;yProgress := A_ScreenHeight - 60
					
					info:
					;Resetar linha e indice
					InputBox, aux1, Estagiário do Estagiário, Informe o ESTAGIÁRIO quantos formularios deseja preencher, , 375, 200, 800, 600, , , Default
					if(ErrorLevel = 1){
					goto, end
					}
					else if(ErrorLevel = 0){
					if ( aux1 = 0 ){
					aux1 = 1
					}
					Else{
					goto, inserir_linhas
					}
					}
					
					inserir_linhas:
					aux2 :=[]
					IndiceTotal := aux1
					aux1 :=1
					While(aux1 <= IndiceTotal){
					
					InputBox, Linha , % "Estagiário do Estagiário - Progresso Atual:  " . aux1 . " / " . IndiceTotal, Informe os ESTAGIÁRIO qual linha da planilha deseja preencher o formulario, , 375, 200, 800, 600, , , Default
					if(ErrorLevel = 1){
						goto, end
					}
					else if(ErrorLevel = 0){
					aux2[aux1] := Linha - 1
					}
					aux1 := aux1 + 1
					}
					
					aux1:=1
					
					;Criar objeto para o arquivo a ser trabalhado e inicializar variáveis
					XL := ComObjCreate("Excel.Application") ;Criar objeto da aplicação Excel
					XL.DisplayAlerts := false
					XL.Workbooks.Open( "C:\Users\Nilson Gasparin\Dropbox\MONTE CRISTO\AUTOMAÇÃO\Dados - Formulario de Solicitação de Acesso.xlsm")
					Salvar := A_WorkingDir . "\" ;Variável com caminho para Salvar dados gerais
					
				;Capturar valores necessários do Excel
				
					Loop 
						{
							Indice := aux2[aux1]
							Linha := Indice + 1
							Nome[Indice] := XL.Sheets(1).Range("A" . Linha).text ;Equipamentos
							;~ UC[Indice] := XL.Sheets(1).Range("B" . Linha).text ; Equipe
							Classe[Indice]:= XL.sheets(1).Range("C" . Linha).text ;Numero da OS
							Rua[Indice]:= XL.Sheets(1).Range("D" . Linha).text ; Ano da OS
							Num[Indice]:=XL.Sheets(1).Range("E" . Linha).text ;Equipe a ser despachada
							CEP[Indice] := XL.Sheets(1).Range("F" . Linha).text ;Data de inicio
							Bairro[Indice] := XL.Sheets(1).Range("G" . Linha).text ;Data de fim
							Cidade[Indice] := XL.Sheets(1).Range("H" . Linha).text ;Data de fim
							Email[Indice] := XL.Sheets(1).Range("I" . Linha).text ;Data de fim
							Tel[Indice] := XL.Sheets(1).Range("J" . Linha).text ;Data de fim
							Cell[Indice] := XL.Sheets(1).Range("K" . Linha).text ;Data de fim
							CPF[Indice] := XL.Sheets(1).Range("L" . Linha).text ;Data de fim
							;~ Pot[Indice] := XL.Sheets(1).Range("M" . Linha).text ;Data de fim
							;~ Tensao[Indice] := XL.Sheets(1).Range("N" . Linha).text ;Data de fim
							;~ Tipo[Indice] := XL.Sheets(1).Range("O" . Linha).text ;Data de fim
							;~ Ramal[Indice] := XL.Sheets(1).Range("P" . Linha).text ;Data de fim
							;~ Remoto[Indice] := XL.Sheets(1).Range("Q" . Linha).text ;Data de fim
							;~ DJ[Indice] := XL.Sheets(1).Range("R" . Linha).text ;Data de fim
							;~ Fuso[Indice] := XL.Sheets(1).Range("T" . Linha).text ;Data de fim
							;~ Longitude[Indice] := XL.Sheets(1).Range("U" . Linha).text ;Data de fim
							;~ Latitude[Indice] := XL.Sheets(1).Range("V" . Linha).text ;Data de fim
							;~ Ligacao[Indice] := XL.Sheets(1).Range("W" . Linha).text ;Data de fim
							;~ Zona[Indice] := XL.Sheets(1).Range("Z" . Linha).text ;Data de fim
							PG[Indice] := XL.Sheets(1).Range("AA" . Linha).text ;Data de fim
							;~ Ncabos[Indice] := XL.Sheets(1).Range("AB" . Linha).text ;Data de fim
							;~ Fase[Indice] := XL.Sheets(1).Range("AC" . Linha).text ;Data de fim
							;~ Neutro[Indice] := XL.Sheets(1).Range("AD" . Linha).text ;Data de fim
							;~ Terra[Indice] := XL.Sheets(1).Range("AE" . Linha).text ;Data de fim
							;~ Nplaca[Indice] := XL.Sheets(1).Range("AF" . Linha).text ;Data de fim
							;~ Placa[Indice] := XL.Sheets(1).Range("AG" . Linha).text ;Data de fim
							;~ Marca_Placa[Indice] := XL.Sheets(1).Range("AH" . Linha).text ;Data de fim
							;~ Area[Indice] := XL.Sheets(1).Range("AJ" . Linha).text ;Data de fim
							;~ PotMod[Indice] := XL.Sheets(1).Range("AI" . Linha).text ;Data de fim
							;~ Ninv[Indice] := XL.Sheets(1).Range("AK" . Linha).text ;Data de fim
							;~ Inv[Indice] := XL.Sheets(1).Range("AL" . Linha).text ;Data de fim
							;~ Marca_Inv[Indice] := XL.Sheets(1).Range("AM" . Linha).text ;Data de fim
							;~ PInv[Indice] := XL.Sheets(1).Range("AN" . Linha).text ;Data de fim
							;~ DJCA[Indice] := XL.Sheets(1).Range("AP" . Linha).text ;Data de fim
							;~ MiniouMic[Indice] := XL.Sheets(1).Range("AR" . Linha).text ;Data de fim
							;~ Acoplado_Trafo[Indice]:=  XL.Sheets(1).Range("AS" . Linha).text ;Data de fim
							;~ Atendimento_Trafo[Indice]:=  XL.Sheets(1).Range("AT" . Linha).text ;Data de fim
							;~ Linha_Inv[Indice]:=  XL.Sheets(1).Range("AX" . Linha).text ;Data de fim
							;~ Linha_Placa[Indice]:=  XL.Sheets(1).Range("AV" . Linha).text ;Data de fim
							;~ Linha_Marca_Inv[Indice]:=  XL.Sheets(1).Range("AW" . Linha).text ;Data de fim
							;~ Linha_Marca_Placa[Indice]:=  XL.Sheets(1).Range("AU" . Linha).text ;Data de fim
							;~ Index_Cidade[Indice]:=  XL.Sheets(1).Range("AY" . Linha).text ;Data de fim
							;~ Index_Tensão[Indice]:=  XL.Sheets(1).Range("AZ" . Linha).text ;Data de fim
							;~ Index_Fuso[Indice]:=  XL.Sheets(1).Range("BA" . Linha).text ;Data de fim
							;~ Index_Trafo_Atendimento[Indice]:=  XL.Sheets(1).Range("BC" . Linha).text ;Data de fim
							aux1 := aux1 + 1
						} Until (aux1 := IndiceTotal )
					
					XL.quit
					
					Indice := aux2[aux1]
					loop1:
						
						
					ToolTip, % "Progresso Atual: " . aux1 . " / " . IndiceTotal, 1780, 1020
					
						Run, https://ecrea.creams.org.br/Autenticacao/Login
						WinActivate, CREA - Conselho Regional de Engenharia e Agronomia 
						WinWaitActive,CREA - Conselho Regional de Engenharia e Agronomia 
						Sleep, 700
						
						
						1:
						
						ImageSearch, X , Y , 1000, 300, 1700, 800, %A_WorkingDir%\lib\1.PNG
						
						if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 1
						
						}
						
						else if( ErrorLevel = 0){
							
						}
						
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Enter}
						Sleep, 1000
						
						2:
						
						;~ ImageSearch, X , Y , 300, 300, 1500, 800, %A_WorkingDir%\lib\2.PNG
						
						;~ if(ErrorLevel = 1){
							
							;~ Sleep, 175
							;~ goto, 2
						
						;~ }
						
						;~ else if( ErrorLevel = 0){
							
						;~ }
					
						;~ Send, {Tab}
						;~ Sleep, 250
						;~ Send, {Enter}
						;~ Sleep, 250
						;~ Send, {Enter}
						;~ Sleep, 600
						;~ Send, {Tab}
						;~ Sleep, 250
						;~ Send, {Tab}
						;~ Sleep, 250
						;~ Send, {Enter}
						;~ Sleep, 500
						
						3:
						
						ImageSearch, X , Y , 300, 300, 1500, 800, %A_WorkingDir%\lib\3.PNG
						
						if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 3
						
						}
						
						else if( ErrorLevel = 0){
							
						}
						
						MouseMove, 860,60
						Sleep, 175
						Click, 1
						Sleep, 175
						Send, ^a
						Sleep, 175
						Send, https://ecrea.creams.org.br/Art
						Sleep, 175
						Send, {Enter}
						
						4:
						
						ImageSearch, X , Y , 300, 300, 1500, 800, %A_WorkingDir%\lib\4.PNG
						
						if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 4
						
						}
						
						else if( ErrorLevel = 0){
							
						}
						
						MouseMove, 470,752
						Sleep, 175
						Click, 1
						Sleep, 175
						
						5:
						
						ImageSearch, X , Y , 300, 300, 1500, 800, %A_WorkingDir%\lib\5.PNG
						
						if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 5
						
						}
						
						else if( ErrorLevel = 0){
							
						}
						
						MouseMove, 650,540
						Sleep, 175
						Click, 1
						Sleep, 175
						Send, INDIVIDUAL
						Sleep, 175
						Send, {Enter}
						Sleep, 300
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Enter}
						Sleep, 1000
						
						6:
						
						ImageSearch, X , Y ,  0, 0, 800, 800,  %A_WorkingDir%\lib\6.PNG
						
						if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 6
						
						}
						
						else if( ErrorLevel = 0){
							
						}
						
						
						MouseMove, 1000,300
						Sleep, 175
						Click, 1
						Sleep, 175
						Send, 35.129
						Sleep, 175
						Send, {Enter}
						Sleep, 300
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Enter}
						Sleep, 1000
						
						
						7:
						
						ImageSearch, X , Y ,  0, 0, 800, 800, %A_WorkingDir%\lib\7.PNG
						
						if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 7
						
						}
						
						else if( ErrorLevel = 0){
							
						}
						
						
						MouseMove, 700,300
						Sleep, 175
						Click, 1
						Sleep, 175
						StringLen, aux,  % CPF[Indice]
						Sleep, 175
						if(aux = 11){
							
						}
						else if(aux = 14 ) {
							Sleep, 175
								Send,PESSOA JURIDICA DE DIREITO PRIVADO
								Sleep, 100
								
							
						}
						else{
							MsgBox % "O valor da celula Não condiz com o esperado " . CPF[Indice] 

						}
						Send, {Enter}
						Sleep, 300
						Send, {Tab}
						Sleep, 175
						SendInput, % CPF[Indice]
						Sleep, 300
						Send, {Tab}
						Sleep, 500
						
						troubleshoot1:
						
						ImageSearch, X , Y ,  1470, 300, 1700, 340, %A_WorkingDir%\lib\troubleshoot1.PNG
						
						if(ErrorLevel = 1){
						
						ImageSearch, X , Y , 0, 400, 1250, 1080, %A_WorkingDir%\lib\troubleshoot2.PNG
						if(ErrorLevel = 1){
							
							Sleep, 175
							goto, troubleshoot1
						
						}
						
						else if( ErrorLevel = 0){
							
							MouseMove, % X + 20, % Y + 20
							sleep, 150
							Click, 1
							Sleep, 150 
							goto, troubleshoot3
						
						}
						
						}
						
						else if( ErrorLevel = 0){
							goto,troubleshoot2
						}
						troubleshoot3:
						
						Sleep, 200
						ImageSearch, X , Y , 1300, 150, 1900, 600, %A_WorkingDir%\lib\troubleshoot3.PNG
						
						if(ErrorLevel = 1){
							
							Sleep, 175
							goto, troubleshoot1
						
						}
						
						else if( ErrorLevel = 0){
							
							MouseMove, 1500, 300
							sleep, 150
							Click, 1
							Sleep, 150 
						
						}
						
						troubleshoot2:
						
						SendInput, % Nome[Indice]
						Sleep, 300
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 300
						SendInput, % CEP[Indice]
						Sleep, 1000
						Send, {Tab}
						Sleep, 175
						Send, {Enter}
						Sleep, 1000
						Send, {Tab}
						Sleep, 500
						Send, {Tab}
						Sleep, 500
						Send, {Tab}
						Sleep, 175
						SendInput, % Rua[Indice]
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						SendInput, % Num[Indice]
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						SendInput, % Bairro[Indice]
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Enter}
						Sleep, 500
						
						8:
						
						ImageSearch, X , Y ,  0, 0, 800, 800, %A_WorkingDir%\lib\8.PNG
						
						if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 8
						
						}
						
						else if( ErrorLevel = 0){
							
						}
						
						
						MouseMove, 580,300
						Sleep, 175
						Click, 1
						Sleep, 175
						FormatTime, dt, %A_Now%, ddMMyyyy
						Sleep, 175
						SendInput, % dt
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						dt1 := dt +  01000000
						aux := StrLen(dt1)
						if(aux=7){
						dt1 = 0%dt1%
						}
						Sleep, 175
						SendInput, % dt1
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						dt2 := dt +  00050000
						 if(aux=7){						
						 dt2 = 0%dt2%
						 }
						Sleep, 175
						;~dt2 := 05052023
						SendInput, % dt2
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, 200000
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Enter}
						Sleep, 200
						Send, {Tab}
						Sleep, 200
						Send, {Tab}
						Sleep, 200
						Send, {Tab}
						Sleep, 200
						Send, {Tab}
						Sleep, 200
						Send, {Tab}
						Sleep, 200
						Send, {Tab}
						Sleep, 200
						Send, {Tab}
						Sleep, 175
						Send, {Enter}
						Sleep, 200
						
						9:
						
						Sleep, 175
						aux = 3
						while(aux>0){
						Send {WheelDown}
						sleep, 50
						aux := aux - 1
						}
						Sleep, 175
						
						ImageSearch, X , Y , 300, 300, 900, 1080, %A_WorkingDir%\lib\9.PNG
						
						if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 9
						
						}
						
						else if( ErrorLevel = 0){
							
							MouseMove, % X + 20, % Y + 20
							sleep, 150
							Click, 1
							Sleep, 150 
						
						}
						
						10:
						
						Sleep, 175
						aux = 3
						while(aux>0){
						Send {WheelDown}
						sleep, 50
						aux := aux - 1
						}
						Sleep, 175
						
						ImageSearch, X , Y , 0, 600, 1250, 1080, %A_WorkingDir%\lib\10.PNG
						
						if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 10
						
						}
						
						else if( ErrorLevel = 0){
							
							MouseMove, % X + 20, % Y + 20
							sleep, 150
							Click, 1
							Sleep, 150 
						
						}
						
						11:
						
						ImageSearch, X , Y ,  0, 0, 800, 800, %A_WorkingDir%\lib\11.PNG
						
						if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 11
						
						}
						
						else if( ErrorLevel = 0){
							
						}
						
						
						MouseMove, 715,380
						Sleep, 175
						Click, 1
						Sleep, 175
						Send, EXECU
						Sleep, 175
						Send, {Enter}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Enter}
						Sleep, 175
						Send, EXECUÇÂO DE OBRA
						Sleep, 175
						Send, {Enter}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Enter}
						Sleep, 300
						Send, DE MICROGERAÇÃO
						Sleep, 1500
						Send, {Enter}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						
						
						
						;~ --------------------------------------------------------------------------------------------------------------------
						;~ ----------------- SANTUARIO DE ODIO AO CREA MS --------------------------------------------------------
						;~ -------------------------------------------------------------------------------------------------------------------
						 
						 ;~ -------              You have no power over me               -------
						 
						 ;~ Jogo de cadeiras para trasnformar STRING em FLOAT em Integer
						 
					    aux5 :=  PG[Indice]  
						aux5 := StrReplace(aux5, "," , ".")
						aux5 := 10000* aux5
						
						;~ Test Case Spammer
						
						
						;~ Send,1
						;~ Sleep, 50
						;~ Send,{Enter}
						;~ Sleep, 50
						;~ Send, % aux1
						;~ Sleep, 50
						;~ Send,{Enter}
						;~ Sleep, 50
						;~ Send,2
						;~ Sleep, 50
						;~ Send,{Enter}
						;~ Sleep, 50
						;~ Send, % PG[Indice]
						;~ Sleep, 50
						;~ Send,{Enter}
						;~ Sleep, 50
						
						;~ Verificando numero de laços necessarios
						aux4:=0
						aux3:=aux5
					
						while(aux3>1000){
						aux3 := aux3/10
						aux4+=1
						}
						
						;~ Here comes the sun
						
						Aux := PG[Indice] 
						Aux := StrSplit(aux5)
						aux4 := aux4 + 4
						aux3 := 0
						While(aux3<aux4){
						Send, % Aux[aux3]
						aux3+=1
						Sleep, 100
						}
						
						;~ DO NOT LOOK BACK ..... JUST WALK
						
						
						Sleep, 175
						Send, {Tab}
						Sleep, 500
						Send, {Enter}
						Sleep, 175
						Send, QUILOWATT
						Sleep, 500
						Send, {Enter}
						Sleep, 175
						Send, {Tab}
						Sleep, 500
						Send, {Enter}
						Sleep, 400
						
						11a:
						
						ImageSearch, X , Y ,  300,300, 800, 800, %A_WorkingDir%\lib\11a.PNG
						
						if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 11a
						
						}
						
						else if( ErrorLevel = 0){
							
						}
						
						
						MouseMove, 715,380
						Sleep, 175
						Click, 1
						Sleep, 175
						Send, EXECU
						Sleep, 175
						Send, {Enter}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Enter}
						Sleep, 175
						Send, PROJETO
						Sleep, 175
						Send, 	{Down}
						Sleep, 175
						Send, 	{Down}
						Sleep, 175
						Send, {Enter}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, {Enter}
						Sleep, 175
						Send, DE MICROGERAÇÃO
						Sleep, 1500
						Send, {Enter}
						Sleep, 175
						Send, {Tab}
						Sleep, 150
						;~ MsgBox, 0 , SITE FILHA DA PUTA DE UM CARALHO - EXXXTA MEEERRRDA NÃO FUNCIONA, % "SÓ DIGIITA ESTA MERDA AQUI " . PG[Indice] ,
						;~ pause
						;~ aux :=  PG[Indice]
						;~ aux := 10000* aux
						;~ SendInput, % aux
						
						 
					    aux5 :=  PG[Indice]  
						aux5 := StrReplace(aux5, "," , ".")
						aux5 := 10000* aux5
						;~ Send,1
						;~ Sleep, 50
						;~ Send,{Enter}
						;~ Sleep, 50
						;~ Send, % aux1
						;~ Sleep, 50
						;~ Send,{Enter}
						;~ Sleep, 50
						;~ Send,2
						;~ Sleep, 50
						;~ Send,{Enter}
						;~ Sleep, 50
						;~ Send, % PG[Indice]
						;~ Sleep, 50
						;~ Send,{Enter}
						;~ Sleep, 50
						aux4:=0
						aux3:=aux5
						while(aux3>1000){
						aux3 := aux3/10
						aux4+=1
						}
						Aux := PG[Indice] 
						Aux := StrSplit(aux5)
						aux4 := aux4 + 4
						aux3 := 0
						While(aux3<aux4){
						Send, % Aux[aux3]
						aux3+=1
						Sleep, 100
						}
						
						;~ SendInput, % aux
						
						Sleep, 175
						Send, {Tab}
						Sleep, 500
						Send, {Enter}
						Sleep, 175
						Send, QUILOWATT
						Sleep, 500
						Send, {Enter}
						Sleep, 175
						Send, {Tab}
						Sleep, 500
						Send, {Enter}
						Sleep, 2000
						
						
						
							12:
						
						ImageSearch, X , Y ,  400, 800, 900, 1000, %A_WorkingDir%\lib\12.PNG
						
							if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 12
						
						}
						
						else if( ErrorLevel = 0){
							
							MouseMove, % X + 20, % Y + 20
							sleep, 150
							Click, 1
							Sleep, 150 
						
						}
						
						13:
						
						ImageSearch, X , Y , 0, 0, 800, 1000, %A_WorkingDir%\lib\13.PNG
						
						if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 13
						
						}
						
						else if( ErrorLevel = 0){
							
						}
						
						
						14:
						
						ImageSearch, X , Y , 0, 300, 1000, 1000, %A_WorkingDir%\lib\14.PNG
						
							if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 14
						
						}
						
						else if( ErrorLevel = 0){
							
							MouseMove, % X + 20, % Y + 20
							sleep, 150
							Click, 1
							Sleep, 150 
						
						}
					
						15:
						
						ImageSearch, X , Y , 0, 0, 800, 1000, %A_WorkingDir%\lib\15.PNG
						
						if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 15
						
						}
						
						else if( ErrorLevel = 0){
							
						}
						
						
						16:
						
						ImageSearch, X , Y , 0, 300, 1000, 1000, %A_WorkingDir%\lib\16.PNG
						
							if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 16
						
						}
						
						else if( ErrorLevel = 0){
							
							MouseMove, % X + 20, % Y + 20
							sleep, 150
							Click, 1
							Sleep, 150 
						
						}
					
					pause
					
						18:
						
						ImageSearch, X , Y , 0, 0, 1280, 1000, %A_WorkingDir%\lib\18.PNG
						
							if(ErrorLevel = 1){
							
							Sleep, 175
							goto, 18
						
						}
						
						else if( ErrorLevel = 0){
							
							Sleep, 300
							MouseMove, % X + 20, % Y + 20
							sleep, 150
							Click, 1
							Sleep, 150 
						
						}
						
						MsgBox, 0 ,  NEXT, pressione esc para a proxima art, 5
						
						pause
						
					if(aux1 < IndiceTotal){
						aux1:= aux1  + 1
					Indice := aux2[aux1]
					goto, loop1
					}
					
					goto, end
					
					
	end:
				ExitApp
				
				Esc::Pause

				
	License:
	
				FormatTime, dt, %A_Now%, yyyyMMdd
				
				
				
				Envsub, dt, dtf,  days
				
				
				
				
				If(Ct = 1){ 
					
				MsgBox, 0 ,  Premmium - Estagiário do Estagiário, Contratando um Estagiario para me substituir, 5
				
				goto, Inicio
				}
			
				
				Else If(dt > 0){
					
					MsgBox, 0 , License Expired - Estagiario, Seu periodo de teste chegou ao final de 30 dias. Entre em contato com o desenvolvedor através do numero +55 (67) 9 98167-0391 e adquira agora mesmo a sua versão premium. ,
					
					goto, end
				
				}
			
				Else { 
					
				dt :=-dt
				MsgBox, 0 ,  Free Version - Estagiário do Estagiário, Contratando um Estagiario para me substituir. Voce possui ainda %dt% dias da versão gratuita. Entre em contato com o desenvolvedor através do numero +55 (67) 9 98167-0391 e adquira agora mesmo a sua versão premium., 5
				
				goto, Inicio
				}
				
				goto, end
				
				
				compensao:
				
				Sleep, 1750
				
				
				
				
				