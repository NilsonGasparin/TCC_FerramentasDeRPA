			#NoEnv  
			SendMode Input  
			SetWorkingDir %A_ScriptDir%  
				CoordMode, Pixel, Window
				CoordMode, Mouse, Window
				SetTitleMatchMode, 2
				SetTitleMatchMode, RegEx
				Ct = 1
				Dtf= 20210512
				Def = Default
				
				
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
					ART := []
					;xProgress := A_ScreenWidth - 350
					;yProgress := A_ScreenHeight - 60
					aux = 0
					
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
							UC[Indice] := XL.Sheets(1).Range("B" . Linha).text ; Equipe
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
							Pot[Indice] := XL.Sheets(1).Range("M" . Linha).text ;Data de fim
							Tensao[Indice] := XL.Sheets(1).Range("N" . Linha).text ;Data de fim
							Tipo[Indice] := XL.Sheets(1).Range("O" . Linha).text ;Data de fim
							Ramal[Indice] := XL.Sheets(1).Range("P" . Linha).text ;Data de fim
							Remoto[Indice] := XL.Sheets(1).Range("Q" . Linha).text ;Data de fim
							DJ[Indice] := XL.Sheets(1).Range("R" . Linha).text ;Data de fim
							Fuso[Indice] := XL.Sheets(1).Range("T" . Linha).text ;Data de fim
							Longitude[Indice] := XL.Sheets(1).Range("U" . Linha).text ;Data de fim
							Latitude[Indice] := XL.Sheets(1).Range("V" . Linha).text ;Data de fim
							Ligacao[Indice] := XL.Sheets(1).Range("W" . Linha).text ;Data de fim
							Zona[Indice] := XL.Sheets(1).Range("Z" . Linha).text ;Data de fim
							PG[Indice] := XL.Sheets(1).Range("AA" . Linha).text ;Data de fim
							Ncabos[Indice] := XL.Sheets(1).Range("AB" . Linha).text ;Data de fim
							Fase[Indice] := XL.Sheets(1).Range("AC" . Linha).text ;Data de fim
							Neutro[Indice] := XL.Sheets(1).Range("AD" . Linha).text ;Data de fim
							Terra[Indice] := XL.Sheets(1).Range("AE" . Linha).text ;Data de fim
							Nplaca[Indice] := XL.Sheets(1).Range("AF" . Linha).text ;Data de fim
							Placa[Indice] := XL.Sheets(1).Range("AG" . Linha).text ;Data de fim
							Marca_Placa[Indice] := XL.Sheets(1).Range("AH" . Linha).text ;Data de fim
							Area[Indice] := XL.Sheets(1).Range("AJ" . Linha).text ;Data de fim
							PotMod[Indice] := XL.Sheets(1).Range("AI" . Linha).text ;Data de fim
							Ninv[Indice] := XL.Sheets(1).Range("AK" . Linha).text ;Data de fim
							Inv[Indice] := XL.Sheets(1).Range("AL" . Linha).text ;Data de fim
							Marca_Inv[Indice] := XL.Sheets(1).Range("AM" . Linha).text ;Data de fim
							PInv[Indice] := XL.Sheets(1).Range("AN" . Linha).text ;Data de fim
							DJCA[Indice] := XL.Sheets(1).Range("AP" . Linha).text ;Data de fim
							MiniouMic[Indice] := XL.Sheets(1).Range("AR" . Linha).text ;Data de fim
							Acoplado_Trafo[Indice]:=  XL.Sheets(1).Range("AS" . Linha).text ;Data de fim
							Atendimento_Trafo[Indice]:=  XL.Sheets(1).Range("AT" . Linha).text ;Data de fim
							Linha_Inv[Indice]:=  XL.Sheets(1).Range("AX" . Linha).text ;Data de fim
							Linha_Placa[Indice]:=  XL.Sheets(1).Range("AV" . Linha).text ;Data de fim
							Linha_Marca_Inv[Indice]:=  XL.Sheets(1).Range("AW" . Linha).text ;Data de fim
							Linha_Marca_Placa[Indice]:=  XL.Sheets(1).Range("AU" . Linha).text ;Data de fim
							Index_Cidade[Indice]:=  XL.Sheets(1).Range("AY" . Linha).text ;Data de fim
							Index_Tensão[Indice]:=  XL.Sheets(1).Range("AZ" . Linha).text ;Data de fim
							Index_Fuso[Indice]:=  XL.Sheets(1).Range("BA" . Linha).text ;Data de fim
							Index_Trafo_Atendimento[Indice]:=  XL.Sheets(1).Range("BC" . Linha).text ;Data de fim
							ART[Indice]:=  XL.Sheets(1).Range("BE" . Linha).text ;Data de fim
							aux1 := aux1 + 1
						} Until (aux1 := IndiceTotal )
					
					XL.quit
					
					Indice := aux2[aux1]
					
					loop1:
					
					ToolTip, % "Progresso Atual: " . aux1 . " / " . IndiceTotal, 1600, 1020
						
						Sleep, 1000
							
							MouseMove, 525, 300
							
							Sleep, 500
							Click, 1
							Sleep, 300
							
						If( MiniouMic[Indice] = "Micro" ){
							sleep, 500
							aux := 11
							while(aux>0) {
								Send,{Down}
								sleep, 50
								aux := aux - 1
							}
						}
						else if( MiniouMic[Indice] = "Mini" ){
							sleep, 500
							aux := 11
							while(aux>0) {
								Send,{Down}
								sleep, 50
								aux := aux - 1
							}
						}
						else if(MiniouMic[Indice] = "0"){
							MsgBox % " Teste                                                          " . Nome[Indice] . "   " . MiniouMic[Indice]  .  "      Mini"
							pause
						}
						else {
							MsgBox % "O valor da celula Não condiz com o esperado " . Nome[Indice] . "   " . MiniouMic[Indice]  .  "      Mini"
							pause
					}
						
						Sleep, 500
						Send, {Enter}
						Sleep, 1000
						Send, {Enter}
						Sleep, 500
						Send, {Tab}
						Sleep, 300
						aux := PInv[Indice] *Ninv[Indice]
						SendInput, % aux
						sleep, 300
						Send, {Tab}
						
						sleep, 300
						Send, {Down}
						sleep, 300
						Send, {Tab}
						
						sleep, 300
						Send, {Down}
						sleep, 300
						Send, {Down}
						sleep, 300
						Send, {Tab}
						
						SendInput, % ART[Indice]
						sleep, 300
						Send, {Tab}
						sleep, 300
						SendInput, % Nome[Indice]
						sleep, 300
						Send, {Tab}
						sleep, 300
						Send, {Tab}
						sleep, 400
						
						
						SendInput, % Cidade[Indice]
						sleep, 500
						
						;~ if(Cidade[Indice] = "Campo Grande" )
						;~ {
							;~ Send, {Down}
							;~ sleep, 300
						;~ }
						;~ else if(Cidade[Indice] = "Jardim" )
						;~ {
							;~ Send, {Down}
							;~ sleep, 300
							;~ Send, {Down}
							;~ sleep, 300
							
						;~ }
						;~ else if(Cidade[Indice] = "Dourados" )
						;~ {
							;~ Send, {Down}
							;~ sleep, 300

							
						;~ }
						;~ else if(Cidade[Indice] = "Guia Lopes da Laguna" )
						;~ {
							;~ sleep, 300
						;~ }
						;~ else{
							;~ MsgBox % "Seleção de cidade encontra-se inativo, Tava com preguiça de programar e não fiz essa parte. Para continuar o procedimento basta selecionar a cidade e clicar em ESC" . Cidade[Indice] 
							;~ pause
						;~ }
						
						Send, {Tab}
							Sleep, 300
							
						if(Zona[Indice] = "URBANO" )
						{
							Send, {Down}
							sleep, 300
							Send, {Down}
							sleep, 300
						}
						else if(Zona[Indice] = "RURAL" )
						{
							Send, {Down}
							sleep, 300
						}
						else{
							MsgBox % "Seleção de cidade encontra-se inativo, Tava com preguiça de programar e não fiz essa parte. Para continuar o procedimento basta selecionar a cidade e clicar em ESC" . Cidade[Indice] 
							pause
						}
						

						Send, {Tab}
						
						Sleep, 300
						SendInput, % Bairro[Indice]
						Sleep, 300
						Send, {Tab}
						
						sleep, 300
						SendInput, % Rua[Indice]
						Sleep, 300
						Send, {Tab}
						
						sleep, 300
						SendInput, % Num[Indice]
						Sleep, 300
						Send, {Tab}
						sleep, 300
						
						sleep, 300
						SendInput, % CEP[Indice]
						Sleep, 300
						Send, {Tab}
						

						Sleep, 300
						Send, {Tab}
						
						sleep, 300
						SendInput, % Nome[Indice]
						sleep, 300
						Send, {Tab}
						sleep, 300
						
						if(StrLen(CPF[Indice]) = 11){
						
						Send, {Down}
							sleep, 300
						
						}
						else if(StrLen(CPF[Indice]) = 14)
						{
							Send, {Down}
							sleep, 300
							Send, {Down}
							sleep, 300
						}
						else{
							MsgBox % "Verificar CPF/CNPJ" . CPF[Indice]
							pause
						}
						
						Send, {Tab}
						
						sleep, 300
						SendInput, % CPF[Indice]
						Sleep, 300
						Send, {Tab}
						
						sleep, 300
						SendInput, % Tel[Indice]
						Sleep, 300
						Send, {Tab}
						
						;GoSub, Is_It_OK
						sleep, 300
						SendInput, % Cell[Indice]
						Sleep, 300
						Send, {Tab}
						
						Sleep, 300
						SendInput, % Email[Indice]
						Sleep, 300
						Send, {Tab}
						
						sleep, 500
							aux := 12
							while(aux>0) {
								Send,{Down}
								sleep, 50
								aux := aux - 1
							}
						Sleep, 300
						Send, {Tab}
						
						sleep, 400
						
						SendInput, % Cidade[Indice]
						sleep, 500
						
						;~ if(Cidade[Indice] = "Campo Grande" )
						;~ {
							;~ sleep, 500
							;~ aux := 20
							;~ while(aux>0) {
								;~ Send,{Down}
								;~ sleep, 50
								;~ aux := aux - 1
							;~ }
						;~ }
						;~ else if(Cidade[Indice] = "Jardim" )
						;~ {
							;~ sleep, 500
							;~ aux := 44
							;~ while(aux>0) {
								;~ Send,{Down}
								;~ sleep, 50
								;~ aux := aux - 1
							;~ }
							
						;~ }
						;~ else if(Cidade[Indice] = "Dourados" )
						;~ {
							;~ sleep, 500
							;~ aux := 32
							;~ while(aux>0) {
								;~ Send,{Down}
								;~ sleep, 50
								;~ aux := aux - 1
							;~ }
							
						;~ }
						;~ else if(Cidade[Indice] = "Guia Lopes da Laguna" )
						;~ {
							;~ sleep, 500
							;~ aux := 36
							;~ while(aux>0) {
								;~ Send,{Down}
								;~ sleep, 50
								;~ aux := aux - 1
							;~ }
						;~ }
						;~ else{
							;~ MsgBox % "Seleção de cidade encontra-se inativo, Tava com preguiça de programar e não fiz essa parte. Para continuar o procedimento basta selecionar a cidade e clicar em ESC" . Cidade[Indice] 
							;~ pause
						;~ }
						
						Send, {Tab}
						Sleep, 300
							
						sleep, 300
						SendInput, % CEP[Indice]
						Sleep, 300
						Send, {Tab}
						
						sleep, 300
						SendInput, % Rua[Indice]
						Sleep, 300
						Send, {Tab}
						
						Sleep, 300
						SendInput, % Bairro[Indice]
						Sleep, 300
						Send, {Tab}
						
						sleep, 300
						SendInput, % Num[Indice]
						Sleep, 300
						Send, {Tab}
						
						sleep, 300
						Send, 64634
						Sleep, 300
						Send, {Tab}
						
						sleep, 300
						Send, (67) 3201-6459
						Sleep, 300
						Send, {Tab}
						
						sleep, 300
						Send,(67) 9 8167-0391
						Sleep, 300
						Send, {Tab}
						
						sleep, 300
						Send,engenharia@montecristoms.com.br
						Sleep, 300
						Send, {Tab}
						
						sleep, 500
						aux := 12
						while(aux>0) {
								Send,{Down}
								sleep, 50
								aux := aux - 1
						}
						Sleep, 300
						Send, {Tab}
						
						sleep, 300
						aux := 20
						while(aux>0) {
								Send,{Down}
								sleep, 50
								aux := aux - 1
							}
						
						Sleep, 300
						Send, {Tab}
						
						sleep, 300
						Send,79004-050
						Sleep, 300
						Send, {Tab}
						
						sleep, 300
						Send,RUA ALUIZIO DE AZEVEDO
						Sleep, 300
						Send, {Tab}
						
						sleep, 300
						Send,JARDIM MONTE LIBANO
						Sleep, 300
						Send, {Tab}
						
						sleep, 300
						Send,616
						Sleep, 300
						Send, {Tab}
						
						sleep, 300
						;~ Send,{Enter}
						;~ Sleep, 300

						
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
				
				sleep, 3000
				
				
				
				
				