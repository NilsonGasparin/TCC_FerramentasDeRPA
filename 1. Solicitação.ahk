			#NoEnv  
			SendMode Input  
			SetWorkingDir %A_ScriptDir%  
				CoordMode, Pixel, Window
				CoordMode, Mouse, Window
				SetTitleMatchMode, 2
				SetTitleMatchMode, RegEx
				Ct = 2
				Dtf= 20221022
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
					aux = 0
					
					info:
					
					aux = 0
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
							aux1 := aux1 + 1
						} Until (aux1 := IndiceTotal )
					
						
							
							Linha_Comp := 1
							Indice_Comp:= 0
						Loop 
							{
								Nome_Comp[Indice_Comp]:= XL.Sheets(2).Range("A" . Linha_Comp).text ;Equipamentos
								UC_Comp[Indice_Comp]:= XL.Sheets(2).Range("B" . Linha_Comp).text ; Equipe
								CPF_Comp[Indice_Comp]:= XL.Sheets(2).Range("C" . Linha_Comp).text ; Equipe
								End_Comp[Indice_Comp]:= XL.sheets(2).Range("D" . Linha_Comp).text ;Numero da OS
								Perc_Comp[Indice_Comp]:= XL.Sheets(2).Range("E" . Linha_Comp).text ; Ano da OS
								Indice_Comp := Indice_Comp + 1
								Linha_Comp := Linha_Comp + 1
							} Until (XL.Sheets(2).Range("A" . Linha_Comp).Value = "")
						
					XL.quit
						IndiceTotal_Comp :=Indice_Comp-1
		
					
					Indice := aux2[aux1]
					
					loop1:
					
					ToolTip, % "Progresso Atual: " . aux1 . " / " . IndiceTotal, 1600, 1020

						test  := WinExist(MEMORIAL_GD_v4-22022022.xlsm)
						if(test=0){
						}
						Else{
						MsgBox, 0 ,  Error, Feche o Excel para proseguir, 7
						pause
						goto, loop1
						}
						
						;Abrir tela para preenchimento do formulario
						Sleep, 175
						Run, "C:\Users\Nilson Gasparin\Dropbox\MONTE CRISTO\AUTOMAÇÃO\MEMORIAL_GD_v4-22022022.xlsm"
						Sleep, 1750
						;~ WinWait,MEMORIAL_GD_v4-22022022.xlsm - Excel
						;~ Sleep, 700
						;~ WinActivate, MEMORIAL_GD_v4-22022022.xlsm - Excel
						;~ WinWaitActive,MEMORIAL_GD_v4-22022022.xlsm - Excel
						;~ Sleep, 700
						WinWait, SEJA BEM-VINDO AO NOVO MEMORIAL DESCRITIVO PARA GERAÇÃO DISTRIBUÍDA DO GRUPO ENERGISA!
						Sleep, 700
						WinActivate, SEJA BEM-VINDO AO NOVO MEMORIAL DESCRITIVO PARA GERAÇÃO DISTRIBUÍDA DO GRUPO ENERGISA!
						WinWaitActive,SEJA BEM-VINDO AO NOVO MEMORIAL DESCRITIVO PARA GERAÇÃO DISTRIBUÍDA DO GRUPO ENERGISA!
						Sleep, 200
						MouseMove, 945, 15
						Sleep, 300
						Click, 1
						Sleep, 300
						WinWait, Microsoft Excel
						Sleep, 700
						WinActivate, Microsoft Excel
						WinWaitActive,Microsoft Excel
						Sleep, 200
						MouseMove, 380, 10
						Sleep, 200
						Click, 1
						Sleep, 500
						
						aux := 0
						
						0.1:
						ImageSearch, X , Y , 0, 0, 1920, 1080, %A_WorkingDir%\lib\UPDATE.PNG
						

						if(ErrorLevel = 1){
						
						aux := aux + 1
						
							if( aux = 50) {
								aux := 0
								goto, 1
							}
						
							Sleep, 300
							goto, 0.1
						
						}
						
						else if( ErrorLevel = 0){
							
						}
						
						sleep, 500
						WinActivate, Microsoft Excel
						WinActivate, Microsoft Excel
						Sleep, 200
						MouseMove, 380, 16
						Sleep, 300
						Click, 1
						Sleep, 4000 
						aux := 0
						1:
						
						ImageSearch, X , Y , 0, 0, 1000, 1000, %A_WorkingDir%\lib\atualizar.PNG
							
							
						if(ErrorLevel = 1){
							Sleep, 200
							goto, 1
						
						}
						
						else if( ErrorLevel = 0){
							
						}
						
							Sleep, 300
							Send, {ALT DOWN}{TAB}{ALT UP}
							sleep, 400
						
						2:
						
						ImageSearch, X , Y , 300, 300, 1920, 1080, %A_WorkingDir%\lib\comecar.PNG
						
						if(ErrorLevel = 1){
							Sleep, 200
							goto, 2
						
						}
						
						else if( ErrorLevel = 0){
							
						}
						
						MouseMove, % X + 20, % Y + 20
						sleep, 150
						Click, 1
						Sleep, 150 

						;preenchimento de dados da planilha
						
							;Tela 1
						
						;Clica e preenche o campo com o numero da UC
						
						micromini:
						Sleep, 300
						
						ImageSearch, X , Y , 0, 0, 1000, 1000, %A_WorkingDir%\lib\MiniMicro.PNG
						
						
							if(ErrorLevel = 1){
							
							Sleep, 175
							goto, micromini
						
						}
						
						else if( ErrorLevel = 0){
							
						}
						
						If( MiniouMic[Indice] = "Micro" ){
							MouseMove, 276, 294
						}
						else if( MiniouMic[Indice] = "Mini" ){
							MouseMove, 276, 320
						}
						else if(MiniouMic[Indice] = "0"){
							MsgBox % " O valor da celula não condiz com o esperado " . Nome[Indice] . "   " . MiniouMic[Indice]  .  "     "
							pause
						}
						else {
							MsgBox % "O valor da celula não condiz com o esperado " . Nome[Indice] . "   " . MiniouMic[Indice]  .  "     "

					}
						
						Sleep, 200
						Click, 1
						Sleep, 1500
						SendInput, % UC[Indice]
						Sleep, 175
						Send, {Tab}
						
						;GoSub, Is_It_OK
						
						;preenche o campo com a classe da UC
						Sleep, 175
						MouseMove, 854, 164
						Sleep, 175
						Click, 1
						Sleep, 300
						if(Classe[Indice] = "RESIDENCIAL" ) {
							Sleep, 500
							Send,{Down}
							Sleep, 500
							Send, {Enter}
						}
						else if(Classe[Indice] = "COMERCIAL" ) {
							
							Sleep, 175
							aux := 2
							Sleep, 175
							while(aux>0) {
								Sleep, 100
								Send,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
							
						}
						
						else if(Classe[Indice] = "INDUSTRIAL" )
						{
								Sleep, 175
							aux := 3
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						}
						else if(Classe[Indice] = "RURAL" )
						{
							Sleep, 175
							aux := 4
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						}
						else if(Classe[Indice] = "PODER PUBLICO " )
						{
							Sleep, 175
							aux := 5
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						}
						else if(Classe[Indice] = "SERVIÇO PUBLICO " )
						{
							Sleep, 175
							aux := 6
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						}
						else{
							MsgBox [,,Erro - Valor Despadronizado, A planilha não está padronizada conforme o estipulado: RESIDENCIAL / COMERCIAL / INDUSTRIAL /  RURAL / PODER PUBLICO / SERVIÇO PUBLICO.  Verifique a planilha e corrija o problema.  Caso queira continuar o procedimento basta selecionar o valor correto e clicar em "ESC"]
							pause
						}
						Sleep, 175
						Send, {Tab}
						;GoSub, Is_It_OK
						;preenche o campo com o nome do Solicitante
						Sleep, 175
						SendInput, % Nome[Indice]
						Sleep, 175
						Send, {Tab}
						
						;preenche o campo com a rua do cliente
						Sleep, 175
						SendInput, % Rua[Indice]
						Sleep, 175
						Send, {Tab}
						
						;preenche o campo com o numero da casa do cliente
						Sleep, 175
						SendInput, % Num[Indice]
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						IfWinExist , Microsoft Excel
							
						{
						WinActivate, Microsoft Excel
						Pause
						}
						
						;preenche o campo com o nome do bairro
						SendInput, % Bairro[Indice]
						Sleep, 175
						Send, {Tab}
						;GoSub, Is_It_OK
						;preenche o campo com a cidade do solicitante. Em virtude da pessima construção da planilha esse campo não pode ser completamente programado, restando ao operador do Script fazer a seleção manual do valor da cidade, em casos diferentes de Campo Grande
						Sleep, 175
						IfWinExist , Microsoft Excel
							
						{
						WinActivate, Microsoft Excel
						Pause
						}
						MouseMove, 725, 226
						Sleep, 175
						Click, 1
						Sleep, 175
							aux := 3
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
						}
						Sleep, 175
						Send, {Enter}
						Sleep, 175
						Send, {Tab}
						
						Sleep, 175
						SendInput, % CEP[Indice]
						Sleep, 175
						Send, {Tab}
						
						Sleep, 175
						SendInput, % Email[Indice]
						Sleep, 175
						Send, {Tab}
						
						MouseMove, 945, 245
						Sleep, 175
						Click, 1
						Sleep, 175
						IfWinExist , Microsoft Excel
							
						{
						WinActivate, Microsoft Excel
						Pause
						}
						if(Index_Cidade[Indice] > 1){
						aux := % Index_Cidade[Indice] -1
						Sleep, 175
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						}
						else{
							MsgBox % "Seleção de cidade encontra-se inativo, Tava com preguiça de programar e não fiz essa parte. Para continuar o procedimento basta selecionar a cidade e clicar em ESC" . Cidade[Indice] 
							pause
						}
						Sleep, 175
						Send, {Tab}
						
						Sleep, 175
						SendInput, % Tel[Indice]
						Sleep, 175
						Send, {Tab}
						
						Sleep, 175
						SendInput, % Cell[Indice]
						Sleep, 175
						Send, {Tab}
						
						;GoSub, Is_It_OK
						Sleep, 175
						SendInput, % CPF[Indice]
						Sleep, 175
						Send, {Tab}
						;GoSub, Is_It_OK
						
						Sleep, 175
						SendInput, % Pot[Indice]
						Sleep, 175
						Send, {Tab}
						
						;GoSub, Is_It_OK
						Sleep, 175
						MouseMove, 854, 326
						Sleep, 175
						Click, 1
						Sleep, 175
						aux := % Index_Tensão[Indice]
						while(aux>0) {
							SendInput,{Down}
							Sleep, 50
							aux := aux - 1
						}
						Sleep, 175
						Send, {Enter}
						Sleep, 175
						Send, {Tab}
						
						;GoSub, Is_It_OK
						Sleep, 175
						MouseMove, 445, 350
						Sleep, 175
						Click, 1
						Sleep, 175
						if(Tipo[Indice] = "MONOFASICO" )
						{
							Sleep, 175
							aux := 1
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						}
						else if(Tipo[Indice] = "BIFASICO" )
						{
							Sleep, 175
							aux := 2
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						}
						
						else if(Tipo[Indice] = "TRIFASICO" )
						{
							Sleep, 175
							aux := 3
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						}
						
						else{
							MsgBox [,,Erro - Valor Despadronizado, A planilha não está padronizada conforme o estipulado: MONOFASICO / BIFASICO / TRIFASICO .  Verifique a planilha e corrija o problema.  Caso queira continuar o procedimento basta selecionar o valor correto e clicar em "ESC"]
							pause
						}
						Sleep, 175
						Send, {Tab}
						
						;GoSub, Is_It_OK
						Sleep, 175
						MouseMove, 445, 375	
						Sleep, 175
						Click, 1
						Sleep, 175
						if(Ramal[Indice] = "AEREO" )
						{
							Sleep, 175
							aux := 1
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						}
						else if(Ramal[Indice] = "SUBTERRANEO" )
						{
							Sleep, 175
							aux := 2
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						}
						else{
							MsgBox [,,Erro - Valor Despadronizado, A planilha não está padronizada conforme o estipulado: AEREO / SUB .  Verifique a planilha e corrija o problema.  Caso queira continuar o procedimento basta selecionar o valor correto e clicar em "ESC"]
							pause
						}
						Sleep, 175
						Send, {Tab}
						
						;GoSub, Is_It_OK
						Sleep, 175
						Send, {Tab}
						
						
						Sleep, 175
						MouseMove, 854, 470	
						Sleep, 175
						Click, 1
						Sleep, 175
							aux := 1
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						Sleep, 175
						Send, {Tab}
						
						
						Sleep, 175
						MouseMove,  854, 512
						Sleep, 175
						Click, 1
						Sleep, 175
							aux := 1
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						Sleep, 175
						Send, {Tab}
						
						
						Sleep, 175
						MouseMove,  854, 555	
						Sleep, 175
						Click, 1
						Sleep, 175
							aux := 1
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						
						
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						MouseMove,  854, 590		
						Sleep, 175
						Click, 1
						Sleep, 175
							aux := 1
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						Sleep, 175
						Send, {Tab}
						
						
						Sleep, 175
						if(Remoto[Indice] = "SIM" )
						{
							MouseMove,  854, 660	
							Sleep, 175
							Click, 1
							Sleep, 175
							aux := 1
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
							
						}
						
						else if(Remoto[Indice] = "NAO" )
						{

						}
						
						else{
							MsgBox [,,Erro - Valor Despadronizado, A planilha não está padronizada conforme o estipulado: SIM / NAO .  Verifique a planilha e corrija o problema.  Caso queira continuar o procedimento basta selecionar o valor correto e clicar em "ESC"]
							pause
						}
						Sleep, 175
						Send, {Tab}
						
						Sleep, 175
							Send, {Tab}
							
						Sleep, 175
						Send, {Tab}
						
						
						
						
						Sleep, 175
						SendInput, % Nome[Indice]
						Sleep, 175
						Send, {Tab}
						;GoSub, Is_It_OK
						
						Sleep, 175
						SendInput, % Tel[Indice]
						Sleep, 175
						Send, {Tab}
						;GoSub, Is_It_OK
						
						Sleep, 175
						SendInput, % Email[Indice]
						Sleep, 175
						Send, {Tab}
						;GoSub, Is_It_OK
						
						Sleep, 175
						MouseMove, 995, 905
						Sleep, 175
						Click, 1
						Sleep, 800
						
						
						
							;Tela 2
						SendInput, % DJ[Indice]
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						Send, 1
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						SendInput, % Demanda[Indice]
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						
						Send, 40
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						
						SendInput, % DJCA[Indice]
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						
						Send, 40
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						
						SendInput, 16
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						
							MouseMove, 150, 310	
							Sleep, 175
							Click, 1
							Sleep, 175
							
						if(Remoto[Indice] = "SIM" )
						{
							aux := 2
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						}
						else if(Remoto[Indice] = "NAO" )
						{
							aux := 1
							while(aux>0) {
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						}
						else{
							MsgBox [,,Erro - Valor Despadronizado, A planilha não está padronizada conforme o estipulado: SIM / NAO .  Verifique a planilha e corrija o problema.  Caso queira continuar o procedimento basta selecionar o valor correto e clicar em "ESC"]
							pause
						}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
							if(Atendimento_Trafo[Indice] = "0" ){
							goto, next1
						}
						
						MouseMove, 450, 310	
							Sleep, 175
							Click, 1
							Sleep, 175
							aux := % Index_Trafo_Atendimento[Indice]
						while(aux>0){
								SendInput,{Down}
								Sleep, 50
								aux := aux - 1
						}
						Sleep, 175
						Send, {Enter}
						
						next1:
						

						Sleep, 175
						Send, {Tab}
						Sleep, 175
						
						;~ Preenche o numero de hastes 
						Send, 3
						Sleep, 175
						Send, {Enter}
						Sleep, 175
						
						MouseMove, 500, 344		
						Sleep, 175
						Click, 1
						Sleep, 175
							aux := % Index_Fuso[Indice] 
							while(aux>0) {
								SendInput,{Down}
								sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						SendInput, % Longitude[Indice]
						Sleep, 175
						Send, {Tab}
						Sleep, 175
					
						SendInput, % Latitude[Indice]
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						MouseMove, 120, 431				
						Sleep, 175
						Click, 1
						Sleep, 175
						
						if(Ligacao[Indice] = "BAIXA" )
						{
							aux := 1
							while(aux>0) {
								SendInput,{Down}
								sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						}
						else if(Ligacao[Indice] = "ALTA" )
						{
							aux := 2
							while(aux>0) {
								SendInput,{Down}
								sleep, 50
								aux := aux - 1
							}
							Sleep, 175
							Send, {Enter}
						}
						Sleep, 175
						Send, {Tab}
						
						SendInput, % Ncabos[Indice]
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						SendInput, % PG[Indice]
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						SendInput, % Fase[Indice]
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						SendInput, % Neutro[Indice] 
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						SendInput, % Terra[Indice]
						Sleep, 175
						Send, {Tab}
						
						
						
						;~ --------------------------------------------------------------------------------------
						;~ ----------------------------------      Em Obras     ----------------------------------
						;~ --------------------------------------------------------------------------------------
						
						Sleep, 175
						MouseMove, 645, 430	
							Sleep, 175
							Click, 1
							Sleep, 175
							FormatTime, Month, %A_Now%, M
							
							while(Month>0){
								SendInput,{Down}
								Sleep, 50
								Month := Month - 1
							}
							Sleep, 175
							Send, {Enter}
							Sleep, 175
							Send, {Tab}
							Sleep, 175
							Send, 2023
							Sleep, 175
							Send, {Tab}
							Sleep, 175
							
							
							MouseMove, 930, 430	
							Sleep, 175
							Click, 1
							Sleep, 175
						if(Zona[Indice] = "RURAL" )
						{
							aux := 2
							while(aux>0) {
								SendInput,{Down}
								sleep, 50
								aux := aux - 1
							}
							Sleep, 175
						}
						else if(Zona[Indice] = "URBANO" )
						{
							aux := 1
							while(aux>0) {
								SendInput,{Down}
								sleep, 50
								aux := aux - 1
							}
							Sleep, 175
						}
						else{
							MsgBox [,,Erro - Valor Despadronizado, A planilha não está padronizada conforme o estipulado: URBANO / RURAL .  Verifique a planilha e corrija o problema.  Caso queira continuar o procedimento basta selecionar o valor correto e clicar em "ESC"]
							pause
						}
						Sleep, 175
						Send, {Enter}
						Sleep, 100
						Send, {Tab}
						Sleep, 100
						Send, {Tab}
						Sleep, 175
						SendInput, % Nplaca[Indice]
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						
						
						; escolhe o modelo da placa
						if( Linha_Marca_Placa[Indice] = "X"){
						MsgBox % "A marca da Placa ainda não foi cadastrada      " . Marca_Inv[Indice] 
							pause
						}
						else if( Linha_Marca_Placa[Indice] = "#N/A"){
						MsgBox % "A marca da Placa ainda não foi cadastrada       " . Marca_Inv[Indice] 
							pause
						}
						else{
								MouseMove, 416, 700
						Sleep, 175
						Click, 1
						Sleep, 175
						aux := Linha_Marca_Placa[Indice] + 1
						while(aux>0){
						SendInput,{Down}	
						sleep, 50
						aux := aux - 1
						}
						Sleep, 175
						Send, {Enter}
						}
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						
						; escolhe a placa
						if( Linha_Placa[Indice] = "X"){
							MsgBox % "A marca da Placa ainda não foi cadastrada       " . Marca_Placa[Indice] 
							pause
						}
						else if( Linha_Placa[Indice] = "#N/A"){
							MsgBox % "A marca da Placa ainda não foi cadastrada       " . Marca_Placa[Indice] 
							pause
						}
						else{
						Sleep, 175
						MouseMove, 700, 700
						sleep, 400
						Click, 1
						sleep, 400
						aux := Linha_Placa[Indice]+ 1
						while(aux>0){
						SendInput,{Down}
						sleep, 50
						aux := aux - 1
						}
						Sleep, 175
						Send, {Enter}
						}
						

						Sleep, 175
						Send, {Tab}
						Sleep, 175
						SendInput, % Area[Indice]
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						SendInput, % PotMod[Indice]
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						
						aux := 9
						while(aux>0){
						SendInput,{Down}
						sleep, 50
						aux := aux - 1
						}
						
						Sleep, 175
						SendInput, % Ninv[Indice]
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						if( Linha_Marca_Inv[Indice] = "X"){
						MsgBox % "A Linha da marca do Inversor ainda não foi cadastrada      " . Marca_Inv[Indice] 
							pause
						}
						if( Linha_Marca_Inv[Indice] = "#N/A"){
						MsgBox % "A marca do Inversor ainda não foi cadastrada      " . Marca_Inv[Indice] 
							pause
						}
						else{
							MouseMove, 400, 1020
						Sleep, 175
						Click, 1
						Sleep, 175
						aux := Linha_Marca_Inv[Indice]+ 1
						while(aux>0){
						SendInput,{Down}
						sleep, 50
						aux := aux - 1
						}
						Sleep, 175
						Send, {Enter}
						}
						Sleep, 175
						Send, {Tab}
						; escolhe o Inversor
						if( Linha_Inv[Indice] = "X"){
						MsgBox % "O modelo do Inversor ainda não foi cadastrado      " . Inv[Indice] 
							pause
						}
						else if( Linha_Inv[Indice] = "#N/A"){
						MsgBox % "O modelo do Inversor ainda não foi cadastrado       " . Inv[Indice] 
							pause
						}
						
						else{
						Sleep, 175
						MouseMove, 730, 1020
						Sleep, 175
						Click, 1
						if(Linha_Inv[Indice] = 0){
						goto, next2
						}
						aux := Linha_Inv[Indice]+ 1
						while(aux>0){
						SendInput,{Down}
						sleep, 50
						aux := aux - 1
						}
						Sleep, 175
						Send, {Enter}
						}

						next2:
						
						Sleep, 175
						Send, {Tab}
						Sleep, 175
						SendInput, % PInv[Indice]
						Sleep, 175
						Send, {Enter}
						Sleep, 175
						
						Pause
						
						if(  Acoplado_Trafo[Indice] = "0" ){
							Sleep, 175
						goto, avancar	
						}
						
						autotrafo:
						
						aux := 9
						while(aux>0){
						SendInput,{Down}
						sleep, 75
						aux := aux - 1
						}
						Sleep, 175
						Send, SIM
						Sleep, 175
						Send, {Enter}
						Sleep, 175
						SendInput, % Acoplado_Trafo[Indice]

						Pause
						
						Sleep, 175
						Send, {Enter}
						Sleep, 175
						
						Acoplado_Trafo[Indice]
						
						Send, {Enter}
						Sleep, 175
						Send, {Enter}
						Sleep, 175
						goto, avancar

						
						Pause
						
						avancar:
						
						Sleep, 175
						aux = 3
						while(aux>0){
						Send {WheelDown}
						sleep, 50
						aux := aux - 1
						}
						Sleep, 175
						
						ImageSearch, X , Y , 1000, 300, 1920, 1080, %A_WorkingDir%\lib\Avancar.PNG
						
						if(ErrorLevel = 1){
							
							Sleep, 175
							goto, avancar
						
						}
						
						else if( ErrorLevel = 0){
							
							MouseMove, % X + 80, % Y + 30
							sleep, 150
							Click, 1
							Sleep, 150 
						
						}
						
						
						WinWait, Message Box Title
						Sleep, 700
						WinActivate, Message Box Title
						WinWaitActive, Message Box Title
						Sleep, 700
							if(Remoto[Indice] = "SIM" )
						{
							Send, S
							goto, cadastro
							
						}
						
						else if(Remoto[Indice] = "NAO" )
						{
							Send, N
							goto, save
						}
						
						cadastro:
						
						Sleep, 1750
						
						Indice_Comp:= 0
						While(Indice_Comp<=IndiceTotal_Comp){
						if (Nome_Comp[Indice_Comp]=Nome[Indice]){
							SendInput, % UC_Comp[Indice_Comp]
							sleep, 150
							Send, {Tab}
							Sleep, 175
							SendInput, % Nome_Comp[Indice_Comp]
							sleep, 150
							Send, {Tab}
							Sleep, 175
							SendInput, % CPF_Comp[Indice_Comp]
							sleep, 150
							Send, {Tab}
							Sleep, 175
							SendInput, % End_Comp[Indice_Comp]
							sleep, 150
							Send, {Tab}
							Sleep, 175
							SendInput, % Perc_Comp[Indice_Comp]
							sleep, 150
							Send, {Tab}
							Sleep, 175
							
						}
						
							Indice_Comp:= Indice_Comp+1
						}

					
					cadastro_next:
					
							Sleep, 175
							MouseMove, 1160, 690	
							sleep, 150
							Click, 1
							Sleep, 1500
							goto, save
								
						
						
					save:


							Sleep, 1750
							MouseMove, 975, 970
							sleep, 150
							Click, 1
							Sleep, 1500 
						
						
						 save2:
						WinWait, Salvar como
						Sleep, 700
						WinActivate, Salvar como
						WinWaitActive, Salvar como
						Sleep, 700
						MouseMove, 550, 60
							sleep, 150
							Click, 1
							Sleep, 150 
						Send, C:\Users\Nilson Gasparin\Dropbox\MONTE CRISTO\AUTOMAÇÃO\ROTATIVO
						sleep, 150
						send, {Enter}
						Sleep, 1750
						send, {Enter}
						Sleep, 175
							MouseMove, 800, 500
							sleep, 150
							Click, 1
							Sleep, 3000 
						Send,{Alt}{Tab}
						Sleep, 1750
						WinWait,Microsoft Excel
						Sleep, 700
						WinActivate, Microsoft Excel
						WinWaitActive, Microsoft Excel
						Sleep, 700
						send, {Enter}
						Sleep, 1750
						WinActivate, Microsoft Excel
						WinWaitActive, Microsoft Excel
						Sleep, 1750
						send, {Enter}
						Sleep, 300
						send, {Enter}
						Sleep, 5000
						
						goto,info
						
					if(aux1 < IndiceTotal){
						aux1:= aux1  + 1
					Indice := aux2[aux1]
					goto, loop1
					}
					
					goto, end
					
					
					
					
				
					

				
					
	
					;~ Is_It_OK:
						
						;~ IfWinExist ahk_class #32770 
						;~ {
							;~ WinActivate ahk_class #32770 
						;~ Sleep, 175
						
						;~ pause
						
						;~ }
						;~ ; erro no campo uc
						;~ IfWinExist APENAS O NÚMERO DA UC
						;~ {
						;~ Sleep, 175
						
						;~ pause
						
						;~ }
						;~ return
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
			
				Else If(Ct = 2){
					
					MsgBox, 0 , Versão do Programador - Estagiário do Estagiário, Contratando um Estagiario para me substituir, 5
					
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
				
				
				
				
				