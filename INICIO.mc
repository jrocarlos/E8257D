my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            E8257D v2
DATE:                  2020-01-14 11:47:44
AUTHOR:                Daniel Sarmento
REVISION:              2
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       1
NUMBER OF LINES:       50
=============================================================================
 STEP    FSC    RANGE NOMINAL        TOLERANCE     MOD1        MOD2  3  4 CON


#------------------CONFIG PLANILHA---------------
  1.001  DISP         CERTIFIQUE-SE QUE A PLANILHA DE DADOS
  1.001  DISP         FOI DEVIDADEMENTE CRIADA COM OS PONTOS E PADRÕES
  1.001  DISP         UTILIZADOS NESTA CALIBRAÇÃO E ESTÁ SALVA COM
  1.001  DISP         O NOME "E8257D" NO ENDEREÇO:
  1.001  DISP         "Z:/Software/PLANILHAS"
#---------------------CONFIG EXCEL-------------------
  1.002  MATH         xlFile = "Z:/Software/PLANILHAS/E8257D.xlt"
  1.003  LIB          COM xlApp = "Excel.Application";
  1.004  LIB          xlApp.Visible = True;
  1.005  LIB          COM xlWB = xlApp.Workbooks;
  1.006  LIB          xlWB.Open(xlFile);
#------------------CONFIG WORKSHEET-------------------
  1.007  LIB          COM xlWS = xlApp.Worksheets["FREQ"];
  1.008  LIB          xlWS.Select();
#-----------------------CONFIG TEST FREQUENCY------------------------
  1.009  OPBR         DESEJA CALIBRAR FREQUÊNCIA?
  1.010  JMPT         1.012
  1.011  JMP          1.013
  1.012  CALL         FREQUENCIA
#-----------------------CONFIG TESTE LINEARITY------------------------
 # 1.017  OPBR         DESEJA CALIBRAR LINEARIDADE?
  #1.018  JMPT         1.020
  #1.019  JMP          1.021
 # 1.020  CALL         E8257D-4
#-----------------------CONFIG TESTE RESPONSE------------------------
  1.013  OPBR         DESEJA CALIBRAR RESPOSTA EM FREQUÊNCIA?
  1.014  JMPT         1.016
  1.015  JMP          1.017
  1.016  CALL         RESPOSTAGEN
#-----------------------CONFIG TEST REFERENCE------------------------
#  1.025  OPBR         DESEJA CALIBRAR FREQUÊNCIA DE REFERÊNCIA?
 # 1.026  JMPT         1.028
 # 1.027  JMP          1.029
 ## 1.028  CALL         E8257D-6
#-----------------------END CAL------------------------
  1.017  DISP         FIM
