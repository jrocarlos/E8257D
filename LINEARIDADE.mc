my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            LINEARIDADE GER
DATE:                  2020-01-14 11:47:07
AUTHOR:                Daniel Sarmento
REVISION:              01
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       1
NUMBER OF LINES:       170
=============================================================================
 STEP    FSC    RANGE NOMINAL        TOLERANCE     MOD1        MOD2  3  4 CON

# SUB ELABORADA COM BASE NA ORIGINAL "RESP", DESENV. POR CARLOS JUNIOR E
# DISPONIVEL NA VERSAO 01 DO PROG. DE ALGUNS GERADORES, COMO E8257D

#--------------------------------VARIÁVEIS---------------------------------
  1.001  MATH         P = 0
  1.002  MATH         LP = 2
  1.003  MATH         CP = 1
  1.004  MATH         T  = 0
  1.005  MATH         L = 0
  1.006  MATH         CALF = 0
  1.007  MATH         LINHA = 2
  1.008  MATH         TEMPO = 5
  1.009  MATH         COLUNA = 5
  1.010  MATH         I = 0
  1.011  MATH         FACTOR = 0
  1.012  MATH         PCT = "PCT"
  1.013  MATH         L2 = 2
  1.014  MATH         L10 = 10

#--------------------------------CLEAR---------------------------------
  1.015  ASK-   R D   N B            P J S U       M C X Z        A  L  T  W
  #---------------------- PLANILHA ----------------------
  1.016  DISP         [32]          ATENÇÃO!!!!
  1.016  DISP
  1.016  DISP         ANTES DE INICIAR INSIRA CORRETAMENTE
  1.016  DISP         OS VALORES DE CAL FACTOR ATUALIZADOS
  1.016  DISP         NA PLANILHA DE DADOS UTILIZADA
  1.016  DISP
  1.016  DISP         UTILIZE PONTO (.) NO LUGAR DE VIRGULA(,)
  1.016  DISP         [32]
  1.016  DISP         [32]     GPIB POWER METER E4419B = 14
  1.016  DISP         [32]     GPIB GERADOR E8257D = 19
#---------------------------CONFIG EXCEL--------------------------
  1.017  LIB          COM xlWS = xlApp.Worksheets["LIN"];
  1.018  LIB          xlWS.Select();
#-----------------------------CONFIG  Nº MEAS----------------
  1.019  MEMI         DIGITE O NÚMERO DE MEDIDAS
  1.020  MATH         A = RND(MEM)
  1.021  IF           A == 0
  1.022  JMP          1.106
  1.023  ENDIF
#-----------------------LOSS INSERTION----------------------
  1.024  MEM2         INSIRA O VALOR DE PERDA DO CONECTOR
  1.025  DO
  1.026  LIB          COM selectedCell2 = xlApp.Cells[L2,L10];
  1.027  LIB          selectedCell2.Select();
  1.028  LIB          selectedCell2.Value2 = [MEM2];
  1.029  MATH         L2 = L2 + 1
  1.030  UNTIL        L2 == 6
#----------------------------CONFIG GENERATOR----------------------
  1.031  RSLT         =
  1.032  IEEE         [@19]*RST
  1.033  IEEE         *CLS
  1.034  IEEE         :OUTP:MOD OFF
 #---------------------------CONFIG METER-----------------------
  1.035  IEEE         [@14]*RST
  1.036  IEEE         *CLS
  1.037  IEEE         SYST:PRES
#----------------ID SENSOR------------------
  1.038  IEEE         SERV:SENS:TYPE?[I$]
  1.039  MATH         SENSOR = MEM2
  1.040  OPBR         VERIQUE A FAIXA DE FREQUENCIA
  1.040  OPBR         DO SENSOR ANTES DE CONTINUAR
  1.040  OPBR
  1.040  OPBR         SENSOR CONECTADO TIPO: [MEM2]
  1.040  OPBR
  1.040  OPBR         DESEJA CONTINUAR?
  1.041  JMPT         1.043
  1.042  JMP          1.106
#-----------------ZERO METER----------------
  1.043  DISP         CONECTE O SENSOR NA PORTA "POWER REF"
  1.043  DISP
  1.043  DISP         [32]   POWER METER         to         SENSOR
  1.043  DISP         [32]
  1.043  DISP         [32]
  1.043  DISP         [32]   POWER REF -------------------> SENSOR
  1.043  DISP         [32]
#  1.044  PIC          SETUP2-1
#----------------CAL FACTOR------------------
  1.044  IEEE         SENS1:CORR:CFAC 99.80PCT
  1.045  IEEE         CAL1:ZERO:AUTO ONCE
  1.046  WAIT         -t 12 ZEROING CHA
  1.047  IEEE         CAL1:RCF 99.8PCT
  1.048  IEEE         CAL1:AUTO ONCE
  1.049  WAIT         -t 8 CALIBRATING CHA
  1.050  IEEE         OUTP:ROSC ON
  1.051  WAIT         -t 3 VERIFIQUE O ZERO
  1.052  OPBR         DESEJA CONTINUAR?
  1.053  JMPT         1.055
  1.054  JMP          1.106
  1.055  IEEE         OUTP:ROSC OFF
  1.056  IEEE         INIT:CONT ON
#-----------------CONFIG POINT------------------
  1.057  DO
  1.058  LIB          COM P1 = xlApp.Cells[LP,CP];
  1.059  LIB          PONTO = P1.Value2;
  1.060  IF           PONTO == 0
  1.061  JMP          1.106
  1.062  ENDIF
  1.063  MATH         CP = CP + 1
  1.064  LIB          COM T1 = xlApp.Cells[LP,CP];
  1.065  LIB          TEX = T1.Value2;
  1.066  MATH         P = PONTO&TEX
#----------------LEVEL-------------------
  1.067  MATH         CP = CP + 1
  1.068  LIB          COM L1 = xlApp.Cells[LP,CP];
  1.069  LIB          L = L1.Value2;
#----------------------END-------------------------------
  1.070  IF           P == 00
  1.071  JMP          1.106
  1.072  ENDIF
#----------------CAL FACTOR-------------------
  1.073  IF           I == 0
  1.074  MATH         I = I + 1
  1.075  JMP          1.104
  1.076  ENDIF
  1.077  LIB          COM F1 = xlApp.Cells[LINHA,11];
  1.078  LIB          CALF = F1.Value2;
  1.079  MATH         FACTOR = CALF&PCT
  1.080  IEEE         SENS1:CORR:CFAC [ V FACTOR]
#----------------------------CONFIG OUT GENERATOR--------------
  1.081  IEEE         [@19]:FREQ [V P]
  1.082  IEEE         POW [V L]
  1.083  IEEE         :POW:STAT ON
  1.084  WAIT         [D2000]
#------------CONFIG IN METER----------------
  1.085  IEEE         [@14]SENS1:FREQ:CW  [V P]
  1.086  IEEE         INIT:CONT ON
  1.087  DO
  1.088  WAIT         -t [V TEMPO] Please Standby
  1.089  IEEE         FETC?[I]
#------------------SAVE DATE----------------
  1.090  LIB          COM selectedCell = xlApp.Cells[LINHA,COLUNA];
  1.091  LIB          selectedCell.Select();
  1.092  LIB          selectedCell.FormulaR1C1 = [MEM];
  1.093  MATH         T = T + 1
  1.094  MATH         COLUNA = COLUNA + 1
  1.095  MATH         CP = CP + 1
  1.096  UNTIL        T == A
#---------------------SAVE LOSS---------------------
  1.097  MATH         T  = 0
  1.098  MATH         COLUNA = 5
  1.099  MATH         LINHA = LINHA + 1
  1.100  MATH         CP = 1
  1.101  MATH         LP = LP + 1
  1.102  UNTIL        PONTO == 0
  1.103  JMP          1.106
#----------------------------SETUP GENERATOR--------------
  1.104  DISP         Connect the generator to the UUT as follows:
  1.104  DISP
  1.104  DISP         [32]   Generator         to         Meter
  1.104  DISP         [32]   OUTPUT -------------------> POWER SENSOR
  1.104  DISP         [32]
 # 1.106  PIC          SETUP2-2
  1.105  JMP          1.077
#------------------RESET------------------
  1.106  IEEE         [@14]*RST
  1.107  IEEE         [@19]*RST
