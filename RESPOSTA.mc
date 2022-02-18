my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            RESPOSTAGEN
DATE:                  2020-01-15 17:25:23
AUTHOR:                Daniel Sarmento
REVISION:
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       1
NUMBER OF LINES:       247
=============================================================================
 STEP    FSC    RANGE NOMINAL        TOLERANCE     MOD1        MOD2  3  4 CON

#--------------------------------VARIABLES---------------------------------
  1.001  MATH         P = 0
  1.002  MATH         LP = 2
  1.003  MATH         CP = 1
  1.004  MATH         T  = 0
  1.005  MATH         L = 0
  1.006  MATH         CALF = 0
  1.007  MATH         LINHA = 2
  1.008  MATH         COLUNA = 5
  1.009  MATH         I = 0
  1.010  MATH         FACTOR = 0
  1.011  MATH         PCT = "PCT"
  1.012  MATH         TEMPO = 5
  1.013  MATH         L2 = 2
  1.014  MATH         L10 = 10
  1.015  MATH         EDUARDO = 0
#----------------------CLEAR-----------------------------
  1.016  ASK-   R D   N B            P J S U       M C X Z        A  L  T  W
  #---------------------- SHEET ----------------------
  1.017  DISP         [32]          ATENÇÃO!!!!
  1.017  DISP
  1.017  DISP         ANTES DE INICIAR INSIRA CORRETAMENTE
  1.017  DISP         OS VALORES DE CAL FACTOR ATUALIZADOS
  1.017  DISP         NA PLANILHA DE DADOS UTILIZADA
  1.017  DISP
  1.017  DISP         UTILIZE PONTO (.) NO LUGAR DE VIRGULA(,)
  1.017  DISP         [32]
  1.017  DISP         [32]     GPIB POWER METER E4419B = 14
  1.017  DISP         [32]     GPIB GERADOR  = 20
#---------------------------CONFIG EXCEL--------------------------
  1.018  LIB          COM xlWS = xlApp.Worksheets["RES"];
  1.019  LIB          xlWS.Select();
#-----------------------LOSS INSERTION----------------------
 # 1.020  MEM2         INSIRA O VALOR DE PERDA DO CONECTOR
 # 1.021  DO
 # 1.022  LIB          COM selectedCell2 = xlApp.Cells[L2,L10];
 # 1.023  LIB          selectedCell2.Select();
 # 1.024  LIB          selectedCell2.Value2 = [MEM2];
 # 1.025  MATH         L2 = L2 + 1
 # 1.026  UNTIL        L2 == 19
#-----------------------------CONFIG  Nº MEAS----------------
  1.020  MEMI         DIGITE O NÚMERO DE MEDIDAS
  1.021  MATH         A = RND(MEM)
  1.022  IF           A == 0
  1.023  JMP          1.151
  1.024  ENDIF
#----------------------------CONFIG GENERATOR----------------------
  1.025  RSLT         =
  1.026  IEEE         [@20]*RST
  1.027  IEEE         *CLS
  1.028  IEEE         :OUTP:MOD OFF
 #---------------------------CONFIG METER-----------------------
  1.029  IEEE         [@14]*RST
  1.030  IEEE         *CLS
  1.031  IEEE         SYST:PRES
#----------------ID SENSOR------------------
  1.032  IEEE         SERV:SENS:TYPE?[I$]
  1.033  MATH         SENSOR = MEM2
  1.034  OPBR         VERIQUE AS FAIXAS DE FREQUENCIA E POTENCIA
  1.034  OPBR         DO SENSOR ANTES DE CONTINUAR
  1.034  OPBR
  1.034  OPBR         SENSOR CONECTADO: [MEM2]
  1.034  OPBR
  1.034  OPBR         DESEJA CONTINUAR?
  1.035  JMPT         1.037
  1.036  JMP          1.151
#----------------------------SETUP REF--------------
  1.037  OPBR         DESEJA FAZER O ZERO/CAL DO SENSOR?
  1.038  JMPT         1.040
  1.039  JMP          1.082
  1.040  DISP         CONECTE O SENSOR AO POWER METER EM 'POWER REF'
  1.040  DISP
  1.040  DISP         [32]   POWER METER         to         SENSOR
  1.040  DISP         [32]
  1.040  DISP         [32]
  1.040  DISP         [32]   POWER REF -------------------> SENSOR
  1.040  DISP         [32]
  #1.045  PIC          SETUP2-1


#----------------ZERO/CAL FACTOR------------------
  1.041  MATH         F1 = CMP (SENSOR,"A")
  1.042  MATH         F1 = CMP (SENSOR,"B")
  1.043  MATH         F1 = CMP (SENSOR,"D")
  1.044  MATH         F1 = CMP (SENSOR,"H")
  1.045  IF           F1 == 1
  1.046  MEM2         ENTRE COM O CAL FACTOR DE REFERENCIA (50 MHz)
  1.047  MATH         REFFACTOR = MEM2&PCT
  1.048  IEEE         SENS1:CORR:CFAC [V REFFACTOR]
  1.049  IEEE         CAL1:ZERO:AUTO ONCE
  1.050  WAIT         -t 12 ZEROING CHA
  1.051  IEEE         CAL1:RCF [V REFFACTOR]
  1.052  IEEE         CAL1:AUTO ONCE
  1.053  WAIT         -t 13 CALIBRATING CHA
  1.054  IEEE         OUTP:ROSC ON
  1.055  WAIT         -t 8 VERIFIQUE O ZERO
  1.056  OPBR         DESEJA CONTINUAR?
  1.057  JMPT         1.059
  1.058  JMP          1.151
  1.059  IEEE         OUTP:ROSC OFF
  1.060  IEEE         INIT:CONT ON
  1.061  ELSE
  1.062  IEEE         [@14]*CLS
  1.063  IEEE         SYST:PRES
  1.064  WAIT         -t 5 PRESETING
  1.065  IEEE         SENS1:FREQ:CW 50MHz
  1.066  IEEE         CAL1:ZERO:AUTO ONCE
  1.067  WAIT         -t 22 ZEROING CHA
  1.068  IEEE         CAL1:AUTO ONCE
  1.069  WAIT         -t 10 CALIBRATING CHA
  1.070  IEEE         AVER:COUN:AUTO ON
  1.071  IEEE         OUTP:ROSC ON
  1.072  WAIT         -t 5 VERIFIQUE O ZERO
  1.073  OPBR         DESEJA CONTINUAR?
  1.074  JMPT         1.076
  1.075  JMP          1.151
  1.076  IEEE         OUTP:ROSC OFF
  1.077  IEEE         INIT:CONT ON
  1.078  MATH         T  = 0
  1.079  MATH         COLUNA = 5
  1.080  MATH         CP = 1
  1.081  MATH         I = 0
  1.082  ENDIF
#----------------------------SETUP GENERATOR--------------
  1.083  DISP         Connect the generator to the UUT as follows:
  1.083  DISP
  1.083  DISP         [32]   Generator         to         Meter
  1.083  DISP         [32]
  1.083  DISP         [32]   OUTPUT -------------------> POWER SENSOR
  1.083  DISP         [32]
#  1.137  PIC          SETUP2-2


#-----------------CONFIG POINT------------------

  1.084  DO
  1.085  LIB          COM P1 = xlApp.Cells[LP,CP];
  1.086  LIB          PONTO = P1.Value2;
  1.087  IF           PONTO == 0
  1.088  JMP          1.151
  1.089  ENDIF
  1.090  MATH         CP = CP + 1
  1.091  LIB          COM T1 = xlApp.Cells[LP,CP];
  1.092  LIB          TEX = T1.Value2;
  1.093  MATH         P = PONTO&TEX
  1.094  MATH         Z1 = CMP  (TEX,"MHz")
  1.095  MATH         Z2 = CMP  (TEX,"GHz")
#----------------LEVEL-------------------
  1.096  MATH         CP = CP + 1
  1.097  LIB          COM L1 = xlApp.Cells[LP,CP];
  1.098  LIB          L = L1.Value2;
  #1.099  IF           P >= 4 && Z2 == 1 && EDUARDO == 0
 # 1.100  IEEE         [@20]:POW:STAT OFF
  #1.101  DISP         SUBSTITUA O SENSOR!!!
  #1.101  DISP         PARA UMA FAIXA ACIMA DE 4 GHz
  #1.102  MATH         EDUARDO = EDUARDO + 1
 # 1.103  JMP          1.029
 # 1.104  ENDIF
#----------------------END-------------------------------
  1.099  IF           P == 00
  1.100  JMP          1.151
  1.101  ENDIF
#----------------SETUP-------------------
  1.102  IF           I == 0
  1.103  MATH         I = I + 1

  1.104  ENDIF
#----------------------------CONFIG OUT GENERATOR--------------
  1.105  IEEE         [@20]:FREQ [V P]
  1.106  IEEE         POW [V L]
  1.107  IEEE         :POW:STAT ON
  1.108  WAIT         [D2000]
#------------CONFIG IN METER----------------
  1.109  LIB          COM F1 = xlApp.Cells[LINHA,11];
  1.110  LIB          CALF = F1.Value2;
  1.111  MATH         S1 = CMP  (SENSOR,"A")
  1.112  MATH         S1 = CMP  (SENSOR,"B")
  1.113  MATH         S1 = CMP  (SENSOR,"D")
  1.114  MATH         S1 = CMP  (SENSOR,"H")
  1.115  IF           S1 == 1
  1.116  MATH         FACTOR = CALF&PCT
  1.117  IEEE         [@14]SENS1:CORR:CFAC [ V FACTOR]
  1.118  ENDIF
  1.119  IEEE         [@14]SENS1:FREQ:CW  [V P]
  1.120  IEEE         INIT:CONT ON
  1.121  DO
  1.122  WAIT         -t [V TEMPO] Please Standby
  1.123  IEEE         FETC?[I]
#------------------SAVE DATE----------------
  1.124  LIB          COM selectedCell = xlApp.Cells[LINHA,COLUNA];
  1.125  LIB          selectedCell.Select();
  1.126  LIB          selectedCell.FormulaR1C1 = [MEM];
  1.127  MATH         T = T + 1
  1.128  MATH         COLUNA = COLUNA + 1
  1.129  MATH         CP = CP + 1
  1.130  UNTIL        T == A
  1.131  MATH         T  = 0
  1.132  MATH         COLUNA = 5
  1.133  MATH         LINHA = LINHA + 1
  1.134  MATH         CP = 1
  1.135  MATH         LS = LP
  1.136  MATH         LP = LP + 1


##### COMPARA O SENSOR QUE SERA USADO AO USADO ANTERIOR PARA
### PERGUNTAR SE QUER EFETUAR NOVO ZERAMENTO


  1.137  LIB          COM S1 = xlApp.Cells[LS,12];
  1.138  LIB          SENS1 = S1.Value2;
  1.139  LIB          COM S2 = xlApp.Cells[LP,12];
  1.140  LIB          SENS2 = S2.Value2;

  1.141  IF           SENS1 != SENS2
  1.142  WAIT         [D2000]
  1.143  IEEE         [@20] :POW:STAT OFF

  1.144  MATH         SENSOR = MEM2
  1.145  OPBR         MODELO DE SENSOR ALTERADO NA PLANILHA!
  1.145  OPBR         SENSOR A SER UTILIZADO: TIPO [V SENS2 ]
  1.145  OPBR         OBSERVE A FAIXA DE OPERAÇAO
  1.145  OPBR         DESEJA REALIZAR O ZERO/CAL DESTE SENSOR?
  1.146  JMPT         1.040
  1.147  JMP          1.082


  1.148  ENDIF


  1.149  UNTIL        PONTO == 0
  1.150  JMP          1.151

#------------------RESET------------------
  1.151  IEEE         [@14]*RST
  1.152  IEEE         [@20]*RST
