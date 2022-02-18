my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            FREQUENCIA
DATE:                  2020-01-10 16:34:43
AUTHOR:                Daniel Sarmento
REVISION:              01
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       1
NUMBER OF LINES:       211
=============================================================================
 STEP    FSC    RANGE NOMINAL        TOLERANCE     MOD1        MOD2  3  4 CON


#------------CONFIG EXCEL---------------
  1.001  LIB          COM xlWS = xlApp.Worksheets["FREQ"];
  1.002  LIB          xlWS.Select();
#------------------CONFIG GENERATOR---------
  1.003  RSLT         =
  1.004  IEEE         [@20]*RST
  1.005  IEEE         *CLS
  1.006  IEEE         :OUTP:MOD OFF
  1.007  IEEE         POW 0
#-------------------CONFIG  Nº MEAS----------------
  1.008  MEMI         DIGITE O NÚMERO DE MEDIDAS
  1.009  MATH         A = RND(MEM)
  1.010  IF           A == 0
  1.011  JMP          1.153
  1.012  ENDIF
#-----------------CONFIG POINT------------------
  1.013  MATH         P = 0
  1.014  MATH         LP = 2
  1.015  MATH         CP = 1
  1.016  MATH         T  = 0
  1.017  MATH         LINHA = 2
  1.018  MATH         COLUNA = 3
  1.019  MATH         EDUARDUINO = 0
  1.020  MATH         M = 0
  1.021  DO
  1.022  LIB          COM P1 = xlApp.Cells[LP,CP];
  1.023  LIB          PONTO = P1.Value2;
  1.024  IF           PONTO == 0
  1.025  JMP          1.153
  1.026  ENDIF
  1.027  MATH         CP = CP + 1
  1.028  LIB          COM T1 = xlApp.Cells[LP,CP];
  1.029  LIB          TEX = T1.Value2;
  1.030  MATH         P = PONTO&TEX
  1.031  MATH         EX = 0
  1.032  MATH         Z1 = CMP  (TEX,"MHz")
  1.033  MATH         Z2 = CMP  (TEX,"kHz")
  1.034  MATH         Z3 = CMP  (TEX,"GHz")
  1.035  MATH         Z4 = CMP (TEX,"Hz")
#esquema pra ver se o ponto é o primeiro a ser executado em toda cal
# ESSA VAR M CONTROLA O DISPLAY, PRA NAO SER EXIBIDO TODA VEZ QUE
# O PROGRAMA ESCOLHER O PONTO A SER CAL

  1.036  MATH         LP2 = LP-1
  1.037  MATH         CP2 = CP +1
  1.038  LIB          COM VALOR = xlApp.Cells[LP2,CP2];
  1.039  LIB          PRIMEIRO = VALOR.Value2;

  1.040  IF           M==0
  1.041  MATH         M = CMP (PRIMEIRO,"1ª MEDIDA")
  1.042  ENDIF


#----------------------GATE-------------------------------
  1.043  IF           PONTO < 1 && Z4 == 1
  1.044  MATH         GATE = 100
  1.045  ELSE
  1.046  MATH         GATE = 10
  1.047  ENDIF

#CONTROLA VAR EM FUNCAO DO PONTO SER > OU < QUE 100 MHZ e 3 GHz

  1.048  IF           (PONTO <=100 && Z1 == 1) || (Z2 == 1)
  1.049  MATH         MENOR = 1
  1.050  MATH         MAIOR = 0
  1.051  MATH         MEIO = 0
  1.052  IF           M == 1
  1.053  DISP         ATENÇÃO
  1.053  DISP         CONECTE O GERADOR AO CH1 DO 53132A
  1.054  MATH         M = 2
  1.055  ENDIF
  1.056  JMP          1.076
  1.057  ELSEIF       (PONTO > 100 && Z1 ==1) || (PONTO <= 3 && Z3==1)
  1.058  MATH         MAIOR = 0
  1.059  MATH         MENOR = 0
  1.060  MATH         MEIO = 1
  1.061  IF           M == 2 || M ==1
  1.062  DISP         ATENÇÃO
  1.062  DISP         CONECTE O GERADOR AO CH3 DO 53132A
  1.063  MATH         M = 3
  1.064  ENDIF
  1.065  JMP          1.076

  1.066  ELSEIF       (PONTO > 3 && Z3 == 1)
  1.067  MATH         MENOR = 0
  1.068  MATH         MAIOR  = 1
  1.069  MATH         MEIO = 0
  1.070  IF           M == 3 || M ==1
  1.071  DISP         ATENÇÃO
  1.071  DISP         CONECTE O GERADOR AO CH2 DO 53152A
  1.072  MATH         M = 0
  1.073  ENDIF
  1.074  JMP          1.091
  1.075  ENDIF




#---------------    CONFIG COUNT - FREQ <= 3 GHz  - 53132A  ----------------
  1.076  IEEE         [@5]*RST
  1.077  IF           MENOR == 1
  1.078  IEEE         :FUNC 'FREQ 1'
  1.079  IF           (Z4 ==1) || (Z2 ==1 && PONTO <=100)
  1.080  IEEE         INP1:FILT ON
  1.081  ENDIF
  1.082  ELSEIF       MEIO == 1
  1.083  IEEE         :FUNC 'FREQ 3'
  1.084  ENDIF
  1.085  IEEE         INIT:CONT OFF
  1.086  IEEE         INP1:COUP DC
  1.087  IEEE         INP1:IMP 50
  1.088  IEEE         EVEN1:LEV:AUTO ON
  1.089  IEEE         INP1:FILT OFF
  1.090  JMP          1.096

#----------------- CONFIG COUNT - FREQ > 3 GHz - 53152A-------------
  1.091  IEEE         [@19]*RST
  1.092  IEEE         :FUNC 'FREQ 2'
  1.093  IEEE         :ROSC:SOUR EXT
  1.094  IEEE         :TRIG:HOLD 1.0
  1.095  IEEE         :FREQ:RES 1 Hz

#----------------------END-------------------------------
  1.096  IF           P == 00
  1.097  JMP          1.153
  1.098  ENDIF


#----------------------------CONFIG OUT GENERATOR--------------
  1.099  IEEE         [@20]:FREQ [V P]
  1.100  IEEE         :POW:STAT ON
  1.101  WAIT         [D2000]


#------------CONFIG IN COUNT ----------------

  1.102  MATH         TEMPO = GATE + (GATE / 2)
  1.103  IF           MENOR == 1 || MEIO == 1
  1.104  IEEE         [@5]FREQ:ARM:STOP:TIM [V GATE]
  1.105  IEEE         INIT:CONT ON
  1.106  ELSEIF       MAIOR == 1
  1.107  IEEE         [@19]FREQ:ARM:STOP:TIM [V GATE]
  1.108  IEEE         INIT:CONT ON
  1.109  ENDIF

  1.110  DO
  1.111  WAIT         -t [V TEMPO] Please Standby
  1.112  IEEE         DATA?[I]
  1.113  IF           Z1 == 1
  1.114  MATH         EX = 1E6
  1.115  ENDIF
  1.116  IF           Z2 == 1
  1.117  MATH         EX = 1E3
  1.118  ENDIF
  1.119  IF           Z3 == 1
  1.120  MATH         EX = 1E9
  1.121  ENDIF
  1.122  IF           Z1 == 1 && PONTO == 1
  1.123  MATH         EX = 1E6
  1.124  ENDIF
  1.125  IF           Z2 == 1 && PONTO == 1
  1.126  MATH         EX = 1E3
  1.127  ENDIF
  1.128  IF           Z3 == 1 && PONTO == 1
  1.129  MATH         EX = 1E9
  1.130  ENDIF
  1.131  MATH         MEM = MEM / EX
  #1.126  MEMCX        0              TOL
  1.132  LIB          COM selectedCell = xlApp.Cells[LINHA,COLUNA];
  1.133  LIB          selectedCell.Select();
  1.134  LIB          selectedCell.FormulaR1C1 = [MEM];
  1.135  MATH         T = T + 1
  1.136  MATH         COLUNA = COLUNA + 1
  1.137  MATH         CP = CP + 1
  1.138  UNTIL        T == A
  1.139  MATH         T  = 0
  1.140  MATH         COLUNA = 3
  1.141  MATH         LINHA = LINHA + 1
  1.142  MATH         CP = 1
  1.143  MATH         LP = LP + 1


  1.144  MATH         Z4 = CMP  (TEX,"MHz")
  1.145  MATH         Z5 = CMP  (TEX,"kHz")
  1.146  MATH         Z6 = CMP  (TEX,"GHz")

  1.147  LIB          COM T2 = xlApp.Cells[LP,CP];
  1.148  LIB          TEX2 = T2.Value2;
  1.149  LIB          COM P2 = xlApp.Cells[LP,CP];
  1.150  LIB          PONTO2 = P2.Value2;

  1.151  UNTIL        PONTO == 0

  1.152  JMP          1.153
#------------------RESET------------------
  1.153  IEEE         [@5]*RST
  1.154  IEEE         [@19]*RST
  1.155  IEEE         [@20]*RST
