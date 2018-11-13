'-- Expressões para utilizar quando necessário buscar datas/períodos em tabelas com campo de data e hora.

'1 - Expressão SELECT: Nessa expressão o formato da data será universal (mm/dd/yyyy)
"SELECT [DATA/HORA] FROM TBL_Log_Abertura WHERE (Format([Data/Hora],'dd/mm/yyyy')) Between #" & txt_data_inicial & "# and #" & txt_data_final & "#"

'2 - Expressão SELECT 2: Nessa expressão o formato da data será convertida (dd/mm/yyyy)
"SELECT [Data/Hora] FROM TBL_Log_Abertura WHERE CDate(Format([Data/Hora],'dd/mm/yyyy')) Between '" & txt_data_inicial & "' and '" & txt_data_final & "'"




