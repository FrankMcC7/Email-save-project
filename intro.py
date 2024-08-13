=TRIM(MID(A1 & " ", FIND(",", A1) + 1, LEN(A1)) & " " & LEFT(A1, FIND(",", A1) - 1))
