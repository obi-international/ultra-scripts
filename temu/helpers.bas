' =RIGHT(B2, LEN(B2) - 3) 
' =ROUND(((" & G2 & " * " & exchangeRate & ") * 0.02) + (((" & G2 & " * " & exchangeRate & ") + ((" & G2 & " * " & exchangeRate & ") * 0.02)) * 0.2), 0)