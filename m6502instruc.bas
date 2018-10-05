Attribute VB_Name = "m6502Instruct"
' This is where all 6502 instructions are kept.
Public Function adc6502()
      Dim tmp As Integer
      
      adrmode (opcode)
      value = get6502memory(savepc)
      
      saveflags = (P And &H1)
      
      sum = A
      sum = (sum + value) And &HFF
      sum = (sum + saveflags) And &HFF
      
      If (sum > &H7F) Or (sum < -&H80) Then
        P = P Or &H40
      Else
        P = (P And &HBF)
      End If
      
      sum = A + (value + saveflags)
      If (sum > &HFF) Then
        P = P Or &H1
      Else
        P = (P And &HFE)
      End If
      
      A = sum And &HFF
      If (P And &H8) Then
             P = (P And &HFE)
              If ((A And &HF) > &H9) Then
                      A = (A + &H6)
              End If
              If ((A And &HF0) > &H90) Then
                      A = A + &H60
                      P = P Or &H1
             End If
      Else
        clockticks6502 = clockticks6502 + 1
      End If
      If (A) Then
      P = (P And &HFD)
      Else
       P = P Or H2
      End If
      If (A And &H80) Then
      P = P Or &H80
      Else
      P = (P And &H7F)
      End If
End Function

Public Function adrmode(opcode As Byte)
  Select Case addrmode(opcode)
    Case "absx6502": absx6502
    Case "immediate6502": immediate6502
    Case "relative6502": relative6502
    Case "absy6502": absy6502
    Case "abs6502": abs6502
    Case "zp6502": zp6502
    Case "zpx6502": zpx6502
    Case "zpy6502": zpy6502
    Case "implied6502": implied6502
    Case "indx6502": indx6502
    Case "indy6502": indy6502
    Case "indirect6502": indirect6502
    Case "indabsx6502": indabsx6502
    Case "indzp6502": indzp6502
  End Select
End Function

Public Function and6502()
  adrmode (opcode)
  value = get6502memory(savepc)
  A = (A And value)
  If (A) Then
    P = (P And &HFD)
  Else
    P = P Or &H2
  End If
  If (A And &H80) Then
    P = P Or &H80
  Else
    P = (P And &H7F)
  End If
End Function
Public Function asl6502()
  adrmode (opcode)
  value = get6502memory(savepc)
  
  P = (P And &HFE) Or ((value \ 128) And &H1)
  value = value * 2
  Call put6502memory((savepc), value And &HFF)
  If (value) Then
    P = (P And &HFD)
  Else
    P = P Or &H2
  End If
  If (value And &H80) Then
    P = P Or &H80
  Else
    P = (P And &H7F)
  End If
End Function


Public Function asla6502()
  If A > 0 Then
    A = A
  End If
  
  P = (P And &HFE) Or ((A \ 128) And &H1)
  A = (A * 2) And &HFF
  
  If (A) Then
    P = (P And &HFD)
  Else
    P = P Or &H2
  End If
  If (A And &H80) Then
    P = P Or &H80
  Else
    P = (P And &H7F)
  End If
End Function

Public Function bcc6502()
  If ((P And &H1) = 0) Then
    adrmode (opcode)
    PC = PC + savepc
    clockticks6502 = clockticks6502 + 1
  Else
    value = gameImage(PC)
    PC = PC + 1
  End If
End Function

Public Function bcs6502()
  If (P And &H1) Then
    adrmode (opcode)
    PC = PC + savepc
    clockticks6502 = clockticks6502 + 1
  Else
    value = gameImage(PC)
    PC = PC + 1
  End If
End Function

Public Function beq6502()
  If (P And &H2) Then
    adrmode (opcode)
    PC = PC + savepc
    clockticks6502 = clockticks6502 + 1
  Else
    value = gameImage(PC)
    PC = PC + 1
  End If
End Function

Public Function bit6502()
  adrmode (opcode)
  value = get6502memory(savepc)
  
  If (value And A) Then
    P = (P And &HFD)
    Else
    P = P Or &H2
  End If
  P = ((P And &H3F) Or (value And &HC0))
End Function

Public Function bmi6502()
  If (P And &H80) Then
    adrmode (opcode)
    PC = PC + savepc
    clockticks6502 = clockticks6502 + 1
  Else
    value = gameImage(savepc)
    PC = PC + 1
  End If
End Function

Public Function bne6502()
  If ((P And &H2) = 0) Then
    adrmode (opcode)
    PC = PC + savepc
    clockticks6502 = clockticks6502 + 1
  Else
    value = gameImage(PC)
    PC = PC + 1
  End If
End Function

Public Function bpl6502()
  If ((P And &H80) = 0) Then
    adrmode (opcode)
    PC = PC + savepc
    clockticks6502 = clockticks6502 + 1
  Else
    value = gameImage(PC)
    PC = PC + 1
  End If
End Function

Public Function brk6502()
  PC = PC + 1
  xx& = put6502memory(&H100 + S, (PC \ 256) And &HFF)
  
  If S <= 0 Then
  S = 0
  Else
  S = S - 1
  End If
  
  xx& = put6502memory(&H100 + S, (PC And &HFF))
  
  If S <= 0 Then
    S = 0
  Else
    S = S - 1
  End If
  
  xx& = put6502memory(&H100 + S, P)
  
  If S <= 0 Then
    S = 0
  Else
    S = S - 1
  End If
  
  P = P Or &H14
  PC = gameImage(65534) Or gameImage(65535) * 2 ^ 8
End Function

Public Function bvc6502()
  If ((P And &H40) = 0) Then
    adrmode (opcode)
    PC = PC + savepc
    clockticks6502 = clockticks6502 + 1
  Else
    value = gameImage(PC)
    PC = PC + 1
  End If
End Function

Public Function bvs6502()
  If (P And &H40) Then
    adrmode (opcode)
    PC = PC + savepc
    clockticks6502 = clockticks6502 + 1
  Else
    value = gameImage(PC)
    PC = PC + 1
  End If
End Function

Public Function clc6502()
  P = P And &HFE
End Function

Public Function cld6502()
  P = P And &HF7
End Function

Public Function cli6502()
  P = P And &HFB
End Function

Public Function clv6502()
  P = P And &HBF
End Function

Public Function cmp6502()
  adrmode (opcode)
  value = get6502memory(savepc)
  
  If (A + &H100 - value) > &HFF Then
    P = P Or &H1
  Else
    P = (P And &HFE)
  End If
  
  value = (A + &H100 - value) And &HFF
  If (value) Then
      P = (P And &HFD)
  Else
    P = (P Or &H2)
  End If
  If (value And &H80) Then
    P = (P Or &H80)
  Else
    P = (P And &H7F)
  End If
End Function

Public Function cpx6502()
  adrmode (opcode)
  value = get6502memory(savepc)
      
  If (x + &H100 - value > &HFF) Then
    P = P Or &H1
  Else
    P = (P And &HFE)
  End If
  
  value = (x + &H100 - value) And &HFF
  If (value) Then
    P = (P And &HFD)
  Else
    P = (P Or &H2)
  End If
  
  If (value And &H80) Then
    P = (P Or &H80)
  Else
    P = (P And &H7F)
  End If
End Function

Public Function cpy6502()
  adrmode (opcode)
  value = get6502memory(savepc)
      
  If (y + &H100 - value > &HFF) Then
    P = (P Or &H1)
  Else
    P = (P And &HFE)
  End If
    value = (y + &H100 - value) And &HFF
  If (value) Then
    P = (P And &HFD)
  Else
    P = (P Or &H2)
  End If
  If (value And &H80) Then
    P = (P Or &H80)
  Else
    P = (P And &H7F)
  End If
End Function

Public Function dec6502()
  adrmode (opcode)
      
  gameImage(savepc) = (gameImage(savepc) - 1) And &HFF
      
  value = get6502memory(savepc)
  If (value) Then
    P = (P And &HFD)
  Else
    P = (P Or &H2)
  End If
  If (value & &H80) Then
    P = (P Or &H80)
  Else
    P = (P And &H7F)
  End If
End Function
Public Function dex6502()
  x = (x - 1) And &HFF
  If (x) Then
    P = (P And &HFD)
  Else
    P = (P Or &H2)
  End If
  If (x And &H80) Then
    P = (P Or &H80)
  Else
    P = (P And &H7F)
  End If
End Function

Public Function dey6502()
  y = (y - 1) And &HFF
      
  If (y) Then
    P = (P And &HFD)
  Else
    P = (P Or &H2)
  End If
  If (y And &H80) Then
    P = P Or &H80
  Else
    P = (P And &H7F)
  End If
End Function

Public Function eor6502()
  adrmode (opcode)
  A = A Xor get6502memory(savepc)
  
  If (A) Then
    P = (P And &HFD)
  Else
    P = (P Or &H2)
  End If
  If (A And &H80) Then
    P = (P Or &H80)
  Else
    P = (P And &H7F)
  End If
End Function

Public Function inc6502()
  adrmode (opcode)
  gameImage(savepc) = (gameImage(savepc) + 1) And &HFF
  
  value = get6502memory(savepc)
  
  If (value) Then
    P = (P And &HFD)
  Else
    P = (P Or &H2)
  End If
  If (value And &H80) Then
    P = (P Or &H80)
  Else
    P = (P And &H7F)
  End If
End Function

Public Function inx6502()
  x = (x + 1) And &HFF
      
  If (x) Then
    P = (P And &HFD)
  Else
    P = (P Or &H2)
  End If
  If (x And &H80) Then
    P = (P Or &H80)
  Else
    P = (P And &H7F)
  End If
End Function

Public Function iny6502()
  y = (y + 1) And &HFF
  If (y) Then
    P = (P And &HFD)
  Else
    P = (P Or &H2)
  End If
  If (y And &H80) Then
    P = (P Or &H80)
  Else
    P = (P And &H7F)
  End If
End Function
Public Function jmp6502()
  adrmode (opcode)
  PC = savepc
End Function
Public Function jsr6502()
  PC = PC + 1
  put6502memory &H100 + S, PC \ 256
  S = (S - 1) And &HFF
  put6502memory &H100 + S, (PC And &HFF)
  S = (S - 1) And &HFF
  PC = PC - 1
  adrmode (opcode)
  PC = savepc
End Function
Public Function lda6502()
  adrmode (opcode)
  A = get6502memory(savepc)
  If (A) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (A And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function
Public Function ldx6502()
  adrmode (opcode)
  x = get6502memory(savepc)
      
  If (x) Then
    P = (P And &HFD)
  Else
    P = (P Or &H2)
  End If
  If (x And &H80) Then
    P = (P Or &H80)
  Else
    P = (P And &H7F)
  End If
End Function
Public Function ldy6502()
  adrmode (opcode)
  y = get6502memory(savepc)
  
  If (y) Then
    P = (P And &HFD)
  Else
    P = (P Or &H2)
  End If
  If (y And &H80) Then
    P = (P Or &H80)
  Else
    P = (P And &H7F)
  End If
End Function
Public Function lsr6502()
  adrmode (opcode)
  value = get6502memory(savepc)
         
  P = ((P And &HFE) Or (value And &H1))
  
  value = value \ 2
  put6502memory (savepc), value And &HFF
  
  If (value = Not 0) Then
    P = (P And &HFD)
  Else
    P = (P Or &H2)
  End If
  If ((value And &H80) = &H80) Then
    P = (P Or &H80)
  Else
    P = (P And &H7F)
  End If
End Function
Public Function lsra6502()
  P = (P And &HFE) Or (A And &H1)
  A = A \ 2
  If (A) Then
    P = (P And &HFD)
  Else
    P = P Or &H2
  End If
  If (A And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function
Public Function nop6502()
'TS: Implemented complex code structure ;)
End Function
Public Function ora6502()
  adrmode (opcode)
  A = A Or get6502memory(savepc)
      
  If (A) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (A And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function
Public Function pha6502()
  gameImage(&H100 + S) = A
  S = (S - 1) And &HFF
End Function
Public Function php6502()
  gameImage(&H100 + S) = P
  S = (S - 1) And &HFF
End Function
Public Function pla6502()
  S = (S + 1) And &HFF
  A = gameImage(S + &H100)
  If (A) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (A And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function
Public Function plp6502()
  S = (S + 1) And &HFF
  P = gameImage(S + &H100) Or &H20
End Function
Public Function rol6502()
  saveflags = (P And &H1)
  adrmode (opcode)
  value = get6502memory(savepc)
      
  If value > 0 Then
    value = value
  End If
  P = (P And &HFE) Or ((value \ 128) And &H1)
  If value >= 126.5 Then
  value = 0
  Else
  value = value * 2
  End If
  value = value Or saveflags
  xx& = put6502memory((savepc), value)
  If (value) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (value And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function
Public Function rola6502()
  saveflags = (P And &H1)
  P = (P And &HFE) Or ((A \ 128) And &H1)
  A = A * 2
  A = A Or saveflags
  If (A) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (A And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function
Public Function ror6502()
  saveflags = (P And &H1)
  adrmode (opcode)
  value = get6502memory(savepc)
      
  If value > 0 Then
    value = value
  End If
  P = (P And &HFE) Or (value And &H1)
  value = value \ 2
  If (saveflags) Then
    value = value Or &H80
  End If
  xx& = put6502memory((savepc), value And &HFF)
  If (value) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (value And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function
Public Function rora6502()
  saveflags = (P And &H1)
  P = (P And &HFE) Or (A And &H1)
  A = A \ 2
  
  If (saveflags) Then
    A = A Or &H80
  End If
  If (A) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (A And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function
Public Function rti6502()
  S = (S + 1) And &HFF
  P = gameImage(S + &H100) Or &H20
  S = (S + 1) And &HFF
  PC = gameImage(S + &H100)
  S = (S + 1) And &HFF
  PC = PC Or (gameImage(S + &H100) * 256)
End Function

Public Function rts6502()
  S = (S + 1) And &HFF
  PC = gameImage(S + &H100)
  S = (S + 1) And &HFF
  PC = PC Or (gameImage(S + &H100) * 256)
  PC = PC + 1
End Function

Public Function sbc6502()
  adrmode (opcode)
  value = get6502memory(savepc) Xor &HFF
  
  saveflags = (P And &H1)
  
  sum = A
  sum = (sum + value) And &HFF
  sum = (sum + (saveflags * 16)) And &HFF
  
  If ((sum > &H7F) Or (sum <= -&H80)) Then
    P = P Or &H40
  Else
    P = P And &HBF
  End If
  
  sum = A + (value + saveflags)
  
  If (sum > &HFF) Then
    P = P Or &H1
  Else
    P = P And &HFE
  End If
  
  A = sum And &HFF
  If (P And &H8) Then
    A = A - &H66
    P = P And &HFE
    If ((A And &HF) > &H9) Then
      A = A + &H6
    End If
    If ((A And &HF0) > &H90) Then
      A = A + &H60
      P = P Or H01
    End If
  Else
    clockticks6502 = clockticks6502 + 1
  End If
  
  If (A) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  
  If (A And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function
Public Function sec6502()
  P = P Or &H1
End Function
Public Function sed6502()
  P = P Or &H8
End Function
Public Function sei6502()
  P = P Or &H4
End Function
Public Function sta6502()
  adrmode (opcode)
  put6502memory (savepc), A
End Function
Public Function stx6502()
  adrmode (opcode)
  put6502memory (savepc), x
End Function
Public Function sty6502()
  adrmode (opcode)
  put6502memory (savepc), y
End Function
Public Function tax6502()
  x = A
  If (x) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (x And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function

Public Function tay6502()
  y = A
  If (y) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (y And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function
Public Function tsx6502()
  x = S
  If (x) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (x And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function

Public Function txa6502()
  A = x
  If (A) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (A And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function
Public Function txs6502()
  S = x
End Function

Public Function tya6502()
  A = y
  If (A) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (A And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function
Public Function bra6502()
  adrmode (opcode)
  PC = PC + savepc
  clockticks6502 = clockticks6502 + 1
End Function
Public Function dea6502()
  A = (A - 1) And &HFF
  
  If (A) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (A And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function
Public Function ina6502()
  A = (A + 1) And &HFF
      
  If (A) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (A And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function
Public Function phx6502()
  put6502memory &H100 + S, x
  S = (S - 1) And &HFF
End Function

Public Function plx6502()
  S = (S + 1) And &HFF
  x = gameImage(1 + &H100)
  If (x) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (x And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function
Public Function phy6502()
  xx& = put6502memory(&H100 + S, y)
  S = (S - 1) And &HFF
End Function
Public Function ply6502()
  S = (S + 1) And &HFF
  
  y = gameImage(S + &H100)
  If (y) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (y And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Function
Public Function stz6502()
  adrmode (opcode)
  put6502memory (savepc), 0
End Function

Public Function tsb6502()
  Dim tmp As Byte
      
  adrmode (opcode)
  tmp = get6502memory(savepc) Or A
  put6502memory (savepc), tmp
      
  If (tmp) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
End Function
Public Function trb6502()
  Dim tmp As Byte
      
  adrmode (opcode)
  tmp = get6502memory(savepc) And (A Xor &HFF)
  put6502memory (savepc), tmp
  
  If (tmp) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
End Function
