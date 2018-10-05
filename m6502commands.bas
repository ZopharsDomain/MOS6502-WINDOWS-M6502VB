Attribute VB_Name = "m6502Commands"

Public Function init6502()
      'TS: Start randomizer
      Randomize
      TICKS(&H0) = 7: instruction(&H0) = "brk6502": addrmode(&H0) = "implied6502"
      TICKS(&H1) = 6: instruction(&H1) = "ora6502": addrmode(&H1) = "indx6502"
      TICKS(&H2) = 2: instruction(&H2) = "nop6502": addrmode(&H2) = "implied6502"
      TICKS(&H3) = 2: instruction(&H3) = "nop6502": addrmode(&H3) = "implied6502"
      TICKS(&H4) = 3: instruction(&H4) = "tsb6502": addrmode(&H4) = "zp6502"
      TICKS(&H5) = 3: instruction(&H5) = "ora6502": addrmode(&H5) = "zp6502"
      TICKS(&H6) = 5: instruction(&H6) = "asl6502": addrmode(&H6) = "zp6502"
      TICKS(&H7) = 2: instruction(&H7) = "nop6502": addrmode(&H7) = "implied6502"
      TICKS(&H8) = 3: instruction(&H8) = "php6502": addrmode(&H8) = "implied6502"
      TICKS(&H9) = 3: instruction(&H9) = "ora6502": addrmode(&H9) = "immediate6502"
      TICKS(&HA) = 2: instruction(&HA) = "asla6502": addrmode(&HA) = "implied6502"
      TICKS(&HB) = 2: instruction(&HB) = "nop6502": addrmode(&HB) = "implied6502"
      TICKS(&HC) = 4: instruction(&HC) = "tsb6502": addrmode(&HC) = "abs6502"
      TICKS(&HD) = 4: instruction(&HD) = "ora6502": addrmode(&HD) = "abs6502"
      TICKS(&HE) = 6: instruction(&HE) = "asl6502": addrmode(&HE) = "abs6502"
      TICKS(&HF) = 2: instruction(&HF) = "nop6502": addrmode(&HF) = "implied6502"
      TICKS(&H10) = 2: instruction(&H10) = "bpl6502": addrmode(&H10) = "relative6502"
      TICKS(&H11) = 5: instruction(&H11) = "ora6502": addrmode(&H11) = "indy6502"
      TICKS(&H12) = 3: instruction(&H12) = "ora6502": addrmode(&H12) = "indzp6502"
      TICKS(&H13) = 2: instruction(&H13) = "nop6502": addrmode(&H13) = "implied6502"
      TICKS(&H14) = 3: instruction(&H14) = "trb6502": addrmode(&H14) = "zp6502"
      TICKS(&H15) = 4: instruction(&H15) = "ora6502": addrmode(&H15) = "zpx6502"
      TICKS(&H16) = 6: instruction(&H16) = "asl6502": addrmode(&H16) = "zpx6502"
      TICKS(&H17) = 2: instruction(&H17) = "nop6502": addrmode(&H17) = "implied6502"
      TICKS(&H18) = 2: instruction(&H18) = "clc6502": addrmode(&H18) = "implied6502"
      TICKS(&H19) = 4: instruction(&H19) = "ora6502": addrmode(&H19) = "absy6502"
      TICKS(&H1A) = 2: instruction(&H1A) = "ina6502": addrmode(&H1A) = "implied6502"
      TICKS(&H1B) = 2: instruction(&H1B) = "nop6502": addrmode(&H1B) = "implied6502"
      TICKS(&H1C) = 4: instruction(&H1C) = "trb6502": addrmode(&H1C) = "abs6502"
      TICKS(&H1D) = 4: instruction(&H1D) = "ora6502": addrmode(&H1D) = "absx6502"
      TICKS(&H1E) = 7: instruction(&H1E) = "asl6502": addrmode(&H1E) = "absx6502"
      TICKS(&H1F) = 2: instruction(&H1F) = "nop6502": addrmode(&H1F) = "implied6502"
      TICKS(&H20) = 6: instruction(&H20) = "jsr6502": addrmode(&H20) = "abs6502"
      TICKS(&H21) = 6: instruction(&H21) = "and6502": addrmode(&H21) = "indx6502"
      TICKS(&H22) = 2: instruction(&H22) = "nop6502": addrmode(&H22) = "implied6502"
      TICKS(&H23) = 2: instruction(&H23) = "nop6502": addrmode(&H23) = "implied6502"
      TICKS(&H24) = 3: instruction(&H24) = "bit6502": addrmode(&H24) = "zp6502"
      TICKS(&H25) = 3: instruction(&H25) = "and6502": addrmode(&H25) = "zp6502"
      TICKS(&H26) = 5: instruction(&H26) = "rol6502": addrmode(&H26) = "zp6502"
      TICKS(&H27) = 2: instruction(&H27) = "nop6502": addrmode(&H27) = "implied6502"
      TICKS(&H28) = 4: instruction(&H28) = "plp6502": addrmode(&H28) = "implied6502"
      TICKS(&H29) = 3: instruction(&H29) = "and6502": addrmode(&H29) = "immediate6502"
      TICKS(&H2A) = 2: instruction(&H2A) = "rola6502": addrmode(&H2A) = "implied6502"
      TICKS(&H2B) = 2: instruction(&H2B) = "nop6502": addrmode(&H2B) = "implied6502"
      TICKS(&H2C) = 4: instruction(&H2C) = "bit6502": addrmode(&H2C) = "abs6502"
      TICKS(&H2D) = 4: instruction(&H2D) = "and6502": addrmode(&H2D) = "abs6502"
      TICKS(&H2E) = 6: instruction(&H2E) = "rol6502": addrmode(&H2E) = "abs6502"
      TICKS(&H2F) = 2: instruction(&H2F) = "nop6502": addrmode(&H2F) = "implied6502"
      TICKS(&H30) = 2: instruction(&H30) = "bmi6502": addrmode(&H30) = "relative6502"
      TICKS(&H31) = 5: instruction(&H31) = "and6502": addrmode(&H31) = "indy6502"
      TICKS(&H32) = 3: instruction(&H32) = "and6502": addrmode(&H32) = "indzp6502"
      TICKS(&H33) = 2: instruction(&H33) = "nop6502": addrmode(&H33) = "implied6502"
      TICKS(&H34) = 4: instruction(&H34) = "bit6502": addrmode(&H34) = "zpx6502"
      TICKS(&H35) = 4: instruction(&H35) = "and6502": addrmode(&H35) = "zpx6502"
      TICKS(&H36) = 6: instruction(&H36) = "rol6502": addrmode(&H36) = "zpx6502"
      TICKS(&H37) = 2: instruction(&H37) = "nop6502": addrmode(&H37) = "implied6502"
      TICKS(&H38) = 2: instruction(&H38) = "sec6502": addrmode(&H38) = "implied6502"
      TICKS(&H39) = 4: instruction(&H39) = "and6502": addrmode(&H39) = "absy6502"
      TICKS(&H3A) = 2: instruction(&H3A) = "dea6502": addrmode(&H3A) = "implied6502"
      TICKS(&H3B) = 2: instruction(&H3B) = "nop6502": addrmode(&H3B) = "implied6502"
      TICKS(&H3C) = 4: instruction(&H3C) = "bit6502": addrmode(&H3C) = "absx6502"
      TICKS(&H3D) = 4: instruction(&H3D) = "and6502": addrmode(&H3D) = "absx6502"
      TICKS(&H3E) = 7: instruction(&H3E) = "rol6502": addrmode(&H3E) = "absx6502"
      TICKS(&H3F) = 2: instruction(&H3F) = "nop6502": addrmode(&H3F) = "implied6502"
      TICKS(&H40) = 6: instruction(&H40) = "rti6502": addrmode(&H40) = "implied6502"
      TICKS(&H41) = 6: instruction(&H41) = "eor6502": addrmode(&H41) = "indx6502"
      TICKS(&H42) = 2: instruction(&H42) = "nop6502": addrmode(&H42) = "implied6502"
      TICKS(&H43) = 2: instruction(&H43) = "nop6502": addrmode(&H43) = "implied6502"
      TICKS(&H44) = 2: instruction(&H44) = "nop6502": addrmode(&H44) = "implied6502"
      TICKS(&H45) = 3: instruction(&H45) = "eor6502": addrmode(&H45) = "zp6502"
      TICKS(&H46) = 5: instruction(&H46) = "lsr6502": addrmode(&H46) = "zp6502"
      TICKS(&H47) = 2: instruction(&H47) = "nop6502": addrmode(&H47) = "implied6502"
      TICKS(&H48) = 3: instruction(&H48) = "pha6502": addrmode(&H48) = "implied6502"
      TICKS(&H49) = 3: instruction(&H49) = "eor6502": addrmode(&H49) = "immediate6502"
      TICKS(&H4A) = 2: instruction(&H4A) = "lsra6502": addrmode(&H4A) = "implied6502"
      TICKS(&H4B) = 2: instruction(&H4B) = "nop6502": addrmode(&H4B) = "implied6502"
      TICKS(&H4C) = 3: instruction(&H4C) = "jmp6502": addrmode(&H4C) = "abs6502"
      TICKS(&H4D) = 4: instruction(&H4D) = "eor6502": addrmode(&H4D) = "abs6502"
      TICKS(&H4E) = 6: instruction(&H4E) = "lsr6502": addrmode(&H4E) = "abs6502"
      TICKS(&H4F) = 2: instruction(&H4F) = "nop6502": addrmode(&H4F) = "implied6502"
      TICKS(&H50) = 2: instruction(&H50) = "bvc6502": addrmode(&H50) = "relative6502"
      TICKS(&H51) = 5: instruction(&H51) = "eor6502": addrmode(&H51) = "indy6502"
      TICKS(&H52) = 3: instruction(&H52) = "eor6502": addrmode(&H52) = "indzp6502"
      TICKS(&H53) = 2: instruction(&H53) = "nop6502": addrmode(&H53) = "implied6502"
      TICKS(&H54) = 2: instruction(&H54) = "nop6502": addrmode(&H54) = "implied6502"
      TICKS(&H55) = 4: instruction(&H55) = "eor6502": addrmode(&H55) = "zpx6502"
      TICKS(&H56) = 6: instruction(&H56) = "lsr6502": addrmode(&H56) = "zpx6502"
      TICKS(&H57) = 2: instruction(&H57) = "nop6502": addrmode(&H57) = "implied6502"
      TICKS(&H58) = 2: instruction(&H58) = "cli6502": addrmode(&H58) = "implied6502"
      TICKS(&H59) = 4: instruction(&H59) = "eor6502": addrmode(&H59) = "absy6502"
      TICKS(&H5A) = 3: instruction(&H5A) = "phy6502": addrmode(&H5A) = "implied6502"
      TICKS(&H5B) = 2: instruction(&H5B) = "nop6502": addrmode(&H5B) = "implied6502"
      TICKS(&H5C) = 2: instruction(&H5C) = "nop6502": addrmode(&H5C) = "implied6502"
      TICKS(&H5D) = 4: instruction(&H5D) = "eor6502": addrmode(&H5D) = "absx6502"
      TICKS(&H5E) = 7: instruction(&H5E) = "lsr6502": addrmode(&H5E) = "absx6502"
      TICKS(&H5F) = 2: instruction(&H5F) = "nop6502": addrmode(&H5F) = "implied6502"
      TICKS(&H60) = 6: instruction(&H60) = "rts6502": addrmode(&H60) = "implied6502"
      TICKS(&H61) = 6: instruction(&H61) = "adc6502": addrmode(&H61) = "indx6502"
      TICKS(&H62) = 2: instruction(&H62) = "nop6502": addrmode(&H62) = "implied6502"
      TICKS(&H63) = 2: instruction(&H63) = "nop6502": addrmode(&H63) = "implied6502"
      TICKS(&H64) = 3: instruction(&H64) = "stz6502": addrmode(&H64) = "zp6502"
      TICKS(&H65) = 3: instruction(&H65) = "adc6502": addrmode(&H65) = "zp6502"
      TICKS(&H66) = 5: instruction(&H66) = "ror6502": addrmode(&H66) = "zp6502"
      TICKS(&H67) = 2: instruction(&H67) = "nop6502": addrmode(&H67) = "implied6502"
      TICKS(&H68) = 4: instruction(&H68) = "pla6502": addrmode(&H68) = "implied6502"
      TICKS(&H69) = 3: instruction(&H69) = "adc6502": addrmode(&H69) = "immediate6502"
      TICKS(&H6A) = 2: instruction(&H6A) = "rora6502": addrmode(&H6A) = "implied6502"
      TICKS(&H6B) = 2: instruction(&H6B) = "nop6502": addrmode(&H6B) = "implied6502"
      TICKS(&H6C) = 5: instruction(&H6C) = "jmp6502": addrmode(&H6C) = "indirect6502"
      TICKS(&H6D) = 4: instruction(&H6D) = "adc6502": addrmode(&H6D) = "abs6502"
      TICKS(&H6E) = 6: instruction(&H6E) = "ror6502": addrmode(&H6E) = "abs6502"
      TICKS(&H6F) = 2: instruction(&H6F) = "nop6502": addrmode(&H6F) = "implied6502"
      TICKS(&H70) = 2: instruction(&H70) = "bvs6502": addrmode(&H70) = "relative6502"
      TICKS(&H71) = 5: instruction(&H71) = "adc6502": addrmode(&H71) = "indy6502"
      TICKS(&H72) = 3: instruction(&H72) = "adc6502": addrmode(&H72) = "indzp6502"
      TICKS(&H73) = 2: instruction(&H73) = "nop6502": addrmode(&H73) = "implied6502"
      TICKS(&H74) = 4: instruction(&H74) = "stz6502": addrmode(&H74) = "zpx6502"
      TICKS(&H75) = 4: instruction(&H75) = "adc6502": addrmode(&H75) = "zpx6502"
      TICKS(&H76) = 6: instruction(&H76) = "ror6502": addrmode(&H76) = "zpx6502"
      TICKS(&H77) = 2: instruction(&H77) = "nop6502": addrmode(&H77) = "implied6502"
      TICKS(&H78) = 2: instruction(&H78) = "sei6502": addrmode(&H78) = "implied6502"
      TICKS(&H79) = 4: instruction(&H79) = "adc6502": addrmode(&H79) = "absy6502"
      TICKS(&H7A) = 4: instruction(&H7A) = "ply6502": addrmode(&H7A) = "implied6502"
      TICKS(&H7B) = 2: instruction(&H7B) = "nop6502": addrmode(&H7B) = "implied6502"
      TICKS(&H7C) = 6: instruction(&H7C) = "jmp6502": addrmode(&H7C) = "indabsx6502"
      TICKS(&H7D) = 4: instruction(&H7D) = "adc6502": addrmode(&H7D) = "absx6502"
      TICKS(&H7E) = 7: instruction(&H7E) = "ror6502": addrmode(&H7E) = "absx6502"
      TICKS(&H7F) = 2: instruction(&H7F) = "nop6502": addrmode(&H7F) = "implied6502"
      TICKS(&H80) = 2: instruction(&H80) = "bra6502": addrmode(&H80) = "relative6502"
      TICKS(&H81) = 6: instruction(&H81) = "sta6502": addrmode(&H81) = "indx6502"
      TICKS(&H82) = 2: instruction(&H82) = "nop6502": addrmode(&H82) = "implied6502"
      TICKS(&H83) = 2: instruction(&H83) = "nop6502": addrmode(&H83) = "implied6502"
      TICKS(&H84) = 2: instruction(&H84) = "sty6502": addrmode(&H84) = "zp6502"
      TICKS(&H85) = 2: instruction(&H85) = "sta6502": addrmode(&H85) = "zp6502"
      TICKS(&H86) = 2: instruction(&H86) = "stx6502": addrmode(&H86) = "zp6502"
      TICKS(&H87) = 2: instruction(&H87) = "nop6502": addrmode(&H87) = "implied6502"
      TICKS(&H88) = 2: instruction(&H88) = "dey6502": addrmode(&H88) = "implied6502"
      TICKS(&H89) = 2: instruction(&H89) = "bit6502": addrmode(&H89) = "immediate6502"
      TICKS(&H8A) = 2: instruction(&H8A) = "txa6502": addrmode(&H8A) = "implied6502"
      TICKS(&H8B) = 2: instruction(&H8B) = "nop6502": addrmode(&H8B) = "implied6502"
      TICKS(&H8C) = 4: instruction(&H8C) = "sty6502": addrmode(&H8C) = "abs6502"
      TICKS(&H8D) = 4: instruction(&H8D) = "sta6502": addrmode(&H8D) = "abs6502"
      TICKS(&H8E) = 4: instruction(&H8E) = "stx6502": addrmode(&H8E) = "abs6502"
      TICKS(&H8F) = 2: instruction(&H8F) = "nop6502": addrmode(&H8F) = "implied6502"
      TICKS(&H90) = 2: instruction(&H90) = "bcc6502": addrmode(&H90) = "relative6502"
      TICKS(&H91) = 6: instruction(&H91) = "sta6502": addrmode(&H91) = "indy6502"
      TICKS(&H92) = 3: instruction(&H92) = "sta6502": addrmode(&H92) = "indzp6502"
      TICKS(&H93) = 2: instruction(&H93) = "nop6502": addrmode(&H93) = "implied6502"
      TICKS(&H94) = 4: instruction(&H94) = "sty6502": addrmode(&H94) = "zpx6502"
      TICKS(&H95) = 4: instruction(&H95) = "sta6502": addrmode(&H95) = "zpx6502"
      TICKS(&H96) = 4: instruction(&H96) = "stx6502": addrmode(&H96) = "zpy6502"
      TICKS(&H97) = 2: instruction(&H97) = "nop6502": addrmode(&H97) = "implied6502"
      TICKS(&H98) = 2: instruction(&H98) = "tya6502": addrmode(&H98) = "implied6502"
      TICKS(&H99) = 5: instruction(&H99) = "sta6502": addrmode(&H99) = "absy6502"
      TICKS(&H9A) = 2: instruction(&H9A) = "txs6502": addrmode(&H9A) = "implied6502"
      TICKS(&H9B) = 2: instruction(&H9B) = "nop6502": addrmode(&H9B) = "implied6502"
      TICKS(&H9C) = 4: instruction(&H9C) = "stz6502": addrmode(&H9C) = "abs6502"
      TICKS(&H9D) = 5: instruction(&H9D) = "sta6502": addrmode(&H9D) = "absx6502"
      TICKS(&H9E) = 5: instruction(&H9E) = "stz6502": addrmode(&H9E) = "absx6502"
      TICKS(&H9F) = 2: instruction(&H9F) = "nop6502": addrmode(&H9F) = "implied6502"
      TICKS(&HA0) = 3: instruction(&HA0) = "ldy6502": addrmode(&HA0) = "immediate6502"
      TICKS(&HA1) = 6: instruction(&HA1) = "lda6502": addrmode(&HA1) = "indx6502"
      TICKS(&HA2) = 3: instruction(&HA2) = "ldx6502": addrmode(&HA2) = "immediate6502"
      TICKS(&HA3) = 2: instruction(&HA3) = "nop6502": addrmode(&HA3) = "implied6502"
      TICKS(&HA4) = 3: instruction(&HA4) = "ldy6502": addrmode(&HA4) = "zp6502"
      TICKS(&HA5) = 3: instruction(&HA5) = "lda6502": addrmode(&HA5) = "zp6502"
      TICKS(&HA6) = 3: instruction(&HA6) = "ldx6502": addrmode(&HA6) = "zp6502"
      TICKS(&HA7) = 2: instruction(&HA7) = "nop6502": addrmode(&HA7) = "implied6502"
      TICKS(&HA8) = 2: instruction(&HA8) = "tay6502": addrmode(&HA8) = "implied6502"
      TICKS(&HA9) = 3: instruction(&HA9) = "lda6502": addrmode(&HA9) = "immediate6502"
      TICKS(&HAA) = 2: instruction(&HAA) = "tax6502": addrmode(&HAA) = "implied6502"
      TICKS(&HAB) = 2: instruction(&HAB) = "nop6502": addrmode(&HAB) = "implied6502"
      TICKS(&HAC) = 4: instruction(&HAC) = "ldy6502": addrmode(&HAC) = "abs6502"
      TICKS(&HAD) = 4: instruction(&HAD) = "lda6502": addrmode(&HAD) = "abs6502"
      TICKS(&HAE) = 4: instruction(&HAE) = "ldx6502": addrmode(&HAE) = "abs6502"
      TICKS(&HAF) = 2: instruction(&HAF) = "nop6502": addrmode(&HAF) = "implied6502"
      TICKS(&HB0) = 2: instruction(&HB0) = "bcs6502": addrmode(&HB0) = "relative6502"
      TICKS(&HB1) = 5: instruction(&HB1) = "lda6502": addrmode(&HB1) = "indy6502"
      TICKS(&HB2) = 3: instruction(&HB2) = "lda6502": addrmode(&HB2) = "indzp6502"
      TICKS(&HB3) = 2: instruction(&HB3) = "nop6502": addrmode(&HB3) = "implied6502"
      TICKS(&HB4) = 4: instruction(&HB4) = "ldy6502": addrmode(&HB4) = "zpx6502"
      TICKS(&HB5) = 4: instruction(&HB5) = "lda6502": addrmode(&HB5) = "zpx6502"
      TICKS(&HB6) = 4: instruction(&HB6) = "ldx6502": addrmode(&HB6) = "zpy6502"
      TICKS(&HB7) = 2: instruction(&HB7) = "nop6502": addrmode(&HB7) = "implied6502"
      TICKS(&HB8) = 2: instruction(&HB8) = "clv6502": addrmode(&HB8) = "implied6502"
      TICKS(&HB9) = 4: instruction(&HB9) = "lda6502": addrmode(&HB9) = "absy6502"
      TICKS(&HBA) = 2: instruction(&HBA) = "tsx6502": addrmode(&HBA) = "implied6502"
      TICKS(&HBB) = 2: instruction(&HBB) = "nop6502": addrmode(&HBB) = "implied6502"
      TICKS(&HBC) = 4: instruction(&HBC) = "ldy6502": addrmode(&HBC) = "absx6502"
      TICKS(&HBD) = 4: instruction(&HBD) = "lda6502": addrmode(&HBD) = "absx6502"
      TICKS(&HBE) = 4: instruction(&HBE) = "ldx6502": addrmode(&HBE) = "absy6502"
      TICKS(&HBF) = 2: instruction(&HBF) = "nop6502": addrmode(&HBF) = "implied6502"
      TICKS(&HC0) = 3: instruction(&HC0) = "cpy6502": addrmode(&HC0) = "immediate6502"
      TICKS(&HC1) = 6: instruction(&HC1) = "cmp6502": addrmode(&HC1) = "indx6502"
      TICKS(&HC2) = 2: instruction(&HC2) = "nop6502": addrmode(&HC2) = "implied6502"
      TICKS(&HC3) = 2: instruction(&HC3) = "nop6502": addrmode(&HC3) = "implied6502"
      TICKS(&HC4) = 3: instruction(&HC4) = "cpy6502": addrmode(&HC4) = "zp6502"
      TICKS(&HC5) = 3: instruction(&HC5) = "cmp6502": addrmode(&HC5) = "zp6502"
      TICKS(&HC6) = 5: instruction(&HC6) = "dec6502": addrmode(&HC6) = "zp6502"
      TICKS(&HC7) = 2: instruction(&HC7) = "nop6502": addrmode(&HC7) = "implied6502"
      TICKS(&HC8) = 2: instruction(&HC8) = "iny6502": addrmode(&HC8) = "implied6502"
      TICKS(&HC9) = 3: instruction(&HC9) = "cmp6502": addrmode(&HC9) = "immediate6502"
      TICKS(&HCA) = 2: instruction(&HCA) = "dex6502": addrmode(&HCA) = "implied6502"
      TICKS(&HCB) = 2: instruction(&HCB) = "nop6502": addrmode(&HCB) = "implied6502"
      TICKS(&HCC) = 4: instruction(&HCC) = "cpy6502": addrmode(&HCC) = "abs6502"
      TICKS(&HCD) = 4: instruction(&HCD) = "cmp6502": addrmode(&HCD) = "abs6502"
      TICKS(&HCE) = 6: instruction(&HCE) = "dec6502": addrmode(&HCE) = "abs6502"
      TICKS(&HCF) = 2: instruction(&HCF) = "nop6502": addrmode(&HCF) = "implied6502"
      TICKS(&HD0) = 2: instruction(&HD0) = "bne6502": addrmode(&HD0) = "relative6502"
      TICKS(&HD1) = 5: instruction(&HD1) = "cmp6502": addrmode(&HD1) = "indy6502"
      TICKS(&HD2) = 3: instruction(&HD2) = "cmp6502": addrmode(&HD2) = "indzp6502"
      TICKS(&HD3) = 2: instruction(&HD3) = "nop6502": addrmode(&HD3) = "implied6502"
      TICKS(&HD4) = 2: instruction(&HD4) = "nop6502": addrmode(&HD4) = "implied6502"
      TICKS(&HD5) = 4: instruction(&HD5) = "cmp6502": addrmode(&HD5) = "zpx6502"
      TICKS(&HD6) = 6: instruction(&HD6) = "dec6502": addrmode(&HD6) = "zpx6502"
      TICKS(&HD7) = 2: instruction(&HD7) = "nop6502": addrmode(&HD7) = "implied6502"
      TICKS(&HD8) = 2: instruction(&HD8) = "cld6502": addrmode(&HD8) = "implied6502"
      TICKS(&HD9) = 4: instruction(&HD9) = "cmp6502": addrmode(&HD9) = "absy6502"
      TICKS(&HDA) = 3: instruction(&HDA) = "phx6502": addrmode(&HDA) = "implied6502"
      TICKS(&HDB) = 2: instruction(&HDB) = "nop6502": addrmode(&HDB) = "implied6502"
      TICKS(&HDC) = 2: instruction(&HDC) = "nop6502": addrmode(&HDC) = "implied6502"
      TICKS(&HDD) = 4: instruction(&HDD) = "cmp6502": addrmode(&HDD) = "absx6502"
      TICKS(&HDE) = 7: instruction(&HDE) = "dec6502": addrmode(&HDE) = "absx6502"
      TICKS(&HDF) = 2: instruction(&HDF) = "nop6502": addrmode(&HDF) = "implied6502"
      TICKS(&HE0) = 3: instruction(&HE0) = "cpx6502": addrmode(&HE0) = "immediate6502"
      TICKS(&HE1) = 6: instruction(&HE1) = "sbc6502": addrmode(&HE1) = "indx6502"
      TICKS(&HE2) = 2: instruction(&HE2) = "nop6502": addrmode(&HE2) = "implied6502"
      TICKS(&HE3) = 2: instruction(&HE3) = "nop6502": addrmode(&HE3) = "implied6502"
      TICKS(&HE4) = 3: instruction(&HE4) = "cpx6502": addrmode(&HE4) = "zp6502"
      TICKS(&HE5) = 3: instruction(&HE5) = "sbc6502": addrmode(&HE5) = "zp6502"
      TICKS(&HE6) = 5: instruction(&HE6) = "inc6502": addrmode(&HE6) = "zp6502"
      TICKS(&HE7) = 2: instruction(&HE7) = "nop6502": addrmode(&HE7) = "implied6502"
      TICKS(&HE8) = 2: instruction(&HE8) = "inx6502": addrmode(&HE8) = "implied6502"
      TICKS(&HE9) = 3: instruction(&HE9) = "sbc6502": addrmode(&HE9) = "immediate6502"
      TICKS(&HEA) = 2: instruction(&HEA) = "nop6502": addrmode(&HEA) = "implied6502"
      TICKS(&HEB) = 2: instruction(&HEB) = "nop6502": addrmode(&HEB) = "implied6502"
      TICKS(&HEC) = 4: instruction(&HEC) = "cpx6502": addrmode(&HEC) = "abs6502"
      TICKS(&HED) = 4: instruction(&HED) = "sbc6502": addrmode(&HED) = "abs6502"
      TICKS(&HEE) = 6: instruction(&HEE) = "inc6502": addrmode(&HEE) = "abs6502"
      TICKS(&HEF) = 2: instruction(&HEF) = "nop6502": addrmode(&HEF) = "implied6502"
      TICKS(&HF0) = 2: instruction(&HF0) = "beq6502": addrmode(&HF0) = "relative6502"
      TICKS(&HF1) = 5: instruction(&HF1) = "sbc6502": addrmode(&HF1) = "indy6502"
      TICKS(&HF2) = 3: instruction(&HF2) = "sbc6502": addrmode(&HF2) = "indzp6502"
      TICKS(&HF3) = 2: instruction(&HF3) = "nop6502": addrmode(&HF3) = "implied6502"
      TICKS(&HF4) = 2: instruction(&HF4) = "nop6502": addrmode(&HF4) = "implied6502"
      TICKS(&HF5) = 4: instruction(&HF5) = "sbc6502": addrmode(&HF5) = "zpx6502"
      TICKS(&HF6) = 6: instruction(&HF6) = "inc6502": addrmode(&HF6) = "zpx6502"
      TICKS(&HF7) = 2: instruction(&HF7) = "nop6502": addrmode(&HF7) = "implied6502"
      TICKS(&HF8) = 2: instruction(&HF8) = "sed6502": addrmode(&HF8) = "implied6502"
      TICKS(&HF9) = 4: instruction(&HF9) = "sbc6502": addrmode(&HF9) = "absy6502"
      TICKS(&HFA) = 4: instruction(&HFA) = "plx6502": addrmode(&HFA) = "implied6502"
      TICKS(&HFB) = 2: instruction(&HFB) = "nop6502": addrmode(&HFB) = "implied6502"
      TICKS(&HFC) = 2: instruction(&HFC) = "nop6502": addrmode(&HFC) = "implied6502"
      TICKS(&HFD) = 4: instruction(&HFD) = "sbc6502": addrmode(&HFD) = "absx6502"
      TICKS(&HFE) = 7: instruction(&HFE) = "inc6502": addrmode(&HFE) = "absx6502"
      TICKS(&HFF) = 2: instruction(&HFF) = "nop6502": addrmode(&HFF) = "implied6502"
End Function