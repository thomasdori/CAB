<%
' ***************************************************************
'
' 8 pt Lucida Sans Typewriter Bitmap Font
' Put together by Tony Stefano
'
' ***************************************************************

' Definitions of chars 32-126

' Font and Letter must be defined to work correctly
Dim Font
Dim Letter(12)

Set Font = Server.CreateObject("Scripting.Dictionary")

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0000000"
Letter(5) = "0000000"
Letter(6) = "0000000"
Letter(7) = "0000000"
Letter(8) = "0000000"
Letter(9) = "0000000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add " ",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0010000"
Letter(3) = "0010000"
Letter(4) = "0010000"
Letter(5) = "0010000"
Letter(6) = "0010000"
Letter(7) = "0010000"
Letter(8) = "0000000"
Letter(9) = "0010000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "!",Letter

Letter(0) = "0000000"
Letter(1) = "0101000"
Letter(2) = "0101000"
Letter(3) = "0101000"
Letter(4) = "0000000"
Letter(5) = "0000000"
Letter(6) = "0000000"
Letter(7) = "0000000"
Letter(8) = "0000000"
Letter(9) = "0000000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add Chr(34),Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0010100"
Letter(3) = "0010100"
Letter(4) = "1111110"
Letter(5) = "0010100"
Letter(6) = "0101000"
Letter(7) = "1111110"
Letter(8) = "0101000"
Letter(9) = "0101000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "#",Letter

Letter(0) = "0000000"
Letter(1) = "0001000"
Letter(2) = "0011110"
Letter(3) = "0101000"
Letter(4) = "0101000"
Letter(5) = "0011000"
Letter(6) = "0001100"
Letter(7) = "0001010"
Letter(8) = "0001010"
Letter(9) = "0111100"
Letter(10) = "0001000"
Letter(11) = "0000000"

Font.Add "$",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0100001"
Letter(3) = "1010010"
Letter(4) = "1010100"
Letter(5) = "1011010"
Letter(6) = "0101101"
Letter(7) = "0010101"
Letter(8) = "0100101"
Letter(9) = "1000010"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "%",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0001000"
Letter(3) = "0010100"
Letter(4) = "0010100"
Letter(5) = "0011000"
Letter(6) = "0101001"
Letter(7) = "1000101"
Letter(8) = "1000010"
Letter(9) = "0111101"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "&",Letter

Letter(0) = "0000000"
Letter(1) = "0010000"
Letter(2) = "0010000"
Letter(3) = "0010000"
Letter(4) = "0000000"
Letter(5) = "0000000"
Letter(6) = "0000000"
Letter(7) = "0000000"
Letter(8) = "0000000"
Letter(9) = "0000000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "'",Letter

Letter(0) = "0000000"
Letter(1) = "0000010"
Letter(2) = "0000100"
Letter(3) = "0001000"
Letter(4) = "0010000"
Letter(5) = "0010000"
Letter(6) = "0010000"
Letter(7) = "0010000"
Letter(8) = "0010000"
Letter(9) = "0001000"
Letter(10) = "0000100"
Letter(11) = "0000010"

Font.Add "(",Letter

Letter(0) = "0000000"
Letter(1) = "0100000"
Letter(2) = "0010000"
Letter(3) = "0001000"
Letter(4) = "0000100"
Letter(5) = "0000100"
Letter(6) = "0000100"
Letter(7) = "0000100"
Letter(8) = "0000100"
Letter(9) = "0001000"
Letter(10) = "0010000"
Letter(11) = "0100000"

Font.Add ")",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0001000"
Letter(3) = "0101010"
Letter(4) = "0010100"
Letter(5) = "0010100"
Letter(6) = "0000000"
Letter(7) = "0000000"
Letter(8) = "0000000"
Letter(9) = "0000000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "*",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0000000"
Letter(5) = "0010000"
Letter(6) = "0010000"
Letter(7) = "1111100"
Letter(8) = "0010000"
Letter(9) = "0010000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "+",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0000000"
Letter(5) = "0000000"
Letter(6) = "0000000"
Letter(7) = "0000000"
Letter(8) = "0011000"
Letter(9) = "0011000"
Letter(10) = "0001000"
Letter(11) = "0010000"

Font.Add ",",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0000000"
Letter(5) = "0000000"
Letter(6) = "0111110"
Letter(7) = "0000000"
Letter(8) = "0000000"
Letter(9) = "0000000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "-",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0000000"
Letter(5) = "0000000"
Letter(6) = "0000000"
Letter(7) = "0000000"
Letter(8) = "0011000"
Letter(9) = "0011000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add ".",Letter

Letter(0) = "0000000"
Letter(1) = "0000100"
Letter(2) = "0000100"
Letter(3) = "0001000"
Letter(4) = "0001000"
Letter(5) = "0001000"
Letter(6) = "0010000"
Letter(7) = "0010000"
Letter(8) = "0100000"
Letter(9) = "0100000"
Letter(10) = "0100000"
Letter(11) = "0000000"

Font.Add "/",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0011100"
Letter(3) = "0100010"
Letter(4) = "0100010"
Letter(5) = "0100010"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0100010"
Letter(9) = "0011100"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "0",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0011000"
Letter(3) = "0101000"
Letter(4) = "0001000"
Letter(5) = "0001000"
Letter(6) = "0001000"
Letter(7) = "0001000"
Letter(8) = "0001000"
Letter(9) = "0111110"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "1",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0111100"
Letter(3) = "0000010"
Letter(4) = "0000010"
Letter(5) = "0000100"
Letter(6) = "0001000"
Letter(7) = "0010000"
Letter(8) = "0100000"
Letter(9) = "0111110"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "2",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0111100"
Letter(3) = "0000010"
Letter(4) = "0000010"
Letter(5) = "0011100"
Letter(6) = "0000010"
Letter(7) = "0000010"
Letter(8) = "0000010"
Letter(9) = "0111100"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "3",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000100"
Letter(3) = "0001100"
Letter(4) = "0010100"
Letter(5) = "0010100"
Letter(6) = "0100100"
Letter(7) = "0111110"
Letter(8) = "0000100"
Letter(9) = "0000100"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "4",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0111100"
Letter(3) = "0100000"
Letter(4) = "0100000"
Letter(5) = "0111000"
Letter(6) = "0000100"
Letter(7) = "0000100"
Letter(8) = "0000100"
Letter(9) = "0111000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "5",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0011110"
Letter(3) = "0100000"
Letter(4) = "0100000"
Letter(5) = "0101100"
Letter(6) = "0110010"
Letter(7) = "0100010"
Letter(8) = "0100010"
Letter(9) = "0011100"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "6",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0111110"
Letter(3) = "0000010"
Letter(4) = "0000100"
Letter(5) = "0001000"
Letter(6) = "0010000"
Letter(7) = "0010000"
Letter(8) = "0100000"
Letter(9) = "0100000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "7",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0011100"
Letter(3) = "0100010"
Letter(4) = "0100010"
Letter(5) = "0011100"
Letter(6) = "0100110"
Letter(7) = "0100010"
Letter(8) = "0100010"
Letter(9) = "0011100"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "8",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0011100"
Letter(3) = "0100010"
Letter(4) = "0100010"
Letter(5) = "0100110"
Letter(6) = "0011010"
Letter(7) = "0000010"
Letter(8) = "0000010"
Letter(9) = "0111100"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "9",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0011000"
Letter(5) = "0011000"
Letter(6) = "0000000"
Letter(7) = "0000000"
Letter(8) = "0011000"
Letter(9) = "0011000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add ":",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0011000"
Letter(5) = "0011000"
Letter(6) = "0000000"
Letter(7) = "0000000"
Letter(8) = "0011000"
Letter(9) = "0011000"
Letter(10) = "0001000"
Letter(11) = "0010000"

Font.Add ";",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0000010"
Letter(5) = "0001100"
Letter(6) = "0110000"
Letter(7) = "0110000"
Letter(8) = "0001100"
Letter(9) = "0000010"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "<",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0000000"
Letter(5) = "1111110"
Letter(6) = "0000000"
Letter(7) = "1111110"
Letter(8) = "0000000"
Letter(9) = "0000000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "=",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "1000000"
Letter(5) = "0110000"
Letter(6) = "0001100"
Letter(7) = "0001100"
Letter(8) = "0110000"
Letter(9) = "1000000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add ">",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0111100"
Letter(3) = "0100010"
Letter(4) = "0000010"
Letter(5) = "0000100"
Letter(6) = "0001000"
Letter(7) = "0001000"
Letter(8) = "0000000"
Letter(9) = "0001000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "?",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0011100"
Letter(3) = "0101010"
Letter(4) = "1010110"
Letter(5) = "1010010"
Letter(6) = "1010110"
Letter(7) = "1001011"
Letter(8) = "0100000"
Letter(9) = "0011100"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "@",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0001000"
Letter(3) = "0010100"
Letter(4) = "0010100"
Letter(5) = "0010100"
Letter(6) = "0100010"
Letter(7) = "0111110"
Letter(8) = "0100010"
Letter(9) = "1000001"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "A",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0111100"
Letter(3) = "0100010"
Letter(4) = "0100010"
Letter(5) = "0111100"
Letter(6) = "0100100"
Letter(7) = "0100010"
Letter(8) = "0100010"
Letter(9) = "0111100"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "B",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0001110"
Letter(3) = "0010000"
Letter(4) = "0100000"
Letter(5) = "0100000"
Letter(6) = "0100000"
Letter(7) = "0100000"
Letter(8) = "0010000"
Letter(9) = "0001110"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "C",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0111000"
Letter(3) = "0100100"
Letter(4) = "0100010"
Letter(5) = "0100010"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0100100"
Letter(9) = "0111000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "D",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0111110"
Letter(3) = "0100000"
Letter(4) = "0100000"
Letter(5) = "0111100"
Letter(6) = "0100000"
Letter(7) = "0100000"
Letter(8) = "0100000"
Letter(9) = "0111110"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "E",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0111110"
Letter(3) = "0100000"
Letter(4) = "0100000"
Letter(5) = "0111100"
Letter(6) = "0100000"
Letter(7) = "0100000"
Letter(8) = "0100000"
Letter(9) = "0100000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "F",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0001110"
Letter(3) = "0010000"
Letter(4) = "0100000"
Letter(5) = "0100000"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0010010"
Letter(9) = "0001110"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "G",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0100010"
Letter(3) = "0100010"
Letter(4) = "0100010"
Letter(5) = "0111110"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0100010"
Letter(9) = "0100010"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "H",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0111110"
Letter(3) = "0001000"
Letter(4) = "0001000"
Letter(5) = "0001000"
Letter(6) = "0001000"
Letter(7) = "0001000"
Letter(8) = "0001000"
Letter(9) = "0111110"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "I",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0011100"
Letter(3) = "0000100"
Letter(4) = "0000100"
Letter(5) = "0000100"
Letter(6) = "0000100"
Letter(7) = "0000100"
Letter(8) = "0000100"
Letter(9) = "0111000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "J",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0100010"
Letter(3) = "0100100"
Letter(4) = "0101000"
Letter(5) = "0110000"
Letter(6) = "0110000"
Letter(7) = "0101000"
Letter(8) = "0100100"
Letter(9) = "0100010"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "K",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0100000"
Letter(3) = "0100000"
Letter(4) = "0100000"
Letter(5) = "0100000"
Letter(6) = "0100000"
Letter(7) = "0100000"
Letter(8) = "0100000"
Letter(9) = "0111110"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "L",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "1000010"
Letter(3) = "1000010"
Letter(4) = "1100110"
Letter(5) = "1100110"
Letter(6) = "1011010"
Letter(7) = "1011010"
Letter(8) = "1000010"
Letter(9) = "1000010"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "M",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0100010"
Letter(3) = "0110010"
Letter(4) = "0110010"
Letter(5) = "0101010"
Letter(6) = "0101010"
Letter(7) = "0100110"
Letter(8) = "0100110"
Letter(9) = "0100010"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "N",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0001000"
Letter(3) = "0010100"
Letter(4) = "0100010"
Letter(5) = "0100010"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0010100"
Letter(9) = "0001000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "O",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0111100"
Letter(3) = "0100010"
Letter(4) = "0100010"
Letter(5) = "0100100"
Letter(6) = "0111000"
Letter(7) = "0100000"
Letter(8) = "0100000"
Letter(9) = "0100000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "P",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0001000"
Letter(3) = "0010100"
Letter(4) = "0100010"
Letter(5) = "0100010"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0010100"
Letter(9) = "0001100"
Letter(10) = "0000010"
Letter(11) = "0000000"

Font.Add "Q",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0111100"
Letter(3) = "0100010"
Letter(4) = "0100010"
Letter(5) = "0100010"
Letter(6) = "0111100"
Letter(7) = "0100100"
Letter(8) = "0100010"
Letter(9) = "0100010"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "R",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0011110"
Letter(3) = "0100000"
Letter(4) = "0100000"
Letter(5) = "0011000"
Letter(6) = "0000100"
Letter(7) = "0000010"
Letter(8) = "0000010"
Letter(9) = "0111100"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "S",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "1111111"
Letter(3) = "0001000"
Letter(4) = "0001000"
Letter(5) = "0001000"
Letter(6) = "0001000"
Letter(7) = "0001000"
Letter(8) = "0001000"
Letter(9) = "0001000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "T",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0100010"
Letter(3) = "0100010"
Letter(4) = "0100010"
Letter(5) = "0100010"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0100010"
Letter(9) = "0011100"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "U",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "1000001"
Letter(3) = "0100010"
Letter(4) = "0100010"
Letter(5) = "0100010"
Letter(6) = "0010100"
Letter(7) = "0010100"
Letter(8) = "0010100"
Letter(9) = "0001000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "V",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "1000001"
Letter(3) = "1000001"
Letter(4) = "1001001"
Letter(5) = "1001001"
Letter(6) = "0110110"
Letter(7) = "0110110"
Letter(8) = "0100010"
Letter(9) = "0100010"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "W",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0100010"
Letter(3) = "0010100"
Letter(4) = "0010100"
Letter(5) = "0001000"
Letter(6) = "0001000"
Letter(7) = "0010100"
Letter(8) = "0100010"
Letter(9) = "0100010"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "X",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0100010"
Letter(3) = "0100010"
Letter(4) = "0010100"
Letter(5) = "0010100"
Letter(6) = "0001000"
Letter(7) = "0001000"
Letter(8) = "0001000"
Letter(9) = "0001000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "Y",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0111110"
Letter(3) = "0000010"
Letter(4) = "0000100"
Letter(5) = "0001000"
Letter(6) = "0001000"
Letter(7) = "0010000"
Letter(8) = "0100000"
Letter(9) = "0111110"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "Z",Letter

Letter(0) = "0000000"
Letter(1) = "0011100"
Letter(2) = "0010000"
Letter(3) = "0010000"
Letter(4) = "0010000"
Letter(5) = "0010000"
Letter(6) = "0010000"
Letter(7) = "0010000"
Letter(8) = "0010000"
Letter(9) = "0010000"
Letter(10) = "0010000"
Letter(11) = "0011100"

Font.Add "[",Letter

Letter(0) = "0000000"
Letter(1) = "1000000"
Letter(2) = "1000000"
Letter(3) = "0100000"
Letter(4) = "0010000"
Letter(5) = "0010000"
Letter(6) = "0001000"
Letter(7) = "0001000"
Letter(8) = "0000100"
Letter(9) = "0000010"
Letter(10) = "0000010"
Letter(11) = "0000000"

Font.Add "\",Letter

Letter(0) = "0000000"
Letter(1) = "0011100"
Letter(2) = "0000100"
Letter(3) = "0000100"
Letter(4) = "0000100"
Letter(5) = "0000100"
Letter(6) = "0000100"
Letter(7) = "0000100"
Letter(8) = "0000100"
Letter(9) = "0000100"
Letter(10) = "0000100"
Letter(11) = "0011100"

Font.Add "]",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0001000"
Letter(3) = "0001000"
Letter(4) = "0010100"
Letter(5) = "0010100"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0000000"
Letter(9) = "0000000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "^",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0000000"
Letter(5) = "0000000"
Letter(6) = "0000000"
Letter(7) = "0000000"
Letter(8) = "0000000"
Letter(9) = "0000000"
Letter(10) = "1111111"
Letter(11) = "0000000"

Font.Add "_",Letter

Letter(0) = "0000000"
Letter(1) = "0010000"
Letter(2) = "0001000"
Letter(3) = "0000000"
Letter(4) = "0000000"
Letter(5) = "0000000"
Letter(6) = "0000000"
Letter(7) = "0000000"
Letter(8) = "0000000"
Letter(9) = "0000000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "`",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0111100"
Letter(5) = "0000010"
Letter(6) = "0011110"
Letter(7) = "0100010"
Letter(8) = "0100110"
Letter(9) = "0011001"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "a",Letter

Letter(0) = "0000000"
Letter(1) = "0100000"
Letter(2) = "0100000"
Letter(3) = "0100000"
Letter(4) = "0101100"
Letter(5) = "0110010"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0110010"
Letter(9) = "0101100"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "b",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0011110"
Letter(5) = "0100000"
Letter(6) = "0100000"
Letter(7) = "0100000"
Letter(8) = "0100000"
Letter(9) = "0011110"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "c",Letter

Letter(0) = "0000000"
Letter(1) = "0000010"
Letter(2) = "0000010"
Letter(3) = "0000010"
Letter(4) = "0011010"
Letter(5) = "0100110"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0100110"
Letter(9) = "0011010"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "d",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0011100"
Letter(5) = "0100010"
Letter(6) = "0111110"
Letter(7) = "0100000"
Letter(8) = "0100000"
Letter(9) = "0011110"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "e",Letter

Letter(0) = "0000000"
Letter(1) = "0000111"
Letter(2) = "0001000"
Letter(3) = "0001000"
Letter(4) = "0111111"
Letter(5) = "0001000"
Letter(6) = "0001000"
Letter(7) = "0001000"
Letter(8) = "0001000"
Letter(9) = "0001000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "f",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0011010"
Letter(5) = "0100110"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0100110"
Letter(9) = "0011010"
Letter(10) = "0000010"
Letter(11) = "0111100"

Font.Add "g",Letter

Letter(0) = "0000000"
Letter(1) = "0100000"
Letter(2) = "0100000"
Letter(3) = "0100000"
Letter(4) = "0101100"
Letter(5) = "0110010"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0100010"
Letter(9) = "0100010"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "h",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000100"
Letter(3) = "0000000"
Letter(4) = "0011100"
Letter(5) = "0000100"
Letter(6) = "0000100"
Letter(7) = "0000100"
Letter(8) = "0000100"
Letter(9) = "0000100"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "i",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000100"
Letter(3) = "0000000"
Letter(4) = "0111100"
Letter(5) = "0000100"
Letter(6) = "0000100"
Letter(7) = "0000100"
Letter(8) = "0000100"
Letter(9) = "0000100"
Letter(10) = "0000100"
Letter(11) = "0111000"

Font.Add "j",Letter

Letter(0) = "0000000"
Letter(1) = "0100000"
Letter(2) = "0100000"
Letter(3) = "0100000"
Letter(4) = "0100100"
Letter(5) = "0101000"
Letter(6) = "0110000"
Letter(7) = "0101000"
Letter(8) = "0100100"
Letter(9) = "0100010"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "k",Letter

Letter(0) = "0000000"
Letter(1) = "0111000"
Letter(2) = "0001000"
Letter(3) = "0001000"
Letter(4) = "0001000"
Letter(5) = "0001000"
Letter(6) = "0001000"
Letter(7) = "0001000"
Letter(8) = "0001000"
Letter(9) = "0001000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "l",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "1011011"
Letter(5) = "1101101"
Letter(6) = "1001001"
Letter(7) = "1001001"
Letter(8) = "1001001"
Letter(9) = "1001001"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "m",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0101100"
Letter(5) = "0110010"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0100010"
Letter(9) = "0100010"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "n",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0011100"
Letter(5) = "0100010"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0100010"
Letter(9) = "0011100"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "o",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0101100"
Letter(5) = "0110010"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0110010"
Letter(9) = "0101100"
Letter(10) = "0100000"
Letter(11) = "0100000"

Font.Add "p",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0011010"
Letter(5) = "0100110"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0100110"
Letter(9) = "0011010"
Letter(10) = "0000010"
Letter(11) = "0000010"

Font.Add "q",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0101110"
Letter(5) = "0110010"
Letter(6) = "0100000"
Letter(7) = "0100000"
Letter(8) = "0100000"
Letter(9) = "0100000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "r",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0011110"
Letter(5) = "0100000"
Letter(6) = "0011000"
Letter(7) = "0000100"
Letter(8) = "0000010"
Letter(9) = "0111100"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "s",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0010000"
Letter(4) = "0111110"
Letter(5) = "0010000"
Letter(6) = "0010000"
Letter(7) = "0010000"
Letter(8) = "0010000"
Letter(9) = "0001110"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "t",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0100010"
Letter(5) = "0100010"
Letter(6) = "0100010"
Letter(7) = "0100010"
Letter(8) = "0100110"
Letter(9) = "0011010"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "u",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0100010"
Letter(5) = "0100010"
Letter(6) = "0010100"
Letter(7) = "0010100"
Letter(8) = "0010100"
Letter(9) = "0001000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "v",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "1000001"
Letter(5) = "1001001"
Letter(6) = "1011101"
Letter(7) = "0110110"
Letter(8) = "0100010"
Letter(9) = "0100010"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "w",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0100010"
Letter(5) = "0010100"
Letter(6) = "0001000"
Letter(7) = "0001000"
Letter(8) = "0010100"
Letter(9) = "0100010"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "x",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0100010"
Letter(5) = "0100010"
Letter(6) = "0010100"
Letter(7) = "0010100"
Letter(8) = "0001000"
Letter(9) = "0001000"
Letter(10) = "0001000"
Letter(11) = "0110000"

Font.Add "y",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0111110"
Letter(5) = "0000100"
Letter(6) = "0001000"
Letter(7) = "0010000"
Letter(8) = "0100000"
Letter(9) = "0111110"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "z",Letter

Letter(0) = "0000000"
Letter(1) = "0001100"
Letter(2) = "0010000"
Letter(3) = "0010000"
Letter(4) = "0010000"
Letter(5) = "0010000"
Letter(6) = "0100000"
Letter(7) = "0010000"
Letter(8) = "0010000"
Letter(9) = "0010000"
Letter(10) = "0010000"
Letter(11) = "0001100"

Font.Add "{",Letter

Letter(0) = "0000000"
Letter(1) = "0001000"
Letter(2) = "0001000"
Letter(3) = "0001000"
Letter(4) = "0001000"
Letter(5) = "0001000"
Letter(6) = "0001000"
Letter(7) = "0001000"
Letter(8) = "0001000"
Letter(9) = "0001000"
Letter(10) = "0001000"
Letter(11) = "0000000"

Font.Add "|",Letter

Letter(0) = "0000000"
Letter(1) = "0011000"
Letter(2) = "0000100"
Letter(3) = "0000100"
Letter(4) = "0000100"
Letter(5) = "0000100"
Letter(6) = "0000010"
Letter(7) = "0000100"
Letter(8) = "0000100"
Letter(9) = "0000100"
Letter(10) = "0000100"
Letter(11) = "0011000"

Font.Add "}",Letter

Letter(0) = "0000000"
Letter(1) = "0000000"
Letter(2) = "0000000"
Letter(3) = "0000000"
Letter(4) = "0000000"
Letter(5) = "0000000"
Letter(6) = "0110010"
Letter(7) = "1001100"
Letter(8) = "0000000"
Letter(9) = "0000000"
Letter(10) = "0000000"
Letter(11) = "0000000"

Font.Add "~",Letter
%>