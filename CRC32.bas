Attribute VB_Name = "Module"
' Attribute VB_Name = "ModStdIO"
Option Explicit

Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, _
    lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, _
    lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long

Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_INPUT_HANDLE = -10&

' // Then declare this array variable Crc32Table
Private crc32Table(256) As Long

Function ReadStdIn(Optional ByVal NumBytes As Long = -1) As String
    Dim StdIn As Long
    Dim Result As Long
    Dim Buffer As String
    Dim BytesRead As Long
    StdIn = GetStdHandle(STD_INPUT_HANDLE)
    Buffer = Space$(1024)
    Do
        Result = ReadFile(StdIn, ByVal Buffer, Len(Buffer), BytesRead, ByVal 0&)
        If Result = 0 Then
            Err.Raise 1001, , "Unable to read from standard input"
        End If
        ReadStdIn = ReadStdIn & Left$(Buffer, BytesRead)
    Loop Until BytesRead < Len(Buffer)
End Function

Sub WriteStdOut(ByVal Text As String)
    Dim StdOut As Long
    Dim Result As Long
    Dim BytesWritten As Long
    StdOut = GetStdHandle(STD_OUTPUT_HANDLE)
    Result = WriteFile(StdOut, ByVal Text, Len(Text), BytesWritten, ByVal 0&)
    If Result = 0 Then
        Err.Raise 1001, , "Unable to write to standard output"
    ElseIf BytesWritten < Len(Text) Then
        Err.Raise 1002, , "Incomplete write operation"
    End If
End Sub

Function Byte4ToLong(ByRef bArray() As Byte, offset As Integer) As Long
    Dim iReturn As Long
    Dim i As Integer

    iReturn = 0
    For i = 0 To 4 - 1
        WriteStdOut ("bArray[" & offset + i & "]=0x" & Right$("00" & Hex$(bArray(offset + i)), 2) & vbCrLf)
        iReturn = iReturn + bArray(offset + i) * (256 ^ ((4 - 1) - i))
    Next i

    Byte4ToLong = iReturn

End Function

Sub InitCrc32(Optional ByVal dwPolynomial As Long = &H4C11DB7)
    crc32Table(0) = &H0
    crc32Table(1) = &H77073096
    crc32Table(2) = &HEE0E612C
    crc32Table(3) = &H990951BA
    crc32Table(4) = &H76DC419
    crc32Table(5) = &H706AF48F
    crc32Table(6) = &HE963A535
    crc32Table(7) = &H9E6495A3
    crc32Table(8) = &HEDB8832
    crc32Table(9) = &H79DCB8A4
    crc32Table(10) = &HE0D5E91E
    crc32Table(11) = &H97D2D988
    crc32Table(12) = &H9B64C2B
    crc32Table(13) = &H7EB17CBD
    crc32Table(14) = &HE7B82D07
    crc32Table(15) = &H90BF1D91
    crc32Table(16) = &H1DB71064
    crc32Table(17) = &H6AB020F2
    crc32Table(18) = &HF3B97148
    crc32Table(19) = &H84BE41DE
    crc32Table(20) = &H1ADAD47D
    crc32Table(21) = &H6DDDE4EB
    crc32Table(22) = &HF4D4B551
    crc32Table(23) = &H83D385C7
    crc32Table(24) = &H136C9856
    crc32Table(25) = &H646BA8C0
    crc32Table(26) = &HFD62F97A
    crc32Table(27) = &H8A65C9EC
    crc32Table(28) = &H14015C4F
    crc32Table(29) = &H63066CD9
    crc32Table(30) = &HFA0F3D63
    crc32Table(31) = &H8D080DF5
    crc32Table(32) = &H3B6E20C8
    crc32Table(33) = &H4C69105E
    crc32Table(34) = &HD56041E4
    crc32Table(35) = &HA2677172
    crc32Table(36) = &H3C03E4D1
    crc32Table(37) = &H4B04D447
    crc32Table(38) = &HD20D85FD
    crc32Table(39) = &HA50AB56B
    crc32Table(40) = &H35B5A8FA
    crc32Table(41) = &H42B2986C
    crc32Table(42) = &HDBBBC9D6
    crc32Table(43) = &HACBCF940
    crc32Table(44) = &H32D86CE3
    crc32Table(45) = &H45DF5C75
    crc32Table(46) = &HDCD60DCF
    crc32Table(47) = &HABD13D59
    crc32Table(48) = &H26D930AC
    crc32Table(49) = &H51DE003A
    crc32Table(50) = &HC8D75180
    crc32Table(51) = &HBFD06116
    crc32Table(52) = &H21B4F4B5
    crc32Table(53) = &H56B3C423
    crc32Table(54) = &HCFBA9599
    crc32Table(55) = &HB8BDA50F
    crc32Table(56) = &H2802B89E
    crc32Table(57) = &H5F058808
    crc32Table(58) = &HC60CD9B2
    crc32Table(59) = &HB10BE924
    crc32Table(60) = &H2F6F7C87
    crc32Table(61) = &H58684C11
    crc32Table(62) = &HC1611DAB
    crc32Table(63) = &HB6662D3D
    crc32Table(64) = &H76DC4190
    crc32Table(65) = &H1DB7106
    crc32Table(66) = &H98D220BC
    crc32Table(67) = &HEFD5102A
    crc32Table(68) = &H71B18589
    crc32Table(69) = &H6B6B51F
    crc32Table(70) = &H9FBFE4A5
    crc32Table(71) = &HE8B8D433
    crc32Table(72) = &H7807C9A2
    crc32Table(73) = &HF00F934
    crc32Table(74) = &H9609A88E
    crc32Table(75) = &HE10E9818
    crc32Table(76) = &H7F6A0DBB
    crc32Table(77) = &H86D3D2D
    crc32Table(78) = &H91646C97
    crc32Table(79) = &HE6635C01
    crc32Table(80) = &H6B6B51F4
    crc32Table(81) = &H1C6C6162
    crc32Table(82) = &H856530D8
    crc32Table(83) = &HF262004E
    crc32Table(84) = &H6C0695ED
    crc32Table(85) = &H1B01A57B
    crc32Table(86) = &H8208F4C1
    crc32Table(87) = &HF50FC457
    crc32Table(88) = &H65B0D9C6
    crc32Table(89) = &H12B7E950
    crc32Table(90) = &H8BBEB8EA
    crc32Table(91) = &HFCB9887C
    crc32Table(92) = &H62DD1DDF
    crc32Table(93) = &H15DA2D49
    crc32Table(94) = &H8CD37CF3
    crc32Table(95) = &HFBD44C65
    crc32Table(96) = &H4DB26158
    crc32Table(97) = &H3AB551CE
    crc32Table(98) = &HA3BC0074
    crc32Table(99) = &HD4BB30E2
    crc32Table(100) = &H4ADFA541
    crc32Table(101) = &H3DD895D7
    crc32Table(102) = &HA4D1C46D
    crc32Table(103) = &HD3D6F4FB
    crc32Table(104) = &H4369E96A
    crc32Table(105) = &H346ED9FC
    crc32Table(106) = &HAD678846
    crc32Table(107) = &HDA60B8D0
    crc32Table(108) = &H44042D73
    crc32Table(109) = &H33031DE5
    crc32Table(110) = &HAA0A4C5F
    crc32Table(111) = &HDD0D7CC9
    crc32Table(112) = &H5005713C
    crc32Table(113) = &H270241AA
    crc32Table(114) = &HBE0B1010
    crc32Table(115) = &HC90C2086
    crc32Table(116) = &H5768B525
    crc32Table(117) = &H206F85B3
    crc32Table(118) = &HB966D409
    crc32Table(119) = &HCE61E49F
    crc32Table(120) = &H5EDEF90E
    crc32Table(121) = &H29D9C998
    crc32Table(122) = &HB0D09822
    crc32Table(123) = &HC7D7A8B4
    crc32Table(124) = &H59B33D17
    crc32Table(125) = &H2EB40D81
    crc32Table(126) = &HB7BD5C3B
    crc32Table(127) = &HC0BA6CAD
    crc32Table(128) = &HEDB88320
    crc32Table(129) = &H9ABFB3B6
    crc32Table(130) = &H3B6E20C
    crc32Table(131) = &H74B1D29A
    crc32Table(132) = &HEAD54739
    crc32Table(133) = &H9DD277AF
    crc32Table(134) = &H4DB2615
    crc32Table(135) = &H73DC1683
    crc32Table(136) = &HE3630B12
    crc32Table(137) = &H94643B84
    crc32Table(138) = &HD6D6A3E
    crc32Table(139) = &H7A6A5AA8
    crc32Table(140) = &HE40ECF0B
    crc32Table(141) = &H9309FF9D
    crc32Table(142) = &HA00AE27
    crc32Table(143) = &H7D079EB1
    crc32Table(144) = &HF00F9344
    crc32Table(145) = &H8708A3D2
    crc32Table(146) = &H1E01F268
    crc32Table(147) = &H6906C2FE
    crc32Table(148) = &HF762575D
    crc32Table(149) = &H806567CB
    crc32Table(150) = &H196C3671
    crc32Table(151) = &H6E6B06E7
    crc32Table(152) = &HFED41B76
    crc32Table(153) = &H89D32BE0
    crc32Table(154) = &H10DA7A5A
    crc32Table(155) = &H67DD4ACC
    crc32Table(156) = &HF9B9DF6F
    crc32Table(157) = &H8EBEEFF9
    crc32Table(158) = &H17B7BE43
    crc32Table(159) = &H60B08ED5
    crc32Table(160) = &HD6D6A3E8
    crc32Table(161) = &HA1D1937E
    crc32Table(162) = &H38D8C2C4
    crc32Table(163) = &H4FDFF252
    crc32Table(164) = &HD1BB67F1
    crc32Table(165) = &HA6BC5767
    crc32Table(166) = &H3FB506DD
    crc32Table(167) = &H48B2364B
    crc32Table(168) = &HD80D2BDA
    crc32Table(169) = &HAF0A1B4C
    crc32Table(170) = &H36034AF6
    crc32Table(171) = &H41047A60
    crc32Table(172) = &HDF60EFC3
    crc32Table(173) = &HA867DF55
    crc32Table(174) = &H316E8EEF
    crc32Table(175) = &H4669BE79
    crc32Table(176) = &HCB61B38C
    crc32Table(177) = &HBC66831A
    crc32Table(178) = &H256FD2A0
    crc32Table(179) = &H5268E236
    crc32Table(180) = &HCC0C7795
    crc32Table(181) = &HBB0B4703
    crc32Table(182) = &H220216B9
    crc32Table(183) = &H5505262F
    crc32Table(184) = &HC5BA3BBE
    crc32Table(185) = &HB2BD0B28
    crc32Table(186) = &H2BB45A92
    crc32Table(187) = &H5CB36A04
    crc32Table(188) = &HC2D7FFA7
    crc32Table(189) = &HB5D0CF31
    crc32Table(190) = &H2CD99E8B
    crc32Table(191) = &H5BDEAE1D
    crc32Table(192) = &H9B64C2B0
    crc32Table(193) = &HEC63F226
    crc32Table(194) = &H756AA39C
    crc32Table(195) = &H26D930A
    crc32Table(196) = &H9C0906A9
    crc32Table(197) = &HEB0E363F
    crc32Table(198) = &H72076785
    crc32Table(199) = &H5005713
    crc32Table(200) = &H95BF4A82
    crc32Table(201) = &HE2B87A14
    crc32Table(202) = &H7BB12BAE
    crc32Table(203) = &HCB61B38
    crc32Table(204) = &H92D28E9B
    crc32Table(205) = &HE5D5BE0D
    crc32Table(206) = &H7CDCEFB7
    crc32Table(207) = &HBDBDF21
    crc32Table(208) = &H86D3D2D4
    crc32Table(209) = &HF1D4E242
    crc32Table(210) = &H68DDB3F8
    crc32Table(211) = &H1FDA836E
    crc32Table(212) = &H81BE16CD
    crc32Table(213) = &HF6B9265B
    crc32Table(214) = &H6FB077E1
    crc32Table(215) = &H18B74777
    crc32Table(216) = &H88085AE6
    crc32Table(217) = &HFF0F6A70
    crc32Table(218) = &H66063BCA
    crc32Table(219) = &H11010B5C
    crc32Table(220) = &H8F659EFF
    crc32Table(221) = &HF862AE69
    crc32Table(222) = &H616BFFD3
    crc32Table(223) = &H166CCF45
    crc32Table(224) = &HA00AE278
    crc32Table(225) = &HD70DD2EE
    crc32Table(226) = &H4E048354
    crc32Table(227) = &H3903B3C2
    crc32Table(228) = &HA7672661
    crc32Table(229) = &HD06016F7
    crc32Table(230) = &H4969474D
    crc32Table(231) = &H3E6E77DB
    crc32Table(232) = &HAED16A4A
    crc32Table(233) = &HD9D65ADC
    crc32Table(234) = &H40DF0B66
    crc32Table(235) = &H37D83BF0
    crc32Table(236) = &HA9BCAE53
    crc32Table(237) = &HDEBB9EC5
    crc32Table(238) = &H47B2CF7F
    crc32Table(239) = &H30B5FFE9
    crc32Table(240) = &HBDBDF21C
    crc32Table(241) = &HCABAC28A
    crc32Table(242) = &H53B39330
    crc32Table(243) = &H24B4A3A6
    crc32Table(244) = &HBAD03605
    crc32Table(245) = &HCDD70693
    crc32Table(246) = &H54DE5729
    crc32Table(247) = &H23D967BF
    crc32Table(248) = &HB3667A2E
    crc32Table(249) = &HC4614AB8
    crc32Table(250) = &H5D681B02
    crc32Table(251) = &H2A6F2B94
    crc32Table(252) = &HB40BBE37
    crc32Table(253) = &HC30C8EA1
    crc32Table(254) = &H5A05DF1B
    crc32Table(255) = &H2D02EF8D
    
End Sub

Function GetCrc32(ByRef bArray() As Byte, offset As Integer, fileSize As Integer) As Long
    Dim crc32Result As Long
    crc32Result = &HFFFFFFFF

    Dim i As Integer
    Dim iLookup As Integer

    For i = 0 To fileSize - 1
        iLookup = (crc32Result And &HFF) Xor bArray(offset + i)
        crc32Result = ((crc32Result And &HFFFFFF00) \ &H100) And &HFFFFFF
        crc32Result = crc32Result Xor crc32Table(iLookup)
    Next

    GetCrc32 = Not (crc32Result)
End Function

Sub Main()
    WriteStdOut ("Hello World!" & vbCrLf)
    Dim fileNum As Integer
    Dim bytes() As Byte
    Dim fileSize As Integer
    Dim index As Integer
    Dim dataSize As Long
    Dim dataCRC As Long
    Dim getCRC As Long
        
    fileNum = FreeFile
    Open "data.bin" For Binary As fileNum
    ReDim bytes(LOF(fileNum) - 1)
    Get fileNum, , bytes
    fileSize = LOF(fileNum)
    Close fileNum
    
    WriteStdOut ("File size=" & fileSize & vbCrLf)

    'For index = 0 To fileSize Step 1
    '    WriteStdOut ("byts[" & index & "]=0x" & Right$("00" & Hex$(bytes(index)), 2) & vbCrLf)
    'Next
    
    dataSize = Byte4ToLong(bytes, 4)
    WriteStdOut ("Data size=" & dataSize & vbCrLf)
    
    If dataSize + 4 + 4 + 4 > fileSize Then
        WriteStdOut ("Data fileSize=" & dataSize & vbCrLf)
        End
    End If
    
    dataCRC = Byte4ToLong(bytes, dataSize + 4 + 4)
    WriteStdOut ("Data CRC=0x" & Right$("00000000" & Hex$(dataCRC), 8) & "," & dataCRC & vbCrLf)

    InitCrc32
    getCRC = GetCrc32(bytes, 0, dataSize + 4 + 4)
    WriteStdOut ("Get CRC=0x" & Right$("00000000" & Hex$(getCRC), 8) & "," & getCRC & vbCrLf)
   
End Sub

