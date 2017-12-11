Option Compare Database
Option Explicit

Private mryOwnerDataTable() As Variant
Private arycboOwnerData() As Variant

Private mryVeterinarianDataTable() As Variant
Private arycboVeterinarianData() As Variant

Function IsInArray(stringToBeFound As String, arr As Variant) As Long
  Dim i As Long
  ' default return value if value not found in array
  IsInArray = -1

  For i = LBound(arr) To UBound(arr)
    If StrComp(stringToBeFound, arr(i), vbTextCompare) = 0 Then
      IsInArray = i
      Exit For
    End If
  Next i
End Function

Private Sub Form_Close()
    'Set target queries/tables
    Dim qryOwnersTable As DAO.Recordset
    
    'Set counters to track progress
    Dim intOwnerCounter As Integer
    Dim intVeterinarianCounter As Integer
    
    'Set Variants cbo populating
    Dim varOwnersElement As Variant
    Dim varVetElement As Variant
    
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'Open query/table
    Set qryOwnersTable = CurrentDb.OpenRecordset("SELECT * FROM qryOwnersTable ORDER BY txtSurname", dbOpenDynaset)
    
    'Set to last record to make sure maximum size is correct
    qryOwnersTable.MoveLast
    Debug.Print "qryOwnersTable.RecordCount 1: " & qryOwnersTable.RecordCount
    
    If Not qryOwnersTable.EOF Then
        'Ensure we begin on the first row
        qryOwnersTable.MoveFirst
    
        'Start OwnerCounter
        intOwnerCounter = 0
        'The size of the array should be equal to the number of rows in the table
        'Need to size the array
        'mryOwnerDataTable is a multidimensional array containing Owner/Vet unique ID and Full Name
        'arycboOwnerData is a normal array containing Full Names to populate the Owner dropdown
        'Had trouble only displaying 1 column of a mry in a cbo, so the ary is used to fill ListIndex is used on mry to retrieve data
        ReDim mryOwnerDataTable(2, qryOwnersTable.RecordCount) As Variant
        ReDim arycboOwnerData(qryOwnersTable.RecordCount) As Variant
        
        Debug.Print "qryOwnersTable.RecordCount 2: " & qryOwnersTable.RecordCount
        
        'Set first cell to be blank for dropdown
        mryOwnerDataTable(0, 0) = ("")
        mryOwnerDataTable(1, 0) = ("")
        arycboOwnerData(0) = ("")
        Debug.Print "mryOwnerDataTable ("; 0 & "): " & mryOwnerDataTable(0, 0) & ", " & mryOwnerDataTable(1, 0)
        
        'Move to next row
        intOwnerCounter = 1
        
        Do Until qryOwnersTable.EOF
            'Get query data and populate the relevant rows with it, repeat until end of query
            'Data is sorted by order of surname
            mryOwnerDataTable(0, intOwnerCounter) = qryOwnersTable.Fields("txtID")
            mryOwnerDataTable(1, intOwnerCounter) = qryOwnersTable.Fields("txtFullname")
            arycboOwnerData(intOwnerCounter) = qryOwnersTable.Fields("txtFullname")
            Debug.Print "mryOwnerDataTable ("; intOwnerCounter & "): " & mryOwnerDataTable(0, intOwnerCounter) & ", " & mryOwnerDataTable(1, intOwnerCounter)
            Debug.Print "arycboOwnerData ("; intOwnerCounter & "): " & arycboOwnerData(intOwnerCounter)
            
            intOwnerCounter = intOwnerCounter + 1
            qryOwnersTable.MoveNext
        Loop
    
    End If
    
    If IsObject(qryOwnersTable) Then Set qryOwnersTable = Nothing
    'Reset counter for cbo populating
    intOwnerCounter = 0
    'Debug.Print UBound(arycboOwnerData)
    'Blank existing values from cbo
    Me.cboOwner.RowSource = ""
    'Fill cbo with ary data
    For Each varOwnersElement In arycboOwnerData
        Me.cboOwner.AddItem Item:=arycboOwnerData(intOwnerCounter), Index:=intOwnerCounter
        intOwnerCounter = intOwnerCounter + 1
    Next
    
    Debug.Print (UBound(mryOwnerDataTable, 1) & "x" & UBound(mryOwnerDataTable, 2))
    Debug.Print (UBound(arycboOwnerData))
End Sub

Private Sub btnAnimal_Click()
    On Error GoTo Err_Open_frmCatPatients_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "frmCatPatients"
    
    stLinkCriteria = "[txtPatientID1]=" & Me.cboAnimals.Value
    Debug.Print stLinkCriteria
    If (Me.cboAnimals.Value > 0) Then
        DoCmd.OpenForm stDocName
        formOpen = "o"
        Forms("frmCatPatients").Form.Recordset.FindFirst "[txtPatientID1] = " & Me.cboAnimals.Value
    End If
Exit_Open_frmCatPatients_Click:
    'DoCmd.Close acForm, "frmCatSamplesQuery", acSavePrompt
    Exit Sub

Err_Open_frmCatPatients_Click:
    MsgBox Err.Description
    Resume Exit_Open_frmCatPatients_Click
End Sub

Private Sub btnEditOwnersForm_Click()
    Call frmOwnersSwitchLock
End Sub

Private Sub cboCountry_Click()
Dim countries(249) As Variant
    countries(0) = ""
    'A
    countries(1) = "AD - Andorra"
    countries(2) = "AE - United Arab Emirates"
    countries(3) = "AF - Afghanistan"
    countries(4) = "AG - Antigua and Barbuda"
    countries(5) = "AI - Anguilla"
    countries(6) = "AL - Albania"
    countries(7) = "AM - Armenia"
    countries(8) = "AO - Angola"
    countries(9) = "AQ - Antarctica"
    countries(10) = "AR - Argentina"
    countries(11) = "AS - American Samoa"
    countries(12) = "AT - Austria"
    countries(13) = "AU - Australia"
    countries(14) = "AW - Aruba"
    countries(15) = "AZ - Azerbaijan"
    'B
    countries(16) = "BA - Bosnia and Herzegovina"
    countries(17) = "BB - Barbados"
    countries(18) = "BD - Bangladesh"
    countries(19) = "BE - Belgium"
    countries(20) = "BF - Burkina Faso"
    countries(21) = "BG - Bulgaria"
    countries(22) = "BH - Bahrain"
    countries(23) = "BI - Burundi"
    countries(24) = "BJ - Benin"
    countries(25) = "BL - Saint Barthelemy"
    countries(26) = "BM - Bermuda"
    countries(27) = "BN - Brunei"
    countries(28) = "BO - Bolivia"
    countries(29) = "BR - Brazil"
    countries(30) = "BS - The Bahamas"
    countries(31) = "BT - Bhutan"
    countries(32) = "BV - Bouvet Island"
    countries(33) = "BW - Botswana"
    countries(34) = "BY - Belarus"
    countries(35) = "BZ - Belize"
    'C
    countries(36) = "CA - Canada"
    countries(37) = "CC - Cocos (Keeling) Islands"
    countries(38) = "CD - Democratic Republic of the Congo"
    countries(39) = "CF - Central African Republic"
    countries(40) = "CG - Republic of the Congo"
    countries(41) = "CH - Switzerland"
    countries(42) = "CI - Cote d'Ivoire"
    countries(43) = "CK - Cook Islands"
    countries(44) = "CL - Chile"
    countries(45) = "CM - Cameroon"
    countries(46) = "CN - China"
    countries(47) = "CO - Colombia"
    countries(48) = "CR - Costa Rica"
    countries(49) = "CU - Cuba"
    countries(50) = "CV - Cape Verde"
    countries(51) = "CW - Curacao"
    countries(52) = "CX - Christmas Island"
    countries(53) = "CY - Cyprus"
    countries(54) = "CZ - Czech Republic"
    'D
    countries(55) = "DE - Germany"
    countries(56) = "DJ - Djibouti"
    countries(57) = "DK - Denmark"
    countries(58) = "DM - Dominica"
    countries(59) = "DO - Dominican Republic"
    countries(60) = "DZ - Algeria"
    'E
    countries(61) = "EC - Ecuador"
    countries(62) = "EE - Estonia"
    countries(63) = "EG - Egypt"
    countries(64) = "EH - Western Sahara"
    countries(65) = "ER - Eritrea"
    countries(66) = "ES - Spain"
    countries(67) = "ET - Ethiopia"
    'F
    countries(68) = "FI - Finland"
    countries(69) = "FJ - Fiji"
    countries(70) = "FK - Falkland Islands (Islas Malvinas)"
    countries(71) = "FM - Federated States of Micronesia"
    countries(72) = "FO - Faroe Islands"
    countries(73) = "FR - France"
    'G
    countries(74) = "GA - Gabon"
    countries(75) = "GB - United Kingdom"
    countries(76) = "GD - Grenada"
    countries(77) = "GE - Georgia"
    countries(78) = "GF - French Guiana"
    countries(79) = "GG - Guernsey"
    countries(80) = "GH - Ghana"
    countries(81) = "GI - Gibraltar"
    countries(82) = "GL - Greenland"
    countries(83) = "GM - Gambia"
    countries(84) = "GN - Guinea"
    countries(85) = "GP - Guadeloupe"
    countries(86) = "GQ - Equatorial Guinea"
    countries(87) = "GR - Greece"
    countries(88) = "GS - South Georgia and the Islands"
    countries(89) = "GT - Guatemala"
    countries(90) = "GU - Guam"
    countries(91) = "GW - Guinea-Bissau"
    countries(92) = "GY - Guyana"
    'H
    countries(93) = "HK - Hong Kong"
    countries(94) = "HM - Heard Island and McDonald Islands"
    countries(95) = "HN - Honduras"
    countries(96) = "HR - Croatia"
    countries(97) = "HT - Haiti"
    countries(98) = "HU - Hungary"
    'I
    countries(99) = "ID - Indonesia"
    countries(100) = "IE - Ireland"
    countries(101) = "IL - Israel"
    countries(102) = "IM - Isle of Man"
    countries(103) = "IN - India"
    countries(104) = "IO - British Indian Ocean Territory"
    countries(105) = "IQ - Iraq"
    countries(106) = "IR - Iran"
    countries(107) = "IS - Iceland"
    countries(108) = "IT - Italy"
    'J
    countries(109) = "JE - Jersey"
    countries(110) = "JM - Jamaica"
    countries(111) = "JO - Jordan"
    countries(112) = "JP - Japan"
    'K
    countries(113) = "KE - Kenya"
    countries(114) = "KG - Kyrgyzstan"
    countries(115) = "KH - Cambodia"
    countries(116) = "KI - Kiribati"
    countries(117) = "KM - Comoros"
    countries(118) = "KN - Saint Kitts and Nevis"
    countries(119) = "KP - North Korea"
    countries(120) = "KR - South Korea"
    countries(121) = "KW - Kuwait"
    countries(122) = "KY - Cayman Islands"
    countries(123) = "KZ - Kazakhstan"
    'L
    countries(124) = "LA - Laos"
    countries(125) = "LB - Lebanon"
    countries(126) = "LC - Saint Lucia"
    countries(127) = "LI - Liechtenstein"
    countries(128) = "LK - Sri Lanka"
    countries(129) = "LR - Liberia"
    countries(130) = "LS - Lesotho"
    countries(131) = "LT - Lithuania"
    countries(132) = "LU - Luxembourg"
    countries(133) = "LV - Latvia"
    countries(134) = "LY - Libya"
    'M
    countries(135) = "MA - Morocco"
    countries(136) = "MC - Monaco"
    countries(137) = "MD - Moldova"
    countries(138) = "ME - Montenegro"
    countries(139) = "MF - Saint Martin"
    countries(140) = "MG - Madagascar"
    countries(141) = "MH - Marshall Islands"
    countries(142) = "MK - Macedonia"
    countries(143) = "ML - Mali"
    countries(144) = "MM - Burma"
    countries(145) = "MN - Mongolia"
    countries(146) = "MO - Macau"
    countries(147) = "MP - Northern Mariana Islands"
    countries(148) = "MQ - Martinique"
    countries(149) = "MR - Mauritania"
    countries(150) = "MS - Montserrat"
    countries(151) = "MT - Malta"
    countries(152) = "MU - Mauritius"
    countries(153) = "MV - Maldives"
    countries(154) = "MW - Malawi"
    countries(155) = "MX - Mexico"
    countries(156) = "MY - Malaysia"
    countries(157) = "MZ - Mozambique"
    'N
    countries(158) = "NA - Namibia"
    countries(159) = "NC - New Caledonia"
    countries(160) = "NE - Niger"
    countries(161) = "NF - Norfolk Island"
    countries(162) = "NG - Nigeria"
    countries(163) = "NI - Nicaragua"
    countries(164) = "NL - Netherlands"
    countries(165) = "NO - Norway"
    countries(166) = "NP - Nepal"
    countries(167) = "NR - Nauru"
    countries(168) = "NU - Niue"
    countries(169) = "NZ - New Zealand"
    'O
    countries(170) = "OM - Oman"
    'P
    countries(171) = "PA - Panama"
    countries(172) = "PE - Peru"
    countries(173) = "PF - French Polynesia"
    countries(174) = "PG - Papua New Guinea"
    countries(175) = "PH - Philippines"
    countries(176) = "PK - Pakistan"
    countries(177) = "PL - Poland"
    countries(178) = "PM - Saint Pierre and Miquelon"
    countries(179) = "PN - Pitcairn Islands"
    countries(180) = "PR - Puerto Rico"
    countries(181) = "PS - Gaza Strip"
    countries(182) = "PS - West Bank"
    countries(183) = "PT - Portugal"
    countries(184) = "PW - Palau"
    countries(185) = "PY - Paraguay"
    'Q
    countries(186) = "QA - Qatar"
    'R
    countries(187) = "RE - Reunion"
    countries(188) = "RO - Romania"
    countries(189) = "RS - Serbia"
    countries(190) = "RU - Russia"
    countries(191) = "RW - Rwanda"
    'S
    countries(192) = "SA - Saudi Arabia"
    countries(193) = "SB - Solomon Islands"
    countries(194) = "SC - Seychelles"
    countries(195) = "SD - Sudan"
    countries(196) = "SE - Sweden"
    countries(197) = "SG - Singapore"
    countries(198) = "SH - Saint Helena, Ascension, and Tristan da Cunha"
    countries(199) = "SI - Slovenia"
    countries(200) = "SJ - Svalbard"
    countries(201) = "SK - Slovakia"
    countries(202) = "SL - Sierra Leone"
    countries(203) = "SM - San Marino"
    countries(204) = "SN - Senegal"
    countries(205) = "SO - Somalia"
    countries(206) = "SR - Suriname"
    countries(207) = "SS - South Sudan"
    countries(208) = "ST - Sao Tome and Principe"
    countries(209) = "SV - El Salvador"
    countries(210) = "SX - Sint Maarten"
    countries(211) = "SY - Syria"
    countries(212) = "SZ - Swaziland"
    'T
    countries(213) = "TC - Turks and Caicos Islands"
    countries(214) = "TD - Chad"
    countries(215) = "TF - French Southern and Antarctic Lands"
    countries(216) = "TG - Togo"
    countries(217) = "TH - Thailand"
    countries(218) = "TJ - Tajikistan"
    countries(219) = "TK - Tokelau"
    countries(220) = "TL - Timor-Leste"
    countries(221) = "TM - Turkmenistan"
    countries(222) = "TN - Tunisia"
    countries(223) = "TO - Tonga"
    countries(224) = "TR - Turkey"
    countries(225) = "TT - Trinidad and Tobago"
    countries(226) = "TV - Tuvalu"
    countries(227) = "TW - Taiwan"
    countries(228) = "TZ - Tanzania"
    'U
    countries(229) = "UA - Ukraine"
    countries(230) = "UG - Uganda"
    countries(231) = "UM - United States Minor Outlying Islands"
    countries(232) = "US - United States"
    countries(233) = "UY - Uruguay"
    countries(234) = "UZ - Uzbekistan"
    'V
    countries(235) = "VA - Vatican City State"
    countries(236) = "VC - Saint Vincent and the Grenadines"
    countries(237) = "VE - Venezuela"
    countries(238) = "VG - British Virgin Islands"
    countries(239) = "VI - Virgin Islands"
    countries(240) = "VN - Vietnam"
    countries(241) = "VU - Vanuatu"
    'W
    countries(242) = "WF - Wallis and Futuna"
    countries(243) = "WS - Samoa"
    'X
    countries(244) = "XK - Kosovo"
    'Y
    countries(245) = "YE - Yemen"
    countries(246) = "YT - Mayotte"
    'Z
    countries(247) = "ZA - South Africa"
    countries(248) = "ZM - Zambia"
    countries(249) = "ZW - Zimbabwe"
    
Dim countrycodes(249) As Variant
    countrycodes(1) = "+376"
    countrycodes(2) = "+971"
    countrycodes(3) = "+93"
    countrycodes(4) = "+1 268"
    countrycodes(5) = "+1 264"
    countrycodes(6) = "+355"
    countrycodes(7) = "+374"
    countrycodes(8) = "+244"
    countrycodes(9) = "+672"
    countrycodes(10) = "+54"
    countrycodes(11) = "+1 684"
    countrycodes(12) = "+43"
    countrycodes(13) = "+61"
    countrycodes(14) = "+297"
    countrycodes(15) = "+994"
    'B
    countrycodes(16) = "+387"
    countrycodes(17) = "+1 246"
    countrycodes(18) = "+880"
    countrycodes(19) = "+32"
    countrycodes(20) = "+226"
    countrycodes(21) = "+359"
    countrycodes(22) = "+973"
    countrycodes(23) = "+257"
    countrycodes(24) = "+229"
    countrycodes(25) = "+590"
    countrycodes(26) = "+1 441"
    countrycodes(27) = "+673"
    countrycodes(28) = "+591"
    countrycodes(29) = "+55"
    countrycodes(30) = "+1 242"
    countrycodes(31) = "+957"
    countrycodes(32) = "+47"
    countrycodes(33) = "+267"
    countrycodes(34) = "+375"
    countrycodes(35) = "+501"
    'C
    countrycodes(36) = "+1"
    countrycodes(37) = "+61"
    countrycodes(38) = "+243"
    countrycodes(39) = "+236"
    countrycodes(40) = "+242"
    countrycodes(41) = "+41"
    countrycodes(42) = "+225"
    countrycodes(43) = "+682"
    countrycodes(44) = "+56"
    countrycodes(45) = "+237"
    countrycodes(46) = "+86"
    countrycodes(47) = "+57"
    countrycodes(48) = "+506"
    countrycodes(49) = "+53"
    countrycodes(50) = "+238"
    countrycodes(51) = "+599 9"
    countrycodes(52) = "+61"
    countrycodes(53) = "+357"
    countrycodes(54) = "+420"
    'D
    countrycodes(55) = "+49"
    countrycodes(56) = "+253"
    countrycodes(57) = "+45"
    countrycodes(58) = "+1 767"
    countrycodes(59) = "+1"
    countrycodes(60) = "+213"
    'E
    countrycodes(61) = "+593"
    countrycodes(62) = "+372"
    countrycodes(63) = "+20"
    countrycodes(64) = "+212"
    countrycodes(65) = "+291"
    countrycodes(66) = "+34"
    countrycodes(67) = "+251"
    'F
    countrycodes(68) = "+358"
    countrycodes(69) = "+679"
    countrycodes(70) = "+500"
    countrycodes(71) = "+691"
    countrycodes(72) = "+298"
    countrycodes(73) = "+33"
    'G
    countrycodes(74) = "+241"
    countrycodes(75) = "+44"
    countrycodes(76) = "+1 473"
    countrycodes(77) = "+995"
    countrycodes(78) = "+594"
    countrycodes(79) = "+44"
    countrycodes(80) = "+233"
    countrycodes(81) = "+350"
    countrycodes(82) = "+299"
    countrycodes(83) = "+220"
    countrycodes(84) = "+224"
    countrycodes(85) = "+590"
    countrycodes(86) = "+224"
    countrycodes(87) = "+30"
    countrycodes(88) = "+500"
    countrycodes(89) = "+502"
    countrycodes(90) = "+1 671"
    countrycodes(91) = "+245"
    countrycodes(92) = "+592"
    'H
    countrycodes(93) = "+852"
    countrycodes(94) = "+672"
    countrycodes(95) = "+504"
    countrycodes(96) = "+385"
    countrycodes(97) = "+509"
    countrycodes(98) = "+36"
    'I
    countrycodes(99) = "+62"
    countrycodes(100) = "+353"
    countrycodes(101) = "+972"
    countrycodes(102) = "+44"
    countrycodes(103) = "+91"
    countrycodes(104) = "+246"
    countrycodes(105) = "+964"
    countrycodes(106) = "+98"
    countrycodes(107) = "+354"
    countrycodes(108) = "+39"
    'J
    countrycodes(109) = "+44"
    countrycodes(110) = "+1 876"
    countrycodes(111) = "+962"
    countrycodes(112) = "+81"
    'K
    countrycodes(113) = "+254"
    countrycodes(114) = "+996"
    countrycodes(115) = "+855"
    countrycodes(116) = "+686"
    countrycodes(117) = "+269"
    countrycodes(118) = "+1 869"
    countrycodes(119) = "+850"
    countrycodes(120) = "+82"
    countrycodes(121) = "+965"
    countrycodes(122) = "+1 345"
    countrycodes(123) = "+7"
    'L
    countrycodes(124) = "+856"
    countrycodes(125) = "+961"
    countrycodes(126) = "+1 758"
    countrycodes(127) = "+423"
    countrycodes(128) = "+94"
    countrycodes(129) = "+231"
    countrycodes(130) = "+266"
    countrycodes(131) = "+370"
    countrycodes(132) = "+352"
    countrycodes(133) = "+371"
    countrycodes(134) = "+218"
    'M
    countrycodes(135) = "+212"
    countrycodes(136) = "+377"
    countrycodes(137) = "+373"
    countrycodes(138) = "+382"
    countrycodes(139) = "+590"
    countrycodes(140) = "+261"
    countrycodes(141) = "+692"
    countrycodes(142) = "+389"
    countrycodes(143) = "+223"
    countrycodes(144) = "+95"
    countrycodes(145) = "+976"
    countrycodes(146) = "+853"
    countrycodes(147) = "+1 670"
    countrycodes(148) = "+596"
    countrycodes(149) = "+222"
    countrycodes(150) = "+1 664"
    countrycodes(151) = "+356"
    countrycodes(152) = "+230"
    countrycodes(153) = "+960"
    countrycodes(154) = "+265"
    countrycodes(155) = "+52"
    countrycodes(156) = "+60"
    countrycodes(157) = "+258"
    'N
    countrycodes(158) = "+264"
    countrycodes(159) = "+687"
    countrycodes(160) = "+227"
    countrycodes(161) = "+672"
    countrycodes(162) = "+234"
    countrycodes(163) = "+505"
    countrycodes(164) = "+31"
    countrycodes(165) = "+47"
    countrycodes(166) = "+977"
    countrycodes(167) = "+674"
    countrycodes(168) = "+683"
    countrycodes(169) = "+64"
    'O
    countrycodes(170) = "+968"
    'P
    countrycodes(171) = "+507"
    countrycodes(172) = "+51"
    countrycodes(173) = "+689"
    countrycodes(174) = "+675"
    countrycodes(175) = "+63"
    countrycodes(176) = "+92"
    countrycodes(177) = "+48"
    countrycodes(178) = "+508"
    countrycodes(179) = "+64"
    countrycodes(180) = "+1"
    countrycodes(181) = "+970"
    countrycodes(182) = "+970"
    countrycodes(183) = "+351"
    countrycodes(184) = "+680"
    countrycodes(185) = "+595"
    'Q
    countrycodes(186) = "+974"
    'R
    countrycodes(187) = "+262"
    countrycodes(188) = "+40"
    countrycodes(189) = "+381"
    countrycodes(190) = "+7"
    countrycodes(191) = "+250"
    'S
    countrycodes(192) = "+966"
    countrycodes(193) = "+677"
    countrycodes(194) = "+248"
    countrycodes(195) = "+249"
    countrycodes(196) = "+46"
    countrycodes(197) = "+65"
    countrycodes(198) = "+290"
    countrycodes(199) = "+386"
    countrycodes(200) = "+47"
    countrycodes(201) = "+421"
    countrycodes(202) = "+232"
    countrycodes(203) = "+378"
    countrycodes(204) = "+221"
    countrycodes(205) = "+252"
    countrycodes(206) = "+597"
    countrycodes(207) = "+211"
    countrycodes(208) = "+239"
    countrycodes(209) = "+503"
    countrycodes(210) = "+1 721"
    countrycodes(211) = "+963"
    countrycodes(212) = "+268"
    'T
    countrycodes(213) = "+1 649"
    countrycodes(214) = "+235"
    countrycodes(215) = "+262"
    countrycodes(216) = "+228"
    countrycodes(217) = "+66"
    countrycodes(218) = "+992"
    countrycodes(219) = "+690"
    countrycodes(220) = "+670"
    countrycodes(221) = "+993"
    countrycodes(222) = "+216"
    countrycodes(223) = "+676"
    countrycodes(224) = "+90"
    countrycodes(225) = "+1 868"
    countrycodes(226) = "+688"
    countrycodes(227) = "+886"
    countrycodes(228) = "+255"
    'U
    countrycodes(229) = "+380"
    countrycodes(230) = "+256"
    countrycodes(231) = "+699"
    countrycodes(233) = "+598"
    countrycodes(232) = "+1"
    countrycodes(234) = "+998"
    'V
    countrycodes(235) = "+39 06 698"
    countrycodes(236) = "+1 784"
    countrycodes(237) = "+58"
    countrycodes(238) = "+1 284"
    countrycodes(239) = "+1 340"
    countrycodes(240) = "+84"
    countrycodes(241) = "+678"
    'W
    countrycodes(242) = "+681"
    countrycodes(243) = "+685"
    'X
    countrycodes(244) = "+383"
    'Y
    countrycodes(245) = "+383"
    countrycodes(246) = "+262"
    'Z
    countrycodes(247) = "+27"
    countrycodes(248) = "+260"
    countrycodes(249) = "+263"
    
    Dim strcountry As String
    Dim strcountrycode As String
    Dim pos
    strcountry = Me.cboCountry.Column(0)
    pos = IsInArray(strcountry, countries)
    strcountrycode = countrycodes(pos)
    Me.txtPhonePrefix.Value = strcountrycode
End Sub

Private Sub btnPrintReport_Click()
    Dim strDocName As String
    Dim strWhere As String
    strDocName = "rprtCatPatients"
    strWhere = "[autoEntryID]=" & Me!autoEntryID
    DoCmd.OpenReport strDocName, acPreview, , strWhere
End Sub

Private Sub Form_Load()
    DoCmd.GoToRecord , , acNewRec
End Sub
