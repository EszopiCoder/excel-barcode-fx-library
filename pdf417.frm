VERSION 5.00
Begin VB.Form pdf417 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Code barre PDF 417 / PDF 417 barcode"
   ClientHeight    =   7935
   ClientLeft      =   300
   ClientTop       =   420
   ClientWidth     =   10440
   Icon            =   "pdf417.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   10440
   Begin VB.CommandButton Command3 
      Caption         =   "Réinitialiser / Init."
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   7320
      Width           =   1600
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8040
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "0"
      Top             =   1980
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "-1"
      Top             =   1980
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   600
      Width           =   10095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fermer / Close"
      Height          =   375
      Left            =   8640
      TabIndex        =   13
      Top             =   120
      Width           =   1600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Copier / Copy"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   7320
      Width           =   1600
   End
   Begin VB.TextBox label1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Code PDF417"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   4440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   2760
      Width           =   5895
   End
   Begin VB.TextBox label5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Grandzebu (Français)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8400
      MouseIcon       =   "pdf417.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "Supplied under GNU GPL license by :"
      Height          =   195
      Left            =   5400
      TabIndex        =   23
      Top             =   7620
      Width           =   2895
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Grandzebu (English)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8400
      MouseIcon       =   "pdf417.frx":0884
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   7620
      Width           =   1935
   End
   Begin VB.Label Label15 
      Caption         =   "True :"
      Height          =   255
      Left            =   8760
      TabIndex        =   21
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "True :"
      Height          =   255
      Left            =   3240
      TabIndex        =   20
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "Number of columns (<1 = auto) :"
      Height          =   255
      Left            =   5400
      TabIndex        =   19
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label12 
      Caption         =   "Security level (-1 = auto) :"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label COnbcol 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   9360
      TabIndex        =   17
      Top             =   1980
      Width           =   375
   End
   Begin VB.Label COsécu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3840
      TabIndex        =   16
      Top             =   1980
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "Réel :"
      Height          =   255
      Left            =   8760
      TabIndex        =   15
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "Réel :"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Nombre de colonnes (<1 = auto) :"
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "Niveau de sécurité (-1 = auto) :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Tapez votre texte ici / Type your text here :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Voici le résultat / Here is the result :"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   2450
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Voici la chaine de code / Here is the code string :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2450
      Width           =   3855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Fourni sous license GNU GPL par :"
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Top             =   7320
      Width           =   2655
   End
End
Attribute VB_Name = "pdf417"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2004 (Grandzebu)
'Ce programme et la police de caractères qui l'accompagne est libre, vous pouvez le redistribuer et/ou le
'modifier selon les termes de la Licence Publique Générale GNU publiée par la Free Software Foundation
'(version 2 ou bien toute autre version ultérieure choisie par vous).
'Les fonctions d'encodage des codes barres sont régies par la Licence Générale Publique Amoindrie GNU (GNU LGPL)
'Ce programme est distribué car potentiellement utile, mais SANS AUCUNE GARANTIE, ni explicite ni implicite,
'y compris les garanties de commercialisation ou d'adaptation dans un but spécifique. Reportez-vous à la Licence
'Publique Générale GNU pour plus de détails.
'
'This program and the font which is supplied with it is free, you can redistribute it and/or
'modify it under the terms of the GNU General Public License as published by the Free Software Foundation
'either version 2 of the License, or (at your option) any later version.
'The barcode encoding functions are governed by the GNU Lesser General Public License (GNU LGPL)
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
'even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General
'Public License for more details.

'V. 2.5.0

Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command3_Click()
  Text1.Text = ""
  Text2.Text = "-1"
  Text3.Text = "0"
  label5.Text = ""
  Text1.SetFocus
End Sub

Private Sub Label16_Click()
  ShellExecute Me.hWnd, "open", "http://grandzebu.net/informatique/codbar-en/codbar.htm", vbNullString, vbNullString, 3
End Sub

Private Sub Text2_LostFocus()
  If Val(Text2) < 0 Then Text2 = "-1"
  If Val(Text2) > 8 Then Text2 = 8
  Call Text1_Change
End Sub

Private Sub Text3_LostFocus()
  If Val(Text3) < 1 Then Text3 = 0
  If Val(Text3) > 30 Then Text3 = 30
  Call Text1_Change
End Sub
  
Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Command2_Click()
  Clipboard.Clear
  Clipboard.SetText label5.Text
End Sub

Private Sub label1_Click()
  Text1.SetFocus
End Sub

Private Sub Label6_Click()
  ShellExecute Me.hWnd, "open", "http://grandzebu.net", vbNullString, vbNullString, 3
End Sub

Private Sub Text1_Change()
  Dim CodeBarre$, sécu%, nbcol%, CodeErr%
  sécu = Val(Text2)
  nbcol = Val(Text3)
  CodeBarre$ = pdf417$(Text1, sécu%, nbcol%, CodeErr%)
  COsécu.Caption = sécu%
  COnbcol.Caption = nbcol%
  If CodeErr% > 1 And CodeBarre$ = "" Then
    label5.Text = "Erreur N° " & CodeErr% & Chr$(13) & Chr$(10)
    Select Case CodeErr%
    Case 2
      label5.Text = label5.Text & "Chaine$ contient trop de données, on dépasse le nombre de 928 MC. / Chaine$ contain too many datas, we go beyong the 928 MCs."
    Case 3
      label5.Text = label5.Text & "Nombre de MC par ligne trop faible, on dépasse 90 lignes. / Number of CWs per row too small, we go beyong 90 rows."
    End Select
  Else
    label5.Text = CodeBarre$
  End If
  label1.Text = CodeBarre$
End Sub

Private Function pdf417$(Chaine$, Optional ByRef sécu%, Optional ByRef nbcol%, Optional ByRef CodeErr%)
  'V 1.5.0
  'Paramètres : Une chaine à encoder
  '             Le niveau de correction souhaité, -1 = automatique.
  '             Le nombre de colonnes de MC de données souhaité, -1 = automatique
  '             Une variable qui pourra récupérer un numéro d'erreur
  'Retour : * une chaine qui, affichée avec la police PDF417.TTF, donne le code barre
  '         * une chaine vide si paramètre fourni incorrect
  '         * Sécu% contient le niveau de correction effectif
  '         * NbCol% contient le nombre de colonnes de MC de données effectif
  '         * Codeerr% contient 0 si pas d'erreur, sinon :
  '           0 : Pas d'erreur
  '           1  : Chaine$ est vide
  '           2  : Chaine$ contient trop de données, on dépasse le nombre de 928 MC.
  '           3  : Nombre de MC par ligne trop faible, on dépasse 90 lignes.
  '           10 : Le niveau de sécurité a été abaissé pour ne pas dépasser les 928 MC.
  
  'Parameters : The string to encode.
  '             The hoped sécurity level, -1 = automatic.
  '             The hoped number of data MC columns, -1 = automatic.
  '             A variable which will can retrieve an error number.
  'Return : * a string which, printed with the PDF417.TTF font, gives the bar code.
  '         * an empty string if the given parameters aren't good.
  '         * sécu% contain le really used sécurity level.
  '         * NbCol% contain the really used number of data CW columns.
  '         * Codeerr% is 0 if no error occured, else :
  '           0  : No error
  '           1  : Chaine$ is empty
  '           2  : Chaine$ contain too many datas, we go beyong the 928 CWs.
  '           3  : Number of CWs per row too small, we go beyong 90 rows.
  '           10 : The sécurity level has being lowers not to exceed the 928 CWs. (It's not an error, only a warning.)
  
  'Variables générales / Global variables
  Dim I%, J%, K%, IndexChaine%, Dummy$, flag As Boolean
  'Découpage en blocs / Splitting into blocks
  Dim Liste%(), IndexListe%
  'Compactage des données / Data compaction
  Dim Longueur%, ChaineMC$, Total
  'Traitement du mode "texte" / "text" mode processing
  Dim ListeT%(), IndexListeT%, CurTable%, ChaineT$, NewTable%
  'Codes de Reed Solomon / Reed Solomon codes
  Dim MCcorrection%()
  'MC de cotés gauche et droit / Left and right side CWs
  Dim C1%, C2%, C3%
  'Sous programme QuelMode / Sub routine QuelMode
  Dim Mode%, CodeASCII%
  'Sous programme Modulo / Sub routine Modulo
  Dim ChaineMod$, Diviseur&, ChaineMult$, Nombre&
  'Tables
  Dim ASCII$
  'Cette chaine décrit le code ASCII pour le mode "texte".
  'ASCII$ contient 95 champs de 4 chiffres correspondant aux car. de valeur ASCII 32 à 126. Les champs sont :
  '  2 chiffres indiquant la ou les tables où se trouvent ce car. (Tables numérotées 1, 2, 4 et 8)
  '  2 chiffres indiquant le N° du car. dans la table
  '  Ex. 0726 en début de chaine : le car. de code 32 est dans les tables 1, 2 et 4 à la ligne 26
  '
  'This string describe the ASCII code for the "text" mode.
  'ASCII$ contain 95 fields of 4 digits which correspond to char. ASCII values 32 to 126. These fields are :
  '  2 digits indicating the table(s) (1 or several) where this char. is located. (Table numbers : 1, 2, 4 and 8)
  '  2 digits indicating the char. number in the table
  '  Sample : 0726 at the beginning of the string : The Char. having code 32 is in the tables 1, 2 and 4 at row 26
  ASCII$ = "07260810082004151218042104100828082308241222042012131216121712190400040104020403040404050406040704080409121408000801042308020825080301000101010201030104010501060107010801090110011101120113011401150116011701180119012001210122012301240125080408050806042408070808020002010202020302040205020602070208020902100211021202130214021502160217021802190220022102220223022402250826082108270809"
  Dim CoefRS$(8)
  'CoefRS$ contient 8 chaines représentant les coefficients sur 3 chiffres des polynomes de calcul des codes de reed Solomon
  'CoefRS$ contain 8 strings describing the factors of the polynomial equations for the reed Solomon codes.
  CoefRS$(0) = "027917"
  CoefRS$(1) = "522568723809"
  CoefRS$(2) = "237308436284646653428379"
  CoefRS$(3) = "274562232755599524801132295116442428295042176065"
  CoefRS$(4) = "361575922525176586640321536742677742687284193517273494263147593800571320803133231390685330063410"
  CoefRS$(5) = "539422006093862771453106610287107505733877381612723476462172430609858822543376511400672762283184440035519031460594225535517352605158651201488502648733717083404097280771840629004381843623264543"
  CoefRS$(6) = "521310864547858580296379053779897444400925749415822093217208928244583620246148447631292908490704516258457907594723674292272096684432686606860569193219129186236287192775278173040379712463646776171491297763156732095270447090507048228821808898784663627378382262380602754336089614087432670616157374242726600269375898845454354130814587804034211330539297827865037517834315550086801004108539"
  CoefRS$(7) = "524894075766882857074204082586708250905786138720858194311913275190375850438733194280201280828757710814919089068569011204796605540913801700799137439418592668353859370694325240216257284549209884315070329793490274877162749812684461334376849521307291803712019358399908103511051008517225289470637731066255917269463830730433848585136538906090002290743199655903329049802580355588188462010134628320479130739071263318374601192605142673687234722384177752607640455193689707805641048060732621895544261852655309697755756060231773434421726528503118049795032144500238836394280566319009647550073914342126032681331792620060609441180791893754605383228749760213054297134054834299922191910532609829189020167029872449083402041656505579481173404251688095497555642543307159924558648055497010"
  CoefRS$(8) = "352077373504035599428207409574118498285380350492197265920155914299229643294871306088087193352781846075327520435543203666249346781621640268794534539781408390644102476499290632545037858916552041542289122272383800485098752472761107784860658741290204681407855085099062482180020297451593913142808684287536561076653899729567744390513192516258240518794395768848051610384168190826328596786303570381415641156237151429531207676710089168304402040708575162864229065861841512164477221092358785288357850836827736707094008494114521002499851543152729771095248361578323856797289051684466533820669045902452167342244173035463651051699591452578037124298332552043427119662777475850764364578911283711472420245288594394511327589777699688043408842383721521560644714559062145873663713159672729"
  CoefRS$(8) = CoefRS$(8) & "624059193417158209563564343693109608563365181772677310248353708410579870617841632860289536035777618586424833077597346269757632695751331247184045787680018066407369054492228613830922437519644905789420305441207300892827141537381662513056252341242797838837720224307631061087560310756665397808851309473795378031647915459806590731425216548249321881699535673782210815905303843922281073469791660162498308155422907817187062016425535336286437375273610296183923116667751353062366691379687842037357720742330005039923311424242749321054669316342299534105667488640672576540316486721610046656447171616464190531297321762752533175134014381433717045111020596284736138646411877669141919045780407164332899165726600325498655357752768223849647063310863251366304282738675410389244031121303263"
  Dim CodageMC$(2)
  'CodageMC$ contient les 3 jeux des 929 MC. Chaque MC est représenté dans la police PDF417.TTF par 3 lettres composant 3 fois 5 bits. Le premier bit toujours à 1
  ' et le dernier toujours à 0 se trouvent dans le caractère séparateur.
  'CodageMC$ contain the 3 sets of the 929 MCs. Each MC is described in the PDF417.TTF font by 3 char. composing 3 time 5 bits. The first bit which is always 1
  ' and the last one which is always 0 are into the separator character.
  CodageMC$(0) = "urAxfsypyunkxdwyozpDAulspBkeBApAseAkprAuvsxhypnkutwxgzfDAplsfBkfrApvsuxyfnkptwuwzflspsyfvspxyftwpwzfxyyrxufkxFwymzonAudsxEyolkucwdBAoksucidAkokgdAcovkuhwxazdnAotsugydlkoswugjdksosidvkoxwuizdtsowydswowjdxwoyzdwydwjofAuFsxCyodkuEwxCjclAocsuEickkocgckcckEcvAohsuayctkogwuajcssogicsgcsacxsoiycwwoijcwicyyoFkuCwxBjcdAoEsuCicckoEguCbcccoEaccEoEDchkoawuDjcgsoaicggoabcgacgDobjcibcFAoCsuBicEkoCguBbcEcoCacEEoCDcECcascagcaacCkuAroBaoBDcCBtfkwpwyezmnAtdswoymlktcwwojFBAmksFAkmvkthwwqzFnAmtstgyFlkmswFksFkgFvkmxwtizFtsmwyFswFsiFxwmyzFwyFyzvfAxpsyuyvdkxowyujqlAvcsxoiqkkvcgxobqkcvcamfAtFswmyqvAmdktEwwmjqtkvgwxqjhlAEkkmcgtEbhkkqsghkcEvAmhstayhvAEtkmgwtajhtkqwwvijhssEsghsgExsmiyhxsEwwmijhwwqyjhwiEyyhyyEyjhyjvFkxmwytjqdAvEsxmiqckvEgxmbqccvEaqcEqcCmFktCwwljqhkmEstCigtAEckvaitCbgskEccmEagscqgamEDEcCEhkmawtDjgxkEgsmaigwsqiimabgwgEgaEgDEiwmbjgywEiigyiEibgybgzjqFAvCsxliqEkvCgxlbqEcvCaqEEvCDqECqEBEFAmCstBighAEEkmCgtBbggkqagvDbggcEEEmCDggEqaDgg"
  CodageMC$(0) = CodageMC$(0) & "CEasmDigisEagmDbgigqbbgiaEaDgiDgjigjbqCkvBgxkrqCcvBaqCEvBDqCCqCBECkmBgtArgakECcmBagacqDamBDgaEECCgaCECBEDggbggbagbDvAqvAnqBBmAqEBEgDEgDCgDBlfAspsweyldksowClAlcssoiCkklcgCkcCkECvAlhssqyCtklgwsqjCsslgiCsgCsaCxsliyCwwlijCwiCyyCyjtpkwuwyhjndAtoswuincktogwubncctoancEtoDlFksmwwdjnhklEssmiatACcktqismbaskngglEaascCcEasEChklawsnjaxkCgstrjawsniilabawgCgaawaCiwlbjaywCiiayiCibCjjazjvpAxusyxivokxugyxbvocxuavoExuDvoCnFAtmswtirhAnEkxviwtbrgkvqgxvbrgcnEEtmDrgEvqDnEBCFAlCssliahACEklCgslbixAagknagtnbiwkrigvrblCDiwcagEnaDiwECEBCaslDiaisCaglDbiysaignbbiygrjbCaDaiDCbiajiCbbiziajbvmkxtgywrvmcxtavmExtDvmCvmBnCktlgwsrraknCcxtrracvnatlDraEnCCraCnCBraBCCklBgskraakCCclBaiikaacnDalBDiicrbaCCCiiEaaCCCBaaBCDglBrabgCDaijgabaCDDijaabDCDrijrvlcxsqvlExsnvlCvlBnBctkqrDcnBEtknrDEvlnrDCnBBrDBCBclAqaDcCBElAnibcaDEnBnibErDnCBBibCaDBibBaDqibqibnxsfvkltkfnAmnAlCAoaBoiDoCAlaBlkpkBdAkosBckkogsebBcckoaBcEkoDBhkkqwsfjBgskqiBggkqbBgaBgDBiwkrjBiiBibBjjlpAsuswhil"
  CodageMC$(0) = CodageMC$(0) & "oksuglocsualoEsuDloCBFAkmssdiDhABEksvisdbDgklqgsvbDgcBEEkmDDgElqDBEBBaskniDisBagknbDiglrbDiaBaDBbiDjiBbbDjbtukwxgyirtucwxatuEwxDtuCtuBlmkstgnqklmcstanqctvastDnqElmCnqClmBnqBBCkklgDakBCcstrbikDaclnaklDbicnraBCCbiEDaCBCBDaBBDgklrDbgBDabjgDbaBDDbjaDbDBDrDbrbjrxxcyyqxxEyynxxCxxBttcwwqvvcxxqwwnvvExxnvvCttBvvBllcssqnncllEssnrrcnnEttnrrEvvnllBrrCnnBrrBBBckkqDDcBBEkknbbcDDEllnjjcbbEnnnBBBjjErrnDDBjjCBBqDDqBBnbbqDDnjjqbbnjjnxwoyyfxwmxwltsowwfvtoxwvvtmtslvtllkossfnlolkmrnonlmlklrnmnllrnlBAokkfDBolkvbDoDBmBAljbobDmDBljbmbDljblDBvjbvxwdvsuvstnkurlurltDAubBujDujDtApAAokkegAocAoEAoCAqsAqgAqaAqDAriArbkukkucshakuEshDkuCkuBAmkkdgBqkkvgkdaBqckvaBqEkvDBqCAmBBqBAngkdrBrgkvrBraAnDBrDAnrBrrsxcsxEsxCsxBktclvcsxqsgnlvEsxnlvCktBlvBAlcBncAlEkcnDrcBnEAlCDrEBnCAlBDrCBnBAlqBnqAlnDrqBnnDrnwyowymwylswotxowyvtxmswltxlksosgfltoswvnvoltmkslnvmltlnvlAkokcfBloksvDnoBlmAklbroDnmBllbrmDnlAkvBlvDnvbrvyzeyzdwyexyuwydxytswetwuswdvxutwtvxtkselsuksdntulstrvu"
  CodageMC$(1) = "ypkzewxdAyoszeixckyogzebxccyoaxcEyoDxcCxhkyqwzfjutAxgsyqiuskxggyqbuscxgausExgDusCuxkxiwyrjptAuwsxiipskuwgxibpscuwapsEuwDpsCpxkuywxjjftApwsuyifskpwguybfscpwafsEpwDfxkpywuzjfwspyifwgpybfwafywpzjfyifybxFAymszdixEkymgzdbxEcymaxEEymDxECxEBuhAxasyniugkxagynbugcxaaugExaDugCugBoxAuisxbiowkuigxbbowcuiaowEuiDowCowBdxAoysujidwkoygujbdwcoyadwEoyDdwCdysozidygozbdyadyDdzidzbxCkylgzcrxCcylaxCEylDxCCxCBuakxDgylruacxDauaExDDuaCuaBoikubgxDroicubaoiEubDoiCoiBcykojgubrcycojacyEojDcyCcyBczgojrczaczDczrxBcykqxBEyknxBCxBBuDcxBquDExBnuDCuDBobcuDqobEuDnobCobBcjcobqcjEobncjCcjBcjqcjnxAoykfxAmxAluBoxAvuBmuBloDouBvoDmoDlcbooDvcbmcblxAexAduAuuAtoBuoBtwpAyeszFiwokyegzFbwocyeawoEyeDwoCwoBthAwqsyfitgkwqgyfbtgcwqatgEwqDtgCtgBmxAtiswrimwktigwrbmwctiamwEtiDmwCmwBFxAmystjiFwkmygtjbFwcmyaFwEmyDFwCFysmziFygmzbFyaFyDFziFzbyukzhghjsyuczhahbwyuEzhDhDyyuCyuBwmkydgzErxqkwmczhrxqcyvaydDxqEwmCxqCwmBxqBtakwngydrviktacwnavicxrawnDviEtaCviCtaBviBmiktbgwnrqykmictb"
  CodageMC$(1) = CodageMC$(1) & "aqycvjatbDqyEmiCqyCmiBqyBEykmjgtbrhykEycmjahycqzamjDhyEEyChyCEyBEzgmjrhzgEzahzaEzDhzDEzrytczgqgrwytEzgngnyytCglzytBwlcycqxncwlEycnxnEytnxnCwlBxnBtDcwlqvbctDEwlnvbExnnvbCtDBvbBmbctDqqjcmbEtDnqjEvbnqjCmbBqjBEjcmbqgzcEjEmbngzEqjngzCEjBgzBEjqgzqEjngznysozgfgfyysmgdzyslwkoycfxloysvxlmwklxlltBowkvvDotBmvDmtBlvDlmDotBvqbovDvqbmmDlqblEbomDvgjoEbmgjmEblgjlEbvgjvysegFzysdwkexkuwkdxkttAuvButAtvBtmBuqDumBtqDtEDugbuEDtgbtysFwkFxkhtAhvAxmAxqBxwekyFgzCrwecyFaweEyFDweCweBsqkwfgyFrsqcwfasqEwfDsqCsqBliksrgwfrlicsraliEsrDliCliBCykljgsrrCycljaCyEljDCyCCyBCzgljrCzaCzDCzryhczaqarwyhEzananyyhCalzyhBwdcyEqwvcwdEyEnwvEyhnwvCwdBwvBsncwdqtrcsnEwdntrEwvntrCsnBtrBlbcsnqnjclbEsnnnjEtrnnjClbBnjBCjclbqazcCjElbnazEnjnazCCjBazBCjqazqCjnaznzioirsrfyziminwrdzzililyikzygozafafyyxozivivyadzyxmyglitzyxlwcoyEfwtowcmxvoyxvwclxvmwtlxvlslowcvtnoslmvrotnmsllvrmtnlvrllDoslvnbolDmrjonbmlDlrjmnblrjlCbolDvajoCbmizoajmCblizmajlizlCbvajvzieifwrFzzididyiczygeaFzywuy"
  CodageMC$(1) = CodageMC$(1) & "gdihzywtwcewsuwcdxtuwstxttskutlusktvnutltvntlBunDulBtrbunDtrbtCDuabuCDtijuabtijtziFiFyiEzygFywhwcFwshxsxskhtkxvlxlAxnBxrDxCBxaDxibxiCzwFcyCqwFEyCnwFCwFBsfcwFqsfEwFnsfCsfBkrcsfqkrEsfnkrCkrBBjckrqBjEkrnBjCBjBBjqBjnyaozDfDfyyamDdzyalwEoyCfwhowEmwhmwElwhlsdowEvsvosdmsvmsdlsvlknosdvlroknmlrmknllrlBboknvDjoBbmDjmBblDjlBbvDjvzbebfwnpzzbdbdybczyaeDFzyiuyadbhzyitwEewguwEdwxuwgtwxtscustuscttvustttvtklulnukltnrulntnrtBDuDbuBDtbjuDbtbjtjfsrpyjdwrozjcyjcjzbFbFyzjhjhybEzjgzyaFyihyyxwEFwghwwxxxxschssxttxvvxkkxllxnnxrrxBBxDDxbbxjFwrmzjEyjEjbCzjazjCyjCjjBjwCowCmwClsFowCvsFmsFlkfosFvkfmkflArokfvArmArlArvyDeBpzyDdwCewauwCdwatsEushusEtshtkdukvukdtkvtAnuBruAntBrtzDpDpyDozyDFybhwCFwahwixsEhsgxsxxkcxktxlvxAlxBnxDrxbpwnuzboybojDmzbqzjpsruyjowrujjoijobbmyjqybmjjqjjmwrtjjmijmbbljjnjjlijlbjkrsCusCtkFukFtAfuAftwDhsChsaxkExkhxAdxAvxBuzDuyDujbuwnxjbuibubDtjbvjjusrxijugrxbjuajuDbtijvibtbjvbjtgrwrjtajtDbsrjtrjsqjsnBxjDxiDxbbxgnyrbxabxDDwrbxrbwqbwn"
  CodageMC$(2) = "pjkurwejApbsunyebkpDwulzeDspByeBwzfcfjkprwzfEfbspnyzfCfDwplzzfBfByyrczfqfrwyrEzfnfnyyrCflzyrBxjcyrqxjEyrnxjCxjBuzcxjquzExjnuzCuzBpzcuzqpzEuznpzCdjAorsufydbkonwudzdDsolydBwokzdAyzdodrsovyzdmdnwotzzdldlydkzynozdvdvyynmdtzynlxboynvxbmxblujoxbvujmujlozoujvozmozlcrkofwuFzcnsodyclwoczckyckjzcucvwohzzctctycszylucxzyltxDuxDtubuubtojuojtcfsoFycdwoEzccyccjzchchycgzykxxBxuDxcFwoCzcEycEjcazcCycCjFjAmrstfyFbkmnwtdzFDsmlyFBwmkzFAyzFoFrsmvyzFmFnwmtzzFlFlyFkzyfozFvFvyyfmFtzyflwroyfvwrmwrltjowrvtjmtjlmzotjvmzmmzlqrkvfwxpzhbAqnsvdyhDkqlwvczhBsqkyhAwqkjhAiErkmfwtFzhrkEnsmdyhnsqtymczhlwEkyhkyEkjhkjzEuEvwmhzzhuzEthvwEtyzhthtyEszhszyduExzyvuydthxzyvtwnuxruwntxrttbuvjutbtvjtmjumjtgrAqfsvFygnkqdwvEzglsqcygkwqcjgkigkbEfsmFygvsEdwmEzgtwqgzgsyEcjgsjzEhEhyzgxgxyEgzgwzycxytxwlxxnxtDxvbxmbxgfkqFwvCzgdsqEygcwqEjgcigcbEFwmCzghwEEyggyEEjggjEazgizgFsqCygEwqCjgEigEbECygayECjgajgCwqBjgCigCbEBjgDjgBigBbCrklfwspzCnsldyClwlczCkyCkjzCuCvwlhzzCtCtyCszyFuCx"
  CodageMC$(2) = CodageMC$(2) & "zyFtwfuwftsrusrtljuljtarAnfstpyankndwtozalsncyakwncjakiakbCfslFyavsCdwlEzatwngzasyCcjasjzChChyzaxaxyCgzawzyExyhxwdxwvxsnxtrxlbxrfkvpwxuzinArdsvoyilkrcwvojiksrciikgrcbikaafknFwtmzivkadsnEyitsrgynEjiswaciisiacbisbCFwlCzahwCEyixwagyCEjiwyagjiwjCazaiziyzifArFsvmyidkrEwvmjicsrEiicgrEbicaicDaFsnCyihsaEwnCjigwrajigiaEbigbCCyaayCCjiiyaajiijiFkrCwvljiEsrCiiEgrCbiEaiEDaCwnBjiawaCiiaiaCbiabCBjaDjibjiCsrBiiCgrBbiCaiCDaBiiDiaBbiDbiBgrAriBaiBDaAriBriAqiAnBfskpyBdwkozBcyBcjBhyBgzyCxwFxsfxkrxDfklpwsuzDdsloyDcwlojDciDcbBFwkmzDhwBEyDgyBEjDgjBazDizbfAnpstuybdknowtujbcsnoibcgnobbcabcDDFslmybhsDEwlmjbgwDEibgiDEbbgbBCyDayBCjbiyDajbijrpkvuwxxjjdArosvuijckrogvubjccroajcEroDjcCbFknmwttjjhkbEsnmijgsrqinmbjggbEajgabEDjgDDCwlljbawDCijiwbaiDCbjiibabjibBBjDDjbbjjjjjFArmsvtijEkrmgvtbjEcrmajEErmDjECjEBbCsnlijasbCgnlbjagrnbjaabCDjaDDBibDiDBbjbibDbjbbjCkrlgvsrjCcrlajCErlDjCCjCBbBgnkrjDgbBajDabBDjDDDArbBrjDrjBcrkqjBErknjBCjBBbAqjBqbAnjBnjAorkfjAmjAlb"
  CodageMC$(2) = CodageMC$(2) & "AfjAvApwkezAoyAojAqzBpskuyBowkujBoiBobAmyBqyAmjBqjDpkluwsxjDosluiDoglubDoaDoDBmwktjDqwBmiDqiBmbDqbAljBnjDrjbpAnustxiboknugtxbbocnuaboEnuDboCboBDmsltibqsDmgltbbqgnvbbqaDmDbqDBliDniBlbbriDnbbrbrukvxgxyrrucvxaruEvxDruCruBbmkntgtwrjqkbmcntajqcrvantDjqEbmCjqCbmBjqBDlglsrbngDlajrgbnaDlDjrabnDjrDBkrDlrbnrjrrrtcvwqrtEvwnrtCrtBblcnsqjncblEnsnjnErtnjnCblBjnBDkqblqDknjnqblnjnnrsovwfrsmrslbkonsfjlobkmjlmbkljllDkfbkvjlvrsersdbkejkubkdjktAeyAejAuwkhjAuiAubAdjAvjBuskxiBugkxbBuaBuDAtiBviAtbBvbDuklxgsyrDuclxaDuElxDDuCDuBBtgkwrDvglxrDvaBtDDvDAsrBtrDvrnxctyqnxEtynnxCnxBDtclwqbvcnxqlwnbvEDtCbvCDtBbvBBsqDtqBsnbvqDtnbvnvyoxzfvymvylnwotyfrxonwmrxmnwlrxlDsolwfbtoDsmjvobtmDsljvmbtljvlBsfDsvbtvjvvvyevydnwerwunwdrwtDsebsuDsdjtubstjttvyFnwFrwhDsFbshjsxAhiAhbAxgkirAxaAxDAgrAxrBxckyqBxEkynBxCBxBAwqBxqAwnBxnlyoszflymlylBwokyfDxolyvDxmBwlDxlAwfBwvDxvtzetzdlyenyulydnytBweDwuBwdbxuDwtbxttzFlyFnyhBwFDwhbwxAiqAinAyokjfAymAylAifAyvkzekzdAyeByuAydBytszp"
  CodeErr% = 0
  If Chaine$ = "" Then CodeErr% = 1: Exit Function
  'Découper la chaine en blocs de caractère de même type : numérique , texte, octet
  'La 1ère colonne du tableau Liste% contient le nombre de caractères, la 2ème le commutateur de mode
  'Split the string in character blocks of the same type : numeric , text, byte
  'The first column of the array Liste% contain the char. number, the second one contain the mode switch
  IndexChaine% = 1
  GoSub QuelMode
  Do
    ReDim Preserve Liste%(1, IndexListe%)
    Liste%(1, IndexListe%) = Mode%
    Do While Liste%(1, IndexListe%) = Mode%
      Liste%(0, IndexListe%) = Liste%(0, IndexListe%) + 1
      IndexChaine% = IndexChaine% + 1
      If IndexChaine% > Len(Chaine$) Then Exit Do
      GoSub QuelMode
    Loop
    IndexListe% = IndexListe% + 1
  Loop Until IndexChaine% > Len(Chaine$)
  'Ne garder le mode numérique que si c'est "rentable", sinon mode "texte" voire "octet"
  'Les seuils de rentabilité ont étés pré-déterminés selon le mode précédent et/ou le mode suivant
  'We retain "numeric" mode only if it's earning, else "text" mode or even "byte" mode
  'The efficiency limits have been pre-defined according to the previous mode and/or the next mode.
  For I% = 0 To IndexListe% - 1
    If Liste%(1, I%) = 902 Then
      If I% = 0 Then 'C'est le premier bloc / It's the first block
        If IndexListe% > 1 Then 'et il y en a d'autres derrière / And there is other blocks behind
          If Liste%(1, I% + 1) = 900 Then
            'Premier bloc et suivi par un bloc de type "texte" / First block and followed by a "text" type block
            If Liste%(0, I%) < 8 Then Liste%(1, I%) = 900
          ElseIf Liste%(1, I% + 1) = 901 Then
            'Premier bloc et suivi par un bloc de type "octet" / First block and followed by a "byte" type block
            If Liste%(0, I%) = 1 Then Liste%(1, I%) = 901
          End If
        End If
      Else
        'C'est pas le premier bloc / It's not the first block
        If I% = IndexListe% - 1 Then
          'C'est le dernier / It's the last one
          If Liste%(1, I% - 1) = 900 Then
            'Il est précédé par un bloc de type "texte" / It's  preceded by a "text" type block
            If Liste%(0, I%) < 7 Then Liste%(1, I%) = 900
          ElseIf Liste%(1, I% - 1) = 901 Then
            'Il est précédé par un bloc de type "octet" / It's  preceded by a "byte" type block
            If Liste%(0, I%) = 1 Then Liste%(1, I%) = 901
          End If
        Else
          'C'est pas le dernier / It's not the last block
          If Liste%(1, I% - 1) = 901 And Liste%(1, I% + 1) = 901 Then
            'Encadré par des blocs de type "octet" / Framed by "byte" type blocks
            If Liste%(0, I%) < 4 Then Liste%(1, I%) = 901
          ElseIf Liste%(1, I% - 1) = 900 And Liste%(1, I% + 1) = 901 Then
            'Précédé par "texte" et suivi par "octet" (Si l'inverse jamais intéressant de changer)
            'Preceded by "text" and followed by "byte" (If the reverse it's never interesting to change)
            If Liste%(0, I%) < 5 Then Liste%(1, I%) = 900
          ElseIf Liste%(1, I% - 1) = 900 And Liste%(1, I% + 1) = 900 Then
            'Encadré par des blocs de type "texte" / Framed by "text" type blocks
            If Liste%(0, I%) < 8 Then Liste%(1, I%) = 900
          End If
        End If
      End If
    End If
  Next
  GoSub Regroupe
  'Ne garder le mode "texte" que si c'est rentable / Maintain "text" mode only if it's earning
  For I% = 0 To IndexListe% - 1
    If Liste%(1, I%) = 900 And I% > 0 Then
      'C'est pas le premier (Si 1er jamais intéressant de changer) / It's not the first (If first, never interesting to change)
      If I% = IndexListe% - 1 Then 'C'est le dernier / It's the last one
        If Liste%(1, I% - 1) = 901 Then
          'Précédé par un bloc de type "octet" / It's  preceded by a "byte" type block
          If Liste%(0, I%) = 1 Then Liste%(1, I%) = 901
        End If
      Else
        'C'est pas le dernier / It's not the last one
        If Liste%(1, I% - 1) = 901 And Liste%(1, I% + 1) = 901 Then
          'Encadré par des blocs de type "octet" / Framed by "byte" type blocks
          If Liste%(0, I%) < 5 Then Liste%(1, I%) = 901
        ElseIf (Liste%(1, I% - 1) = 901 And Liste%(1, I% + 1) <> 901) Or (Liste%(1, I% - 1) <> 901 And Liste%(1, I% + 1) = 901) Then
          'Un bloc "octet" devant ou derrière / A "byte" block ahead or behind
          If Liste%(0, I%) < 3 Then Liste%(1, I%) = 901
        End If
      End If
    End If
  Next
  GoSub Regroupe
  'Maintenant on compacte les données dans les MC, les MC sont stockées sur 3 car. dans une grande chaine : ChaineMC$
  'Now we compress datas into the MCs, the MCs are stored in 3 char. in a large string : ChaineMC$
  IndexChaine% = 1
  For I% = 0 To IndexListe% - 1
    'Donc 3 modes de compactage / Thus 3 compaction modes
    Select Case Liste%(1, I%)
    Case 900 'Texte
      ReDim ListeT%(1, Liste%(0, I%))
      'ListeT% contiendra le numéro de table(s) et la valeur de chaque caractère
      'Numéros de table codés sur les 4 bits de poids faibles, soit en décimal 1, 2, 4, 8
      'ListeT% will contain the table number(s) (1 ou several) and the value of each char.
      'Table number encoded in the 4 less weight bits, that is in decimal 1, 2, 4, 8
      For IndexListeT% = 0 To Liste%(0, I%) - 1
        CodeASCII% = Asc(Mid$(Chaine$, IndexChaine% + IndexListeT%, 1))
        Select Case CodeASCII%
        Case 9 'HT
          ListeT%(0, IndexListeT%) = 12
          ListeT%(1, IndexListeT%) = 12
        Case 10 'LF
          ListeT%(0, IndexListeT%) = 8
          ListeT%(1, IndexListeT%) = 15
        Case 13 'CR
          ListeT%(0, IndexListeT%) = 12
          ListeT%(1, IndexListeT%) = 11
        Case Else
          ListeT%(0, IndexListeT%) = Mid$(ASCII$, CodeASCII% * 4 - 127, 2)
          ListeT%(1, IndexListeT%) = Mid$(ASCII$, CodeASCII% * 4 - 125, 2)
        End Select
      Next
      CurTable% = 1 'Table par défaut / Default table
      ChaineT$ = ""
      'Les données sont stockées sur 2 car. dans la chaine TableT$ / Datas are stored in 2 char. in the string TableT$
      For J% = 0 To Liste%(0, I%) - 1
        If (ListeT%(0, J%) And CurTable%) > 0 Then
          'Le car. est dans la table courante / The char. is in the current table
          ChaineT$ = ChaineT$ & Format(ListeT%(1, J%), "00")
        Else
          'Faut changer de table / Obliged to change the table
          flag = False 'True si on change de table pour un seul car. / True if we change the table only for 1 char.
          If J% = Liste%(0, I%) - 1 Then
            flag = True
          Else
            If (ListeT%(0, J%) And ListeT%(0, J% + 1)) = 0 Then flag = True 'Pas de table commune avec le car. suivant / No common table with the next char.
          End If
          If flag Then
            'On change de table pour 1 seul car., Chercher un commutateur fugitif
            'We change only for 1 char., Look for a temporary switch
            If (ListeT%(0, J%) And 1) > 0 And CurTable% = 2 Then
              'Table 2 vers 1 pour 1 car. --> T_MAJ / Table 2 to 1 for 1 char. --> T_UPP
              ChaineT$ = ChaineT$ & "27" & Format(ListeT%(1, J%), "00")
            ElseIf (ListeT%(0, J%) And 8) > 0 Then
              'Table 1 ou 2 ou 4 vers table 8 pour 1 car. --> T_PON / Table 1 or 2 or 4 to table 8 for 1 char. --> T_PUN
              ChaineT$ = ChaineT$ & "29" & Format(ListeT%(1, J%), "00")
            Else
              'Pas de commutateur fugitif / No temporary switch available
              flag = False
            End If
          End If
          If Not flag Then 'On re-teste flag qui a peut-être changé ci-dessus ! donc ELSE pas possible / We test again flag which is perhaps changed ! Impossible tio use ELSE statement
            '
            'On doit utiliser un commutateur à basculement
            'Déterminer la nouvelle table à utiliser
            'We must use a bi-state switch
            'Looking for the new table to use
            If J% = Liste%(0, I%) - 1 Then
              NewTable% = ListeT%(0, J%)
            Else
              NewTable% = IIf((ListeT%(0, J%) And ListeT%(0, J% + 1)) = 0, ListeT%(0, J%), ListeT%(0, J%) And ListeT%(0, J% + 1))
            End If
            'Ne garder que la première s'il y en a plusieurs de possible / Maintain the first if several tables are possible
            Select Case NewTable%
            Case 3, 5, 7, 9, 11, 13, 15
              NewTable% = 1
            Case 6, 10, 14
              NewTable% = 2
            Case 12
              NewTable% = 4
            End Select
            'Choisir le commutateur, parfois il faut 2 commutateurs de suite / Select the switch, on occasion we must use 2 switchs consecutively
            Select Case CurTable%
            Case 1
              Select Case NewTable%
              Case 2
                ChaineT$ = ChaineT$ & "27"
              Case 4
                ChaineT$ = ChaineT$ & "28"
              Case 8
                ChaineT$ = ChaineT$ & "2825"
              End Select
            Case 2
              Select Case NewTable%
              Case 1
                ChaineT$ = ChaineT$ & "2828"
              Case 4
                ChaineT$ = ChaineT$ & "28"
              Case 8
                ChaineT$ = ChaineT$ & "2825"
              End Select
            Case 4
              Select Case NewTable%
              Case 1
                ChaineT$ = ChaineT$ & "28"
              Case 2
                ChaineT$ = ChaineT$ & "27"
              Case 8
                ChaineT$ = ChaineT$ & "25"
              End Select
            Case 8
              Select Case NewTable%
              Case 1
                ChaineT$ = ChaineT$ & "29"
              Case 2
                ChaineT$ = ChaineT$ & "2927"
              Case 4
                ChaineT$ = ChaineT$ & "2928"
              End Select
            End Select
            CurTable% = NewTable%
            ChaineT$ = ChaineT$ & Format(ListeT%(1, J%), "00") 'On ajoute enfin le car. / At last we add the char.
          End If
        End If
      Next
      If Len(ChaineT$) Mod 4 > 0 Then ChaineT$ = ChaineT$ & "29" 'Bourrage si nb de car. impair / Padding if number of char. is odd
      'Maintenant traduire la chaine ChaineT$ en MCs
      'Now translate the string ChaineT$ into CWs
      If I% > 0 Then ChaineMC$ = ChaineMC$ & "900" 'Mettre en place le commutateur sauf si premier bloc car mode "texte" par défaut / Set up the switch exept for the first block because "text" is the default
      For J% = 1 To Len(ChaineT$) Step 4
        ChaineMC$ = ChaineMC$ & Format(Mid$(ChaineT$, J%, 2) * 30 + Mid$(ChaineT$, J% + 2, 2), "000")
      Next
    Case 901 'Octet
      'Choisir le commutateur parmi les 3 possibles / Select the switch between the 3 possible
      If Liste%(0, I%) = 1 Then
        '1 seul octet, c'est immédiat
        ChaineMC$ = ChaineMC$ & "913" & Format(Asc(Mid$(Chaine$, IndexChaine%, 1)), "000")
      Else
        'Choisir le commutateur selon qu'on a un multiple de 6 octets ou non
        'Select the switch for perfect multiple of 6 bytes or no
        If Liste%(0, I%) Mod 6 = 0 Then
          ChaineMC$ = ChaineMC$ & "924"
        Else
          ChaineMC$ = ChaineMC$ & "901"
        End If
        J% = 0
        Do While J% < Liste%(0, I%)
          Longueur% = Liste%(0, I%) - J%
          If Longueur% >= 6 Then
            'prendre des paquets de 6 /Take groups of 6
            Longueur% = 6
            Total = 0
            For K% = 0 To Longueur% - 1
              Total = Total + (Asc(Mid$(Chaine$, IndexChaine% + J% + K%, 1)) * 256 ^ (Longueur% - 1 - K%))
            Next
            ChaineMod$ = Format(Total, "general number")
            Dummy$ = ""
            Do
              Diviseur& = 900
              GoSub Modulo
              Dummy$ = Format(Diviseur&, "000") & Dummy$
              ChaineMod$ = ChaineMult$
              If ChaineMult$ = "" Then Exit Do
            Loop
            ChaineMC$ = ChaineMC$ & Dummy$
          Else
            'S'il reste un paquet de moins de 6 octets / If it remain a group of less than 6 bytes
            For K% = 0 To Longueur% - 1
              ChaineMC$ = ChaineMC$ & Format(Asc(Mid$(Chaine$, IndexChaine% + J% + K%, 1)), "000")
            Next
          End If
          J% = J% + Longueur%
        Loop
      End If
    Case 902 'Numérique / Numeric
      ChaineMC$ = ChaineMC$ & "902"
      J% = 0
      Do While J% < Liste%(0, I%)
        Longueur% = Liste%(0, I%) - J%
        If Longueur% > 44 Then Longueur% = 44
        ChaineMod$ = "1" & Mid$(Chaine$, IndexChaine% + J%, Longueur%)
        Dummy$ = ""
        Do
          Diviseur& = 900
          GoSub Modulo
          Dummy$ = Format(Diviseur&, "000") & Dummy$
          ChaineMod$ = ChaineMult$
          If ChaineMult$ = "" Then Exit Do
        Loop
        ChaineMC$ = ChaineMC$ & Dummy$
        J% = J% + Longueur%
      Loop
      Debug.Print ChaineMC
    End Select
    IndexChaine% = IndexChaine% + Liste%(0, I%)
  Next
  'ChaineMC$ contient la liste des MC (sur 3 chiffres) représentant les données
  'On s'occupe maintenant du niveau de correction
  'ChaineMC$ contain the MC list (on 3 digits) depicting the datas
  'Now we take care of the correction level
  Longueur% = Len(ChaineMC$) / 3
  If sécu% < 0 Then
    'Détermination auto. du niveau de correction en fonction des recommandations de la norme
    'Fixing auto. the correction level according to the standard recommendations
    If Longueur% < 41 Then
      sécu% = 2
    ElseIf Longueur% < 161 Then
      sécu% = 3
    ElseIf Longueur% < 321 Then
      sécu% = 4
    Else
      sécu% = 5
    End If
  End If
  'On s'occupe maintenant du nombre de MC par ligne / Now we take care of the number of CW per row
  Longueur% = Longueur% + 1 + (2 ^ (sécu% + 1))
  If nbcol% > 30 Then nbcol% = 30
  If nbcol% < 1 Then
    'Avec une police haute de 3 modules, pour obtenir un code-barre "carré"
    'x = nb. de col. | Largeur en module = 69 + 17x | Hauteur en module = 3t / x (t étant le nb total de MC)
    'On a donc 69 + 17x = 3t/x <=> 17x²+69x-3t=0 - Le discriminant est 69²-4*17*-3t = 4761+204t donc x=SQR(discr.)-69/2*17
    '
    'With a 3 modules high font, for getting a "square" bar code
    'x = nb. of col. | Width by module = 69 + 17x | Height by module = 3t / x (t is the total number of MCs)
    'Thus we have 69 + 17x = 3t/x <=> 17x²+69x-3t=0 - Discriminant is 69²-4*17*-3t = 4761+204t thus x=SQR(discr.)-69/2*17
    nbcol% = (Sqr(204# * Longueur% + 4761) - 69) / (34 / 1.3)   '1.3 = coeff. de pondération déterminé au pif après essais / 1.3 = balancing factor determined at a guess after tests
    If nbcol% = 0 Then nbcol% = 1
  End If
  'Si on dépasse 928 MC on essaye de réduire le niveau de correction
  'If we go beyong 928 CWs we try to reduce the correction level
  Do While sécu% > 0
    'Calcul du nombre total de MC en tenant compte du rembourrage pour compléter les lignes
    'Calculation of the total number of CW with the padding
    Longueur% = Len(ChaineMC$) / 3 + 1 + (2 ^ (sécu% + 1))
    Longueur% = (Longueur% \ nbcol% + IIf(Longueur% Mod nbcol% > 0, 1, 0)) * nbcol%
    If Longueur% < 929 Then Exit Do
    'On doit réduire le niveau de sécurité pour tout faire rentrer
    'We must reduce security level
    sécu% = sécu% - 1
    CodeErr% = 10
  Loop
  If Longueur% > 928 Then CodeErr% = 2: Exit Function
  If Longueur% / nbcol% > 90 Then CodeErr% = 3: Exit Function
  'Calcul du rembourrage / Padding calculation
  Longueur% = Len(ChaineMC$) / 3 + 1 + (2 ^ (sécu% + 1))
  I% = 0
  If Longueur% \ nbcol% < 3 Then
    I% = nbcol% * 3 - Longueur%   'Il faut au moins 3 lignes dans le code / A bar code must have at least 3 row
  Else
    If Longueur% Mod nbcol% > 0 Then I% = nbcol% - (Longueur% Mod nbcol%)
  End If
  'On ajoute le rembourrage / We add the padding
  Do While I% > 0
    ChaineMC$ = ChaineMC$ & "900"
    I% = I% - 1
  Loop
  'On ajoute le descripteur de longueur / We add the length descriptor
  ChaineMC$ = Format(Len(ChaineMC$) / 3 + 1, "000") & ChaineMC$
  'On s'occupe maintenant des codes de Reed Solomon / Now we take care of the Reed Solomon codes
  Longueur% = Len(ChaineMC$) / 3
  K% = 2 ^ (sécu% + 1)
  ReDim MCcorrection%(K% - 1)
  Total = 0
  For I% = 0 To Longueur% - 1
    Total = (Mid$(ChaineMC$, I% * 3 + 1, 3) + MCcorrection%(K% - 1)) Mod 929
    For J% = K% - 1 To 0 Step -1
      If J% = 0 Then
        MCcorrection%(J%) = (929 - (Total * Mid$(CoefRS$(sécu%), J% * 3 + 1, 3)) Mod 929) Mod 929
      Else
        MCcorrection%(J%) = (MCcorrection%(J% - 1) + 929 - (Total * Mid$(CoefRS$(sécu%), J% * 3 + 1, 3)) Mod 929) Mod 929
      End If
    Next
  Next
  For J% = 0 To K% - 1
    If MCcorrection%(J%) <> 0 Then MCcorrection%(J%) = 929 - MCcorrection%(J%)
  Next
  'On va ajouter les codes de correction à la chaine / We add theses codes to the string
  For I% = K% - 1 To 0 Step -1
    ChaineMC$ = ChaineMC$ & Format(MCcorrection%(I%), "000")
  Next
  'La chaine des MC est terminée
  'Calcul des paramètres pour les MC de cotés gauche et droit
  'The CW string is finished
  'Calculation of parameters for the left and right side CWs
  C1% = (Len(ChaineMC$) / 3 / nbcol% - 1) \ 3
  C2% = sécu% * 3 + (Len(ChaineMC$) / 3 / nbcol% - 1) Mod 3
  C3% = nbcol% - 1
  'On encode chaque ligne / We encode each row
  For I% = 0 To Len(ChaineMC$) / 3 / nbcol% - 1
    Dummy$ = Mid$(ChaineMC$, I% * nbcol% * 3 + 1, nbcol% * 3)
    K% = (I% \ 3) * 30
    Select Case I% Mod 3
    Case 0
      Dummy$ = Format(K% + C1%, "000") & Dummy$ & Format(K% + C3%, "000")
    Case 1
      Dummy$ = Format(K% + C2%, "000") & Dummy$ & Format(K% + C1%, "000")
    Case 2
      Dummy$ = Format(K% + C3%, "000") & Dummy$ & Format(K% + C2%, "000")
    End Select
    pdf417$ = pdf417$ & "+*" 'Commencer par car. de start et séparateur / Start with a start char. and a separator
    For J% = 0 To Len(Dummy$) / 3 - 1
      pdf417$ = pdf417$ & Mid$(CodageMC$(I% Mod 3), Mid$(Dummy$, J% * 3 + 1, 3) * 3 + 1, 3) & "*"
    Next
    pdf417$ = pdf417$ & "-" & Chr$(13) & Chr$(10) 'Ajouter car. de stop et CRLF / Add a stop char. and a CRLF
  Next
  Exit Function
Regroupe:
  'Regrouper les blocs de même type / Bring together same type blocks
  If IndexListe% > 1 Then
    I% = 1
    Do While I% < IndexListe%
      If Liste%(1, I% - 1) = Liste%(1, I%) Then
        'Regroupement / Bringing together
        Liste%(0, I% - 1) = Liste%(0, I% - 1) + Liste%(0, I%)
        J% = I% + 1
        'Réduction de la liste / Decrease the list
        Do While J% < IndexListe%
          Liste%(0, J% - 1) = Liste%(0, J%)
          Liste%(1, J% - 1) = Liste%(1, J%)
          J% = J% + 1
        Loop
        IndexListe% = IndexListe% - 1
        I% = I% - 1
      End If
      I% = I% + 1
    Loop
  End If
Return
QuelMode:
  CodeASCII% = Asc(Mid$(Chaine$, IndexChaine%, 1))
  Select Case CodeASCII%
  Case 48 To 57
    Mode% = 902
  Case 9, 10, 13, 32 To 126
    Mode% = 900
  Case Else
    Mode% = 901
  End Select
Return
Modulo:
  'ChaineMod$ représente un très grand nombre sur plus de 9 chiffres
  'Diviseur& est le diviseur, contient le résultat au retour
  'ChaineMult$ contient au retour le résultat de la division entière
  '
  'ChaineMod$ depict a very large number having more than 9 digits
  'Diviseur& is the divisor, contain the result after return
  'ChaineMult$ contain after return the result of the integer division
  ChaineMult$ = ""
  Nombre& = 0
  Do While ChaineMod$ <> ""
    Nombre& = Nombre& * 10 + Left$(ChaineMod$, 1) 'Abaisse un chiffre / Put down a digit
    ChaineMod$ = Mid$(ChaineMod$, 2)
    If Nombre& < Diviseur& Then
      If ChaineMult$ <> "" Then ChaineMult$ = ChaineMult$ & "0"
    Else
      ChaineMult$ = ChaineMult$ & Nombre& \ Diviseur&
    End If
    Nombre& = Nombre& Mod Diviseur& 'Récupère le reste / get the remainder
  Loop
  Diviseur& = Nombre&
Return
End Function
