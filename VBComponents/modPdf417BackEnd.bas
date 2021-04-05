Attribute VB_Name = "modPdf417BackEnd"
Option Explicit
'Source: http://grandzebu.net/informatique/codbar-en/pdf417.htm
'Tables
Dim ASCII
Dim CoefRS(8)
Dim CodageMC(2)

Private Sub initASCII()
'This string describes the ASCII for the "text" mode
'ASCII contain 95 fields of 4 digits which correspond to characters ASCII 32 to 126. These fields :
'   2 digits indicating the table(s) (1 or several) where this character is located. (Table numbers : 1, 2, 4, and 8)
'   2 digits indicating the char. number in the table
'   Sample : 0726 at the beginning of the string : The Char. having code 32 is in the tables 1, 2 and 4 at row 26
ASCII = "07260810082004151218042104100828082308241222042012131216121712190400040104020403040404050406040704080409121408000801042308020825080301000101010201030104010501060107010801090110011101120113011401150116011701180119012001210122012301240125080408050806042408070808020002010202020302040205020602070208020902100211021202130214021502160217021802190220022102220223022402250826082108270809"
End Sub

Private Sub initCoefRS()
'CoefRS contain 8 strings describing the factors of the polynomial equations for the reed Solomon codes.
CoefRS(0) = "027917"
CoefRS(1) = "522568723809"
CoefRS(2) = "237308436284646653428379"
CoefRS(3) = "274562232755599524801132295116442428295042176065"
CoefRS(4) = "361575922525176586640321536742677742687284193517273494263147593800571320803133231390685330063410"
CoefRS(5) = "539422006093862771453106610287107505733877381612723476462172430609858822543376511400672762283184440035519031460594225535517352605158651201488502648733717083404097280771840629004381843623264543"
CoefRS(6) = "521310864547858580296379053779897444400925749415822093217208928244583620246148447631292908490704516258457907594723674292272096684432686606860569193219129186236287192775278173040379712463646776171491297763156732095270447090507048228821808898784663627378382262380602754336089614087432670616157374242726600269375898845454354130814587804034211330539297827865037517834315550086801004108539"
CoefRS(7) = "524894075766882857074204082586708250905786138720858194311913275190375850438733194280201280828757710814919089068569011204796605540913801700799137439418592668353859370694325240216257284549209884315070329793490274877162749812684461334376849521307291803712019358399908103511051008517225289470637731066255917269463830730433848585136538906090002290743199655903329049802580355588188462010134628320479130739071263318374601192605142673687234722384177752607640455193689707805641048060732621895544261852655309697755756060231773434421726528503118049795032144500238836394280566319009647550073914342126032681331792620060609441180791893754605383228749760213054297134054834299922191910532609829189020167029872449083402041656505579481173404251688095497555642543307159924558648055497010"
CoefRS(8) = "352077373504035599428207409574118498285380350492197265920155914299229643294871306088087193352781846075327520435543203666249346781621640268794534539781408390644102476499290632545037858916552041542289122272383800485098752472761107784860658741290204681407855085099062482180020297451593913142808684287536561076653899729567744390513192516258240518794395768848051610384168190826328596786303570381415641156237151429531207676710089168304402040708575162864229065861841512164477221092358785288357850836827736707094008494114521002499851543152729771095248361578323856797289051684466533820669045902452167342244173035463651051699591452578037124298332552043427119662777475850764364578911283711472420245288594394511327589777699688043408842383721521560644714559062145873663713159672729"
CoefRS(8) = CoefRS(8) & "624059193417158209563564343693109608563365181772677310248353708410579870617841632860289536035777618586424833077597346269757632695751331247184045787680018066407369054492228613830922437519644905789420305441207300892827141537381662513056252341242797838837720224307631061087560310756665397808851309473795378031647915459806590731425216548249321881699535673782210815905303843922281073469791660162498308155422907817187062016425535336286437375273610296183923116667751353062366691379687842037357720742330005039923311424242749321054669316342299534105667488640672576540316486721610046656447171616464190531297321762752533175134014381433717045111020596284736138646411877669141919045780407164332899165726600325498655357752768223849647063310863251366304282738675410389244031121303263"
End Sub

Private Sub initCodageMC()
'CodageMC contain the 3 sets of the 929 MCs. Each MC is described in the PDF417.TTF font by 3 char. composing 3 time 5 bits. The first bit which is always 1
' and the last one which is always 0 are into the separator character.
CodageMC(0) = "urAxfsypyunkxdwyozpDAulspBkeBApAseAkprAuvsxhypnkutwxgzfDAplsfBkfrApvsuxyfnkptwuwzflspsyfvspxyftwpwzfxyyrxufkxFwymzonAudsxEyolkucwdBAoksucidAkokgdAcovkuhwxazdnAotsugydlkoswugjdksosidvkoxwuizdtsowydswowjdxwoyzdwydwjofAuFsxCyodkuEwxCjclAocsuEickkocgckcckEcvAohsuayctkogwuajcssogicsgcsacxsoiycwwoijcwicyyoFkuCwxBjcdAoEsuCicckoEguCbcccoEaccEoEDchkoawuDjcgsoaicggoabcgacgDobjcibcFAoCsuBicEkoCguBbcEcoCacEEoCDcECcascagcaacCkuAroBaoBDcCBtfkwpwyezmnAtdswoymlktcwwojFBAmksFAkmvkthwwqzFnAmtstgyFlkmswFksFkgFvkmxwtizFtsmwyFswFsiFxwmyzFwyFyzvfAxpsyuyvdkxowyujqlAvcsxoiqkkvcgxobqkcvcamfAtFswmyqvAmdktEwwmjqtkvgwxqjhlAEkkmcgtEbhkkqsghkcEvAmhstayhvAEtkmgwtajhtkqwwvijhssEsghsgExsmiyhxsEwwmijhwwqyjhwiEyyhyyEyjhyjvFkxmwytjqdAvEsxmiqckvEgxmbqccvEaqcEqcCmFktCwwljqhkmEstCigtAEckvaitCbgskEccmEagscqgamEDEcCEhkmawtDjgxkEgsmaigwsqiimabgwgEgaEgDEiwmbjgywEiigyiEibgybgzjqFAvCsxliqEkvCgxlbqEcvCaqEEvCDqECqEBEFAmCstBighAEEkmCgtBbggkqagvDbggcEEEmCDggEqaDgg"
CodageMC(0) = CodageMC(0) & "CEasmDigisEagmDbgigqbbgiaEaDgiDgjigjbqCkvBgxkrqCcvBaqCEvBDqCCqCBECkmBgtArgakECcmBagacqDamBDgaEECCgaCECBEDggbggbagbDvAqvAnqBBmAqEBEgDEgDCgDBlfAspsweyldksowClAlcssoiCkklcgCkcCkECvAlhssqyCtklgwsqjCsslgiCsgCsaCxsliyCwwlijCwiCyyCyjtpkwuwyhjndAtoswuincktogwubncctoancEtoDlFksmwwdjnhklEssmiatACcktqismbaskngglEaascCcEasEChklawsnjaxkCgstrjawsniilabawgCgaawaCiwlbjaywCiiayiCibCjjazjvpAxusyxivokxugyxbvocxuavoExuDvoCnFAtmswtirhAnEkxviwtbrgkvqgxvbrgcnEEtmDrgEvqDnEBCFAlCssliahACEklCgslbixAagknagtnbiwkrigvrblCDiwcagEnaDiwECEBCaslDiaisCaglDbiysaignbbiygrjbCaDaiDCbiajiCbbiziajbvmkxtgywrvmcxtavmExtDvmCvmBnCktlgwsrraknCcxtrracvnatlDraEnCCraCnCBraBCCklBgskraakCCclBaiikaacnDalBDiicrbaCCCiiEaaCCCBaaBCDglBrabgCDaijgabaCDDijaabDCDrijrvlcxsqvlExsnvlCvlBnBctkqrDcnBEtknrDEvlnrDCnBBrDBCBclAqaDcCBElAnibcaDEnBnibErDnCBBibCaDBibBaDqibqibnxsfvkltkfnAmnAlCAoaBoiDoCAlaBlkpkBdAkosBckkogsebBcckoaBcEkoDBhkkqwsfjBgskqiBggkqbBgaBgDBiwkrjBiiBibBjjlpAsuswhil"
CodageMC(0) = CodageMC(0) & "oksuglocsualoEsuDloCBFAkmssdiDhABEksvisdbDgklqgsvbDgcBEEkmDDgElqDBEBBaskniDisBagknbDiglrbDiaBaDBbiDjiBbbDjbtukwxgyirtucwxatuEwxDtuCtuBlmkstgnqklmcstanqctvastDnqElmCnqClmBnqBBCkklgDakBCcstrbikDaclnaklDbicnraBCCbiEDaCBCBDaBBDgklrDbgBDabjgDbaBDDbjaDbDBDrDbrbjrxxcyyqxxEyynxxCxxBttcwwqvvcxxqwwnvvExxnvvCttBvvBllcssqnncllEssnrrcnnEttnrrEvvnllBrrCnnBrrBBBckkqDDcBBEkknbbcDDEllnjjcbbEnnnBBBjjErrnDDBjjCBBqDDqBBnbbqDDnjjqbbnjjnxwoyyfxwmxwltsowwfvtoxwvvtmtslvtllkossfnlolkmrnonlmlklrnmnllrnlBAokkfDBolkvbDoDBmBAljbobDmDBljbmbDljblDBvjbvxwdvsuvstnkurlurltDAubBujDujDtApAAokkegAocAoEAoCAqsAqgAqaAqDAriArbkukkucshakuEshDkuCkuBAmkkdgBqkkvgkdaBqckvaBqEkvDBqCAmBBqBAngkdrBrgkvrBraAnDBrDAnrBrrsxcsxEsxCsxBktclvcsxqsgnlvEsxnlvCktBlvBAlcBncAlEkcnDrcBnEAlCDrEBnCAlBDrCBnBAlqBnqAlnDrqBnnDrnwyowymwylswotxowyvtxmswltxlksosgfltoswvnvoltmkslnvmltlnvlAkokcfBloksvDnoBlmAklbroDnmBllbrmDnlAkvBlvDnvbrvyzeyzdwyexyuwydxytswetwuswdvxutwtvxtkselsuksdntulstrvu"
CodageMC(1) = "ypkzewxdAyoszeixckyogzebxccyoaxcEyoDxcCxhkyqwzfjutAxgsyqiuskxggyqbuscxgausExgDusCuxkxiwyrjptAuwsxiipskuwgxibpscuwapsEuwDpsCpxkuywxjjftApwsuyifskpwguybfscpwafsEpwDfxkpywuzjfwspyifwgpybfwafywpzjfyifybxFAymszdixEkymgzdbxEcymaxEEymDxECxEBuhAxasyniugkxagynbugcxaaugExaDugCugBoxAuisxbiowkuigxbbowcuiaowEuiDowCowBdxAoysujidwkoygujbdwcoyadwEoyDdwCdysozidygozbdyadyDdzidzbxCkylgzcrxCcylaxCEylDxCCxCBuakxDgylruacxDauaExDDuaCuaBoikubgxDroicubaoiEubDoiCoiBcykojgubrcycojacyEojDcyCcyBczgojrczaczDczrxBcykqxBEyknxBCxBBuDcxBquDExBnuDCuDBobcuDqobEuDnobCobBcjcobqcjEobncjCcjBcjqcjnxAoykfxAmxAluBoxAvuBmuBloDouBvoDmoDlcbooDvcbmcblxAexAduAuuAtoBuoBtwpAyeszFiwokyegzFbwocyeawoEyeDwoCwoBthAwqsyfitgkwqgyfbtgcwqatgEwqDtgCtgBmxAtiswrimwktigwrbmwctiamwEtiDmwCmwBFxAmystjiFwkmygtjbFwcmyaFwEmyDFwCFysmziFygmzbFyaFyDFziFzbyukzhghjsyuczhahbwyuEzhDhDyyuCyuBwmkydgzErxqkwmczhrxqcyvaydDxqEwmCxqCwmBxqBtakwngydrviktacwnavicxrawnDviEtaCviCtaBviBmiktbgwnrqykmictb"
CodageMC(1) = CodageMC(1) & "aqycvjatbDqyEmiCqyCmiBqyBEykmjgtbrhykEycmjahycqzamjDhyEEyChyCEyBEzgmjrhzgEzahzaEzDhzDEzrytczgqgrwytEzgngnyytCglzytBwlcycqxncwlEycnxnEytnxnCwlBxnBtDcwlqvbctDEwlnvbExnnvbCtDBvbBmbctDqqjcmbEtDnqjEvbnqjCmbBqjBEjcmbqgzcEjEmbngzEqjngzCEjBgzBEjqgzqEjngznysozgfgfyysmgdzyslwkoycfxloysvxlmwklxlltBowkvvDotBmvDmtBlvDlmDotBvqbovDvqbmmDlqblEbomDvgjoEbmgjmEblgjlEbvgjvysegFzysdwkexkuwkdxkttAuvButAtvBtmBuqDumBtqDtEDugbuEDtgbtysFwkFxkhtAhvAxmAxqBxwekyFgzCrwecyFaweEyFDweCweBsqkwfgyFrsqcwfasqEwfDsqCsqBliksrgwfrlicsraliEsrDliCliBCykljgsrrCycljaCyEljDCyCCyBCzgljrCzaCzDCzryhczaqarwyhEzananyyhCalzyhBwdcyEqwvcwdEyEnwvEyhnwvCwdBwvBsncwdqtrcsnEwdntrEwvntrCsnBtrBlbcsnqnjclbEsnnnjEtrnnjClbBnjBCjclbqazcCjElbnazEnjnazCCjBazBCjqazqCjnaznzioirsrfyziminwrdzzililyikzygozafafyyxozivivyadzyxmyglitzyxlwcoyEfwtowcmxvoyxvwclxvmwtlxvlslowcvtnoslmvrotnmsllvrmtnlvrllDoslvnbolDmrjonbmlDlrjmnblrjlCbolDvajoCbmizoajmCblizmajlizlCbvajvzieifwrFzzididyiczygeaFzywuy"
CodageMC(1) = CodageMC(1) & "gdihzywtwcewsuwcdxtuwstxttskutlusktvnutltvntlBunDulBtrbunDtrbtCDuabuCDtijuabtijtziFiFyiEzygFywhwcFwshxsxskhtkxvlxlAxnBxrDxCBxaDxibxiCzwFcyCqwFEyCnwFCwFBsfcwFqsfEwFnsfCsfBkrcsfqkrEsfnkrCkrBBjckrqBjEkrnBjCBjBBjqBjnyaozDfDfyyamDdzyalwEoyCfwhowEmwhmwElwhlsdowEvsvosdmsvmsdlsvlknosdvlroknmlrmknllrlBboknvDjoBbmDjmBblDjlBbvDjvzbebfwnpzzbdbdybczyaeDFzyiuyadbhzyitwEewguwEdwxuwgtwxtscustuscttvustttvtklulnukltnrulntnrtBDuDbuBDtbjuDbtbjtjfsrpyjdwrozjcyjcjzbFbFyzjhjhybEzjgzyaFyihyyxwEFwghwwxxxxschssxttxvvxkkxllxnnxrrxBBxDDxbbxjFwrmzjEyjEjbCzjazjCyjCjjBjwCowCmwClsFowCvsFmsFlkfosFvkfmkflArokfvArmArlArvyDeBpzyDdwCewauwCdwatsEushusEtshtkdukvukdtkvtAnuBruAntBrtzDpDpyDozyDFybhwCFwahwixsEhsgxsxxkcxktxlvxAlxBnxDrxbpwnuzboybojDmzbqzjpsruyjowrujjoijobbmyjqybmjjqjjmwrtjjmijmbbljjnjjlijlbjkrsCusCtkFukFtAfuAftwDhsChsaxkExkhxAdxAvxBuzDuyDujbuwnxjbuibubDtjbvjjusrxijugrxbjuajuDbtijvibtbjvbjtgrwrjtajtDbsrjtrjsqjsnBxjDxiDxbbxgnyrbxabxDDwrbxrbwqbwn"
CodageMC(2) = "pjkurwejApbsunyebkpDwulzeDspByeBwzfcfjkprwzfEfbspnyzfCfDwplzzfBfByyrczfqfrwyrEzfnfnyyrCflzyrBxjcyrqxjEyrnxjCxjBuzcxjquzExjnuzCuzBpzcuzqpzEuznpzCdjAorsufydbkonwudzdDsolydBwokzdAyzdodrsovyzdmdnwotzzdldlydkzynozdvdvyynmdtzynlxboynvxbmxblujoxbvujmujlozoujvozmozlcrkofwuFzcnsodyclwoczckyckjzcucvwohzzctctycszylucxzyltxDuxDtubuubtojuojtcfsoFycdwoEzccyccjzchchycgzykxxBxuDxcFwoCzcEycEjcazcCycCjFjAmrstfyFbkmnwtdzFDsmlyFBwmkzFAyzFoFrsmvyzFmFnwmtzzFlFlyFkzyfozFvFvyyfmFtzyflwroyfvwrmwrltjowrvtjmtjlmzotjvmzmmzlqrkvfwxpzhbAqnsvdyhDkqlwvczhBsqkyhAwqkjhAiErkmfwtFzhrkEnsmdyhnsqtymczhlwEkyhkyEkjhkjzEuEvwmhzzhuzEthvwEtyzhthtyEszhszyduExzyvuydthxzyvtwnuxruwntxrttbuvjutbtvjtmjumjtgrAqfsvFygnkqdwvEzglsqcygkwqcjgkigkbEfsmFygvsEdwmEzgtwqgzgsyEcjgsjzEhEhyzgxgxyEgzgwzycxytxwlxxnxtDxvbxmbxgfkqFwvCzgdsqEygcwqEjgcigcbEFwmCzghwEEyggyEEjggjEazgizgFsqCygEwqCjgEigEbECygayECjgajgCwqBjgCigCbEBjgDjgBigBbCrklfwspzCnsldyClwlczCkyCkjzCuCvwlhzzCtCtyCszyFuCx"
CodageMC(2) = CodageMC(2) & "zyFtwfuwftsrusrtljuljtarAnfstpyankndwtozalsncyakwncjakiakbCfslFyavsCdwlEzatwngzasyCcjasjzChChyzaxaxyCgzawzyExyhxwdxwvxsnxtrxlbxrfkvpwxuzinArdsvoyilkrcwvojiksrciikgrcbikaafknFwtmzivkadsnEyitsrgynEjiswaciisiacbisbCFwlCzahwCEyixwagyCEjiwyagjiwjCazaiziyzifArFsvmyidkrEwvmjicsrEiicgrEbicaicDaFsnCyihsaEwnCjigwrajigiaEbigbCCyaayCCjiiyaajiijiFkrCwvljiEsrCiiEgrCbiEaiEDaCwnBjiawaCiiaiaCbiabCBjaDjibjiCsrBiiCgrBbiCaiCDaBiiDiaBbiDbiBgrAriBaiBDaAriBriAqiAnBfskpyBdwkozBcyBcjBhyBgzyCxwFxsfxkrxDfklpwsuzDdsloyDcwlojDciDcbBFwkmzDhwBEyDgyBEjDgjBazDizbfAnpstuybdknowtujbcsnoibcgnobbcabcDDFslmybhsDEwlmjbgwDEibgiDEbbgbBCyDayBCjbiyDajbijrpkvuwxxjjdArosvuijckrogvubjccroajcEroDjcCbFknmwttjjhkbEsnmijgsrqinmbjggbEajgabEDjgDDCwlljbawDCijiwbaiDCbjiibabjibBBjDDjbbjjjjjFArmsvtijEkrmgvtbjEcrmajEErmDjECjEBbCsnlijasbCgnlbjagrnbjaabCDjaDDBibDiDBbjbibDbjbbjCkrlgvsrjCcrlajCErlDjCCjCBbBgnkrjDgbBajDabBDjDDDArbBrjDrjBcrkqjBErknjBCjBBbAqjBqbAnjBnjAorkfjAmjAlb"
CodageMC(2) = CodageMC(2) & "AfjAvApwkezAoyAojAqzBpskuyBowkujBoiBobAmyBqyAmjBqjDpkluwsxjDosluiDoglubDoaDoDBmwktjDqwBmiDqiBmbDqbAljBnjDrjbpAnustxiboknugtxbbocnuaboEnuDboCboBDmsltibqsDmgltbbqgnvbbqaDmDbqDBliDniBlbbriDnbbrbrukvxgxyrrucvxaruEvxDruCruBbmkntgtwrjqkbmcntajqcrvantDjqEbmCjqCbmBjqBDlglsrbngDlajrgbnaDlDjrabnDjrDBkrDlrbnrjrrrtcvwqrtEvwnrtCrtBblcnsqjncblEnsnjnErtnjnCblBjnBDkqblqDknjnqblnjnnrsovwfrsmrslbkonsfjlobkmjlmbkljllDkfbkvjlvrsersdbkejkubkdjktAeyAejAuwkhjAuiAubAdjAvjBuskxiBugkxbBuaBuDAtiBviAtbBvbDuklxgsyrDuclxaDuElxDDuCDuBBtgkwrDvglxrDvaBtDDvDAsrBtrDvrnxctyqnxEtynnxCnxBDtclwqbvcnxqlwnbvEDtCbvCDtBbvBBsqDtqBsnbvqDtnbvnvyoxzfvymvylnwotyfrxonwmrxmnwlrxlDsolwfbtoDsmjvobtmDsljvmbtljvlBsfDsvbtvjvvvyevydnwerwunwdrwtDsebsuDsdjtubstjttvyFnwFrwhDsFbshjsxAhiAhbAxgkirAxaAxDAgrAxrBxckyqBxEkynBxCBxBAwqBxqAwnBxnlyoszflymlylBwokyfDxolyvDxmBwlDxlAwfBwvDxvtzetzdlyenyulydnytBweDwuBwdbxuDwtbxttzFlyFnyhBwFDwhbwxAiqAinAyokjfAymAylAifAyvkzekzdAyeByuAydBytszp"
End Sub

Public Function PDF417String(Chain, Optional ByRef security = -1, Optional ByRef nbcol = 1, Optional ByRef CodeErr)
'Parameters :   The string to encode
'               The desired security level (default = -1)
'               The desired number of data MC columns (automatic = -1)
'               A variable which can retrieve an error number
'Return :       * a string which, printed with the PDF417.tff font, gives the barcode
'               * an empty string if the given parameters aren't good
'               * security contain used security level
'               * nbcol contain used number of data CW columns
'               * CodeErr is 0 if no error occurred, else:
'                   0  : No error
'                   1  : Chain is empty
'                   2  : Chain contains too many datas, we go beyond the 928 CWs
'                   3  : Number of CWs per row too small, we go beyond 90 rows
'                   10 : The security level been lowered not to exceed the 928 CWs (not an error, only a warning)

'Global variables
Dim i, j, K, IndexChain, Dummy, Flag As Boolean
'Splitting into blocks
Dim List(), IndexList
'Data compaction
Dim Length, ChainMC, Total
'"Text" mode processing
Dim ListT(), IndexListT, CurTabl, ChainT, NewTabl
'Reed Solomon codes
Dim MCcorrection()
'Left and Right side CWs
Dim C1, C2, C3
'Subroutine QuelMode
Dim Mode, CodeASCII
'Subroutine modulo
Dim ChainMod, Divisor, ChainMult, Number
'Tables
Call initASCII
Call initCoefRS
Call initCodageMC
CodeErr = 0
If Chain = "" Then CodeErr = 1: Exit Function
If nbcol < 1 Then nbcol = 1
'Split the string in character blocks of the same type : numeric , text, byte
'The first column of the array List contain the char. number, the second one contain the mode switch
IndexChain = 1
GoSub QuelMode
Do
  ReDim Preserve List(1, IndexList)
  List(1, IndexList) = Mode
  Do While List(1, IndexList) = Mode
    List(0, IndexList) = List(0, IndexList) + 1
    IndexChain = IndexChain + 1
    If IndexChain > Len(Chain) Then Exit Do
    GoSub QuelMode
  Loop
  IndexList = IndexList + 1
Loop Until IndexChain > Len(Chain)
'We retain "numeric" mode only if it's earning, else "text" mode or even "byte" mode
'The efficiency limits have been pre-defined according to the previous mode and/or the next mode.
For i = 0 To IndexList - 1
  If List(1, i) = 902 Then
    If i = 0 Then 'It's the first block
      If IndexList > 1 Then 'And there is other blocks behind
        If List(1, i + 1) = 900 Then
          'First block and followed by a "text" type block
          If List(0, i) < 8 Then List(1, i) = 900
          ElseIf List(1, i + 1) = 901 Then
          'First block and followed by a "byte" type block
          If List(0, i) = 1 Then List(1, i) = 901
        End If
      End If
    Else 'It's not the first block
      If i = IndexList - 1 Then
        'It's the last one
        If List(1, i - 1) = 900 Then
          'It's  preceded by a "text" type block
          If List(0, i) < 7 Then List(1, i) = 900
          ElseIf List(1, i - 1) = 901 Then
            'It's  preceded by a "byte" type block
            If List(0, i) = 1 Then List(1, i) = 901
          End If
      Else
        'It's not the last block
        If List(1, i - 1) = 901 And List(1, i + 1) = 901 Then
          'Framed by "byte" type blocks
          If List(0, i) < 4 Then List(1, i) = 901
        ElseIf List(1, i - 1) = 900 And List(1, i + 1) = 901 Then
          'Preceded by "text" and followed by "byte" (If the reverse it's never interesting to change)
          If List(0, i) < 5 Then List(1, i) = 900
        ElseIf List(1, i - 1) = 900 And List(1, i + 1) = 900 Then
          'Framed by "text" type blocks
          If List(0, i) < 8 Then List(1, i) = 900
        End If
      End If
    End If
  End If
Next
GoSub Regroupe
'Maintain "text" mode only if it's earning
For i = 0 To IndexList - 1
  If List(1, i) = 900 And i > 0 Then
    'It's not the first (If first, never interesting to change)
    If i = IndexList - 1 Then 'C'est le dernier / It's the last one
      If List(1, i - 1) = 901 Then
        'It's  preceded by a "byte" type block
        If List(0, i) = 1 Then List(1, i) = 901
      End If
    Else
      'It's not the last one
      If List(1, i - 1) = 901 And List(1, i + 1) = 901 Then
        'Framed by "byte" type blocks
        If List(0, i) < 5 Then List(1, i) = 901
      ElseIf (List(1, i - 1) = 901 And List(1, i + 1) <> 901) Or (List(1, i - 1) <> 901 And List(1, i + 1) = 901) Then
        'A "byte" block ahead or behind
        If List(0, i) < 3 Then List(1, i) = 901
      End If
    End If
  End If
Next
GoSub Regroupe
'Now we compress datas into the MCs, the MCs are stored in 3 char. in a large string : ChainMC
IndexChain = 1
For i = 0 To IndexList - 1
  'Thus 3 compaction modes
  Select Case List(1, i)
    Case 900 'Text
      ReDim ListT(1, List(0, i))
      'ListT will contain the table number(s) (1 ou several) and the value of each char.
      'Table number encoded in the 4 less weight bits, that is in decimal 1, 2, 4, 8
      For IndexListT = 0 To List(0, i) - 1
        CodeASCII = Asc(Mid(Chain, IndexChain + IndexListT, 1))
        Select Case CodeASCII
          Case 9 'HT
            ListT(0, IndexListT) = 12
            ListT(1, IndexListT) = 12
          Case 10 'LF
            ListT(0, IndexListT) = 8
            ListT(1, IndexListT) = 15
          Case 13 'CR
            ListT(0, IndexListT) = 12
            ListT(1, IndexListT) = 11
          Case Else
            ListT(0, IndexListT) = Mid(ASCII, CodeASCII * 4 - 127, 2)
            ListT(1, IndexListT) = Mid(ASCII, CodeASCII * 4 - 125, 2)
          End Select
      Next
      CurTabl = 1 'Default table
      ChainT = ""
      'Datas are stored in 2 char. in the string TableT
      For j = 0 To List(0, i) - 1
        If (ListT(0, j) And CurTabl) > 0 Then
          'The char. is in the current table
          ChainT = ChainT & Format(ListT(1, j), "00")
        Else
          'Obliged to change the table
          Flag = False 'True if we change the table only for 1 char.
          If j = List(0, i) - 1 Then
            Flag = True
          Else
            If (ListT(0, j) And ListT(0, j + 1)) = 0 Then Flag = True 'No common table with the next char.
          End If
          If Flag Then
            'We change only for 1 char., Look for a temporary switch
            If (ListT(0, j) And 1) > 0 And CurTabl = 2 Then
              'Table 2 to 1 for 1 char. --> T_UPP
              ChainT = ChainT & "27" & Format(ListT(1, j), "00")
            ElseIf (ListT(0, j) And 8) > 0 Then
              'Table 1 or 2 or 4 to table 8 for 1 char. --> T_PUN
              ChainT = ChainT & "29" & Format(ListT(1, j), "00")
            Else
              'No temporary switch available
              Flag = False
            End If
          End If
          If Not Flag Then 'We test again flag which is perhaps changed ! Impossible to use ELSE statement
            '
            'We must use a bi-state switch
            'Looking for the new table to use
            If j = List(0, i) - 1 Then
              NewTabl = ListT(0, j)
            Else
              NewTabl = IIf((ListT(0, j) And ListT(0, j + 1)) = 0, ListT(0, j), ListT(0, j) And ListT(0, j + 1))
            End If
            'Maintain the first if several tables are possible
            Select Case NewTabl
            Case 3, 5, 7, 9, 11, 13, 15
              NewTabl = 1
            Case 6, 10, 14
              NewTabl = 2
            Case 12
              NewTabl = 4
            End Select
            'Select the switch, on occasion we must use 2 switchs consecutively
            Select Case CurTabl
            Case 1
              Select Case NewTabl
              Case 2
                ChainT = ChainT & "27"
              Case 4
                ChainT = ChainT & "28"
              Case 8
                ChainT = ChainT & "2825"
              End Select
            Case 2
              Select Case NewTabl
              Case 1
                ChainT = ChainT & "2828"
              Case 4
                ChainT = ChainT & "28"
              Case 8
                ChainT = ChainT & "2825"
              End Select
            Case 4
              Select Case NewTabl
              Case 1
                ChainT = ChainT & "28"
              Case 2
                ChainT = ChainT & "27"
              Case 8
                ChainT = ChainT & "25"
              End Select
            Case 8
              Select Case NewTabl
              Case 1
                ChainT = ChainT & "29"
              Case 2
                ChainT = ChainT & "2927"
              Case 4
                ChainT = ChainT & "2928"
              End Select
            End Select
            CurTabl = NewTabl
            ChainT = ChainT & Format(ListT(1, j), "00") 'At last we add the char.
          End If
        End If
      Next
      If Len(ChainT) Mod 4 > 0 Then ChainT = ChainT & "29" 'Padding if number of char. is odd
      'Now translate the string ChainT into CWs
      If i > 0 Then ChainMC = ChainMC & "900" 'Set up the switch exept for the first block because "text" is the default
      For j = 1 To Len(ChainT) Step 4
        ChainMC = ChainMC & Format(Mid(ChainT, j, 2) * 30 + Mid(ChainT, j + 2, 2), "000")
      Next
    Case 901 'Octet
      'Select the switch between the 3 possible
      If List(0, i) = 1 Then
        '1 seul octet, c'est immédiat
        ChainMC = ChainMC & "913" & Format(Asc(Mid(Chain, IndexChain, 1)), "000")
      Else
        'Select the switch for perfect multiple of 6 bytes or no
        If List(0, i) Mod 6 = 0 Then
          ChainMC = ChainMC & "924"
        Else
          ChainMC = ChainMC & "901"
        End If
        j = 0
        Do While j < List(0, i)
          Length = List(0, i) - j
          If Length >= 6 Then
            'Take groups of 6
            Length = 6
            Total = 0
            For K = 0 To Length - 1
              Total = Total + (Asc(Mid(Chain, IndexChain + j + K, 1)) * 256 ^ (Length - 1 - K))
            Next
            ChainMod = Format(Total, "general number")
            Dummy = ""
            Do
              Divisor = 900
              GoSub Modulo
              Dummy = Format(Divisor, "000") & Dummy
              ChainMod = ChainMult
              If ChainMult = "" Then Exit Do
            Loop
            ChainMC = ChainMC & Dummy
          Else
            'If it remain a group of less than 6 bytes
            For K = 0 To Length - 1
              ChainMC = ChainMC & Format(Asc(Mid(Chain, IndexChain + j + K, 1)), "000")
            Next
          End If
          j = j + Length
        Loop
      End If
    Case 902 'Numeric
      ChainMC = ChainMC & "902"
      j = 0
      Do While j < List(0, i)
        Length = List(0, i) - j
        If Length > 44 Then Length = 44
        ChainMod = "1" & Mid(Chain, IndexChain + j, Length)
        Dummy = ""
        Do
          Divisor = 900
          GoSub Modulo
          Dummy = Format(Divisor, "000") & Dummy
          ChainMod = ChainMult
          If ChainMult = "" Then Exit Do
        Loop
        ChainMC = ChainMC & Dummy
        j = j + Length
      Loop
      'Debug.Print ChainMC
    End Select
    IndexChain = IndexChain + List(0, i)
  Next
  'ChainMC contain the MC list (on 3 digits) depicting the datas
  'Now we take care of the correction level
  Length = Len(ChainMC) / 3
  If security < 0 Then
    'Fixing auto. the correction level according to the standard recommendations
    If Length < 41 Then
      security = 2
    ElseIf Length < 161 Then
      security = 3
    ElseIf Length < 321 Then
      security = 4
    Else
      security = 5
    End If
  End If
  'Now we take care of the number of CW per row
  Length = Length + 1 + (2 ^ (security + 1))
  If nbcol > 30 Then nbcol = 30
  If nbcol < 1 Then
    'With a 3 modules high font, for getting a "square" bar code
    'x = nb. of col. | Width by module = 69 + 17x | Height by module = 3t / x (t is the total number of MCs)
    'Thus we have 69 + 17x = 3t/x <=> 17x²+69x-3t=0 - Discriminant is 69²-4*17*-3t = 4761+204t thus x=SQR(discr.)-69/2*17
    nbcol = (Sqr(204 * Length + 4761) - 69) / (34 / 1.3)   '1.3 = balancing factor determined at a guess after tests
    If nbcol = 0 Then nbcol = 1
  End If
  'If we go beyond 928 CWs we try to reduce the correction level
  Do While security > 0
    'Calculation of the total number of CW with the padding
    Length = Len(ChainMC) / 3 + 1 + (2 ^ (security + 1))
    Length = (Length \ nbcol + IIf(Length Mod nbcol > 0, 1, 0)) * nbcol
    If Length < 929 Then Exit Do
    'We must reduce security level
    security = security - 1
    CodeErr = 10
  Loop
  If Length > 928 Then CodeErr = 2: Exit Function
  If Length / nbcol > 90 Then CodeErr = 3: Exit Function
  'Padding calculation
  Length = Len(ChainMC) / 3 + 1 + (2 ^ (security + 1))
  i = 0
  If Length / nbcol < 3 Then
    i = nbcol * 3 - Length   'A bar code must have at least 3 row
  Else
    If Length Mod nbcol > 0 Then i = nbcol - (Length Mod nbcol)
  End If
  'We add the padding
  Do While i > 0
    ChainMC = ChainMC & "900"
    i = i - 1
  Loop
  'We add the length descriptor
  ChainMC = Format(Len(ChainMC) / 3 + 1, "000") & ChainMC
  'Now we take care of the Reed Solomon codes
  Length = Len(ChainMC) / 3
  K = 2 ^ (security + 1)
  ReDim MCcorrection(K - 1)
  Total = 0
  For i = 0 To Length - 1
    Total = (Mid(ChainMC, i * 3 + 1, 3) + MCcorrection(K - 1)) Mod 929
    For j = K - 1 To 0 Step -1
      If j = 0 Then
        MCcorrection(j) = (929 - (Total * Mid(CoefRS(security), j * 3 + 1, 3)) Mod 929) Mod 929
      Else
        MCcorrection(j) = (MCcorrection(j - 1) + 929 - (Total * Mid(CoefRS(security), j * 3 + 1, 3)) Mod 929) Mod 929
      End If
    Next
  Next
  For j = 0 To K - 1
    If MCcorrection(j) <> 0 Then MCcorrection(j) = 929 - MCcorrection(j)
  Next
  'We add theses codes to the string
  For i = K - 1 To 0 Step -1
    ChainMC = ChainMC & Format(MCcorrection(i), "000")
  Next
  'The CW string is finished
  'Calculation of parameters for the left and right side CWs
  C1 = (Len(ChainMC) / 3 / nbcol - 1) \ 3
  C2 = security * 3 + (Len(ChainMC) / 3 / nbcol - 1) Mod 3
  C3 = nbcol - 1
  'We encode each row
  For i = 0 To Len(ChainMC) / 3 / nbcol - 1
    Dummy = Mid(ChainMC, i * nbcol * 3 + 1, nbcol * 3)
    K = (i \ 3) * 30
    Select Case i Mod 3
    Case 0
      Dummy = Format(K + C1, "000") & Dummy & Format(K + C3, "000")
    Case 1
      Dummy = Format(K + C2, "000") & Dummy & Format(K + C1, "000")
    Case 2
      Dummy = Format(K + C3, "000") & Dummy & Format(K + C2, "000")
    End Select
    PDF417String = PDF417String & "+*" 'Start with a start char. and a separator
    For j = 0 To Len(Dummy) / 3 - 1
      PDF417String = PDF417String & Mid(CodageMC(i Mod 3), Mid(Dummy, j * 3 + 1, 3) * 3 + 1, 3) & "*"
    Next
    PDF417String = PDF417String & "-" & Chr(13) & Chr(10) 'Add a stop char. and a CRLF
  Next
  Exit Function
Regroupe:
  'Bring together same type blocks
  If IndexList > 1 Then
    i = 1
    Do While i < IndexList
      If List(1, i - 1) = List(1, i) Then
        'Bringing together
        List(0, i - 1) = List(0, i - 1) + List(0, i)
        j = i + 1
        'Decrease the list
        Do While j < IndexList
          List(0, j - 1) = List(0, j)
          List(1, j - 1) = List(1, j)
          j = j + 1
        Loop
        IndexList = IndexList - 1
        i = i - 1
      End If
      i = i + 1
    Loop
  End If
Return
QuelMode:
  CodeASCII = Asc(Mid(Chain, IndexChain, 1))
  Select Case CodeASCII
  Case 48 To 57
    Mode = 902
  Case 9, 10, 13, 32 To 126
    Mode = 900
  Case Else
    Mode = 901
  End Select
Return
Modulo:
  'ChainMod depict a very large number having more than 9 digits
  'Divisor is the divisor, contain the result after return
  'ChainMult contain after return the result of the integer division
  ChainMult = ""
  Number = 0
  Do While ChainMod <> ""
    Number = Number * 10 + Left(ChainMod, 1) 'Put down a digit
    ChainMod = Mid(ChainMod, 2)
    If Number < Divisor Then
      If ChainMult <> "" Then ChainMult = ChainMult & "0"
    Else
      ChainMult = ChainMult & Number \ Divisor
    End If
    Number = Number Mod Divisor 'Get the remainder
  Loop
  Divisor = Number
Return
End Function

Public Function PDF417ToBinary(strPDF417 As String)
'Define variables
Dim LCaseAlpha(25) As String, UCaseAlpha(5) As String 'Data words and left/right indicators
Dim SpecialChar(3) As String
Dim BinPDF417() As String
Dim i As Long, j As Long
Dim BinRow As String
'Define array values
LCaseAlpha(0) = "00110" 'a
LCaseAlpha(1) = "00111" 'b
LCaseAlpha(2) = "01000" 'c...
LCaseAlpha(3) = "01001"
LCaseAlpha(4) = "01010"
LCaseAlpha(5) = "01011"
LCaseAlpha(6) = "01100"
LCaseAlpha(7) = "01101"
LCaseAlpha(8) = "01110"
LCaseAlpha(9) = "01111"
LCaseAlpha(10) = "10000"
LCaseAlpha(11) = "10001"
LCaseAlpha(12) = "10010"
LCaseAlpha(13) = "10011"
LCaseAlpha(14) = "10100"
LCaseAlpha(15) = "10101"
LCaseAlpha(16) = "10110"
LCaseAlpha(17) = "10111"
LCaseAlpha(18) = "11000"
LCaseAlpha(19) = "11001"
LCaseAlpha(20) = "11010"
LCaseAlpha(21) = "11011"
LCaseAlpha(22) = "11100"
LCaseAlpha(23) = "11101"
LCaseAlpha(24) = "11110"
LCaseAlpha(25) = "11111"
UCaseAlpha(0) = "00000" 'A
UCaseAlpha(1) = "00001" 'B
UCaseAlpha(2) = "00010" 'C...
UCaseAlpha(3) = "00011"
UCaseAlpha(4) = "00100"
UCaseAlpha(5) = "00101"
SpecialChar(0) = "01" '* (separator)
SpecialChar(1) = "1111111101010100" '+ (start pattern)
SpecialChar(3) = "11111101000101001" '- (end pattern)
'Split strPDF417 by each row (remove Chr(13) and Chr(10) off the end)
BinPDF417 = Split(Left(strPDF417, Len(strPDF417) - 2), Chr(13) & Chr(10))
'Loop through each character and convert to binary
For i = 0 To UBound(BinPDF417)
    BinRow = ""
    For j = 1 To Len(BinPDF417(i))
        Select Case Asc(Mid(BinPDF417(i), j, 1))
            Case 42 To 45 'Special characters (*, +, -)
                BinRow = BinRow & SpecialChar(Asc(Mid(BinPDF417(i), j, 1)) - 42)
            Case 65 To 70 'A-F
                BinRow = BinRow & UCaseAlpha(Asc(Mid(BinPDF417(i), j, 1)) - 65)
            Case 97 To 122 'a-z
                BinRow = BinRow & LCaseAlpha(Asc(Mid(BinPDF417(i), j, 1)) - 97)
            Case Else
                PDF417ToBinary = "ERROR: Character not found (ASCII=" & Asc(Mid(BinPDF417(i), j, 1)) & ")"
                Exit Function
        End Select
    Next j
    BinPDF417(i) = BinRow
Next i
' Return PDF417 as binary
PDF417ToBinary = BinPDF417
End Function
