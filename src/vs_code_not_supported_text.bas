Attribute VB_Name = "vs_code_not_supported_text"
Option Explicit

'vs-code not supported text as Dictionary write in VBA environment
'then export, also noted this module do not modify from vs-code
'if needed modify from VBA environment only then export

Private Function CreateVsCodeNotSupportedOrBengaliTxtDictionary()

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = CreateObject("Scripting.Dictionary")
    
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "sourceDataAsDicUpIssuingStatusCurrencyNumberFormat", "_([$€-2]"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "expNoAndDtBengaliTxt", "BGK&ªwc bs I ZvwiL"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "mlcNoAndDtBengaliTxt", "gvóvi Gj wm bs I ZvwiL"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "localB2bLcDesBengaliTxt", "7| ‡jvKvj e¨vK Uz e¨vK Gj/wm Gi weeiY t"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "infoAboutStockBengaliTxt", "9| gRy` m¤úwK©Z Z_¨vw`"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "totalUsedRawMetarialsBengaliTxt", "‡gvU e¨eüZ c‡Y¨i"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "charCode151", "—"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "charCode151WithSlash", "\—"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "pioneerDenimLimitedUpNoBengaliTxt", "cvBIwbqvi †Wwbg wjwg‡UW. BD wc bs-"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "exportLcSalesContractBengaliTxt", "ißvwb FYcÎ/†mjm K›Uªv± t"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "upAppNoPart1BengaliTxt", "†gmvm©  cvBIwbqvi ‡Wwbgm wjwg‡UW KZ©„K `vwLjK…Z BDwc Av‡e`b bs-"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "upAppNoPart2BengaliTxt", " Gi Z_¨vw` wb‡gœ Dc¯’vcb Kiv n‡jv t"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "bbLcScNoAndDtBengaliTxt", "weweGjwm/†mjm K›Uªv± bs I ZvwiL"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "udIpExpNoAndDtBengaliTxt", "BDwW/AvBwc b¤^i I ZvwiL ms‡kvabxmn/BGK&ªwc bs I ZvwiL"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "denimFabricsBengaliTxt", "‡Wwbg Kvco"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "issuingBankNameAndAddressBengaliTxt", "Bm¨yK…Z e¨vs‡Ki bvg I wVKvbv"
    
    Set CreateVsCodeNotSupportedOrBengaliTxtDictionary = vsCodeNotSupportedOrBengaliTxtDictionary
    
End Function
