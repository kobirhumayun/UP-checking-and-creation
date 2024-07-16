Attribute VB_Name = "vs_code_not_supported_text"
Option Explicit

'vs-code not supported text as Dictionary write in VBA environment
'then export, also noted this module do not modify from vs-code
'if needed modify from VBA environment only then export

Private Function CreateVsCodeNotSupportedOrBengaliTxtDictionary()

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = CreateObject("Scripting.Dictionary")
    
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "sourceDataAsDicUpIssuingStatusCurrencyNumberFormat", "_([$�-2]"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "expNoAndDtBengaliTxt", "BGK&�wc bs I ZvwiL"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "mlcNoAndDtBengaliTxt", "gv�vi Gj wm bs I ZvwiL"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "localB2bLcDesBengaliTxt", "7| �jvKvj e�vK Uz e�vK Gj/wm Gi weeiY t"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "infoAboutStockBengaliTxt", "9| gRy` m��wK�Z Z_�vw`"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "totalUsedRawMetarialsBengaliTxt", "�gvU e�e�Z c�Y�i"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "charCode151", "�"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "charCode151WithSlash", "\�"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "pioneerDenimLimitedUpNoBengaliTxt", "cvBIwbqvi �Wwbg wjwg�UW. BD wc bs-"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "exportLcSalesContractBengaliTxt", "i�vwb FYc�/�mjm K�U�v� t"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "upAppNoPart1BengaliTxt", "�gmvm�  cvBIwbqvi �Wwbgm wjwg�UW KZ��K `vwLjK�Z BDwc Av�e`b bs-"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "upAppNoPart2BengaliTxt", " Gi Z_�vw` wb�g� Dc��vcb Kiv n�jv t"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "bbLcScNoAndDtBengaliTxt", "weweGjwm/�mjm K�U�v� bs I ZvwiL"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "udIpExpNoAndDtBengaliTxt", "BDwW/AvBwc b�^i I ZvwiL ms�kvabxmn/BGK&�wc bs I ZvwiL"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "denimFabricsBengaliTxt", "�Wwbg Kvco"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "issuingBankNameAndAddressBengaliTxt", "Bm�yK�Z e�vs�Ki bvg I wVKvbv"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "submittedInfoBengaliTxt", "c�wZ�v�bi `vwLjK�Z Z_�"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "foundCorrectBengaliTxt", "mwVK cvIqv wM�q�Q"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "buyerNameBengaliTxt", "��Zv c�wZ�v�bi bvg"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "udNoBengaliTxt", "BDwW bs-"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "mLcExpIpNoBengaliTxt", "gv�vi Gjwm/ AvBwc/BG�wc bs I ZvwiL"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "sellerNameBengaliTxt", "we��Zv c�wZ�v�bi bvg"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "bbLcValueBengaliTxt", "weweGj wm bs I g~j� (gv:W:)"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "qtyOfGoodsYdsBengaliTxt", "c�b�I cwigvb (MR)"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "mLcValueBengaliTxt", "gv�vi Gj wm g~j�"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "mLcValidityBengaliTxt", "gv�vi Gj wm �gqv`"
    vsCodeNotSupportedOrBengaliTxtDictionary.Add "rawMaterialNameAndDescriptionBengaliTxt", "KuvPvgv�ji bvg I eY�bv"
    
    Set CreateVsCodeNotSupportedOrBengaliTxtDictionary = vsCodeNotSupportedOrBengaliTxtDictionary
    
End Function
