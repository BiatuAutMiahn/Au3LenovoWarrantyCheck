#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Res\ohiohealth.ico
#AutoIt3Wrapper_Outfile_x64=..\WarrantyLookup.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=Quickly check Lenovo Warranty
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_ProductVersion=1.0.0.0
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;Lenovo Warranty Check
#include-once
#include <Array.au3>
#include <Date.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiStatusBar.au3>
#include <WindowsConstants.au3>
#include <GuiEdit.au3>
#include "Includes\Common\json.au3"
Global Const $HTTP_STATUS_OK = 200
Global Const $sTitle="Lenovo Warranty Lookup"
Global $bStatTimer,$iStatTimer
#Region ### START Koda GUI section ### Form=
$hWnd = GUICreate($sTitle, 370, 57, 192, 124)
GUISetFont(10, 400, 0, "Consolas")
$idInput = GUICtrlCreateInput("", 8, 8, 353, 23)
_GUICtrlEdit_SetCueBanner($idInput, "Serial or Asset Number, then Enter.", True)
$idHotKey = GUICtrlCreateDummy()
Dim $AccelKeys[1][2] = [["{ENTER}", $idHotKey]]; Set accelerators
GUISetAccelerators($AccelKeys)
$idStatus = _GUICtrlStatusBar_Create($hWnd)
_GUICtrlStatusBar_SetText($idStatus,"Initializing...")
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###
_GUICtrlStatusBar_SetText($idStatus,"Ready")
While Sleep(1)
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
        Case $idHotKey
            ;ConsoleWrite(_GuiCtrlGetFocus($hWnd)&"=="&$idHost&@CRLF)
            If _GuiCtrlGetFocus($hWnd)<>$idInput Then ContinueLoop
            _DoCheck(GUICtrlRead($idInput))
            GUICtrlSetData($idInput,"")
    EndSwitch
    If $bStatTimer And TimerDiff($iStatTimer)>=10000 Then
        $bStatTimer=False
        _GUICtrlStatusBar_SetText($idStatus,"Ready")
    EndIf
WEnd

Func _SetResultStat($sMsg)
    _GUICtrlStatusBar_SetText($idStatus,$sMsg)
    $iStatTimer=TimerInit()
    $bStatTimer=True
EndFunc

Func _InvalidSN()
    SoundPlay(@WindowsDir&"\Media\Windows Critical Stop.wav")
    _SetResultStat("Error: Invalid serial number, or unknown format.")
EndFunc

Func _FailLookup()
    SoundPlay(@WindowsDir&"\Media\Windows Critical Stop.wav")
    _SetResultStat("Failed to lookup warranty")
EndFunc

Func _DoCheck($sSN)
    $sSN=StringLower($sSN)
    If StringLeft($sSN,2)=="oh" Then
        _GUICtrlStatusBar_SetText($idStatus,"Getting serial from ServiceNow...")
        $sSerial=_snGetHwSerial($sSN)
        If @error Then Return _InvalidSN()
    Else
        _GUICtrlStatusBar_SetText($idStatus,"Getting warranty info...")
        $ret=HttpGet("https://pcsupport.lenovo.com/us/en/api/v4/mse/getproducts","productId="&$sSN)
        $oJson = Json_Decode($ret)
        If @error Then Return _InvalidSN()
        If Not IsArray($oJson) Then Return _InvalidSN()
        If UBound($oJson,1)<>1 Then Return _InvalidSN()
        If Not IsObj($oJson[0]) Then Return _InvalidSN()
        $sSerial=Json_ObjGet($oJson[0],"Serial")
    EndIf
    ;ConsoleWrite($sSerial&@CRLF)
    ;ConsoleWrite(Json_Encode($oJson[0], $JSON_PRETTY_PRINT + $JSON_UNESCAPED_SLASHES, "   ")&@CRLF)
    ;Exit
    _GUICtrlStatusBar_SetText($idStatus,"Validating Serial with Lenovo...")
    $ret=HttpPost("https://pcsupport.lenovo.com/us/en/api/v4/upsell/redport/getIbaseInfo",'{"serialNumber":"'&$sSerial&'","country":"us","language":"en"}')
    $oJson = Json_Decode($ret)
    If @error Then Return _InvalidSN()
    ;ConsoleWrite(Json_Encode($oJson, $JSON_PRETTY_PRINT + $JSON_UNESCAPED_SLASHES, "   ")&@CRLF)
    If Not Json_ObjExists($oJson,"data") Then Return _FailLookup()
    $oData=Json_ObjGet($oJson,"data")
    If Not IsObj($oData) Then Return _FailLookup()
    If Not Json_ObjExists($oData,"currentWarranty") Then Return _FailLookup()
    $oCW=Json_ObjGet($oData,"currentWarranty")
    If Not Json_ObjExists($oCW,"endDate") Then Return _FailLookup()
    ;ConsoleWrite(&@CRLF)
    $sExpDate=StringReplace(Json_ObjGet($oCW, "endDate"),'-','.')
    $sNow=_NowCalcDate()
    $iDays=_DateDiff('D',$sNow,$sExpDate)
    If $iDays<0 Then
        $sMsg="Warranty expired on "&$sExpDate
        SoundPlay(@WindowsDir&"\Media\Windows Hardware Fail.wav")
    Else
        $sMsg="Warranty expires on "&$sExpDate
        SoundPlay(@WindowsDir&"\Media\Windows Foreground.wav")
    EndIf
    _SetResultStat($sMsg)
EndFunc

Func HttpPost($sURL, $sData = "")
    Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")
    $oHTTP.Open("POST", $sURL, False)
    If (@error) Then Return SetError(1, 0, 0)
    $oHTTP.SetRequestHeader("Content-Type", "application/json")
    $oHTTP.Send($sData)
    If (@error) Then Return SetError(2, 0, 0)
    If ($oHTTP.Status <> $HTTP_STATUS_OK) Then Return SetError(3, 0, 0)
    Return SetError(0, 0, $oHTTP.ResponseText)
EndFunc

Func HttpGet($sURL, $sData = "")
    Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")
    $oHTTP.Open("GET", $sURL & "?" & $sData, False)
    If (@error) Then Return SetError(1, 0, 0)
    $oHTTP.Send()
    If (@error) Then Return SetError(2, 0, 0)
    If ($oHTTP.Status <> $HTTP_STATUS_OK) Then Return SetError(3, 0, 0)
    Return SetError(0, 0, $oHTTP.ResponseText)
EndFunc

Func _GuiCtrlGetFocus($GuiRef)
    Local $hwnd = ControlGetHandle($GuiRef, "", ControlGetFocus($GuiRef))
    Local $result = DllCall("user32.dll", "int", "GetDlgCtrlID", "hwnd", $hwnd)
    Return $result[0]
EndFunc

Func _snGetHwSerial($sAsset)
    Local $sRet
    $oQuery=_queryHardware("asset_tag="&$sAsset)
    $aRet=_SoapGetRecordsArray($oQuery)
    Return _SoapGetAttr($oQuery,"serial_number")
EndFunc

Func _queryHardware($sQuery)
    $SoapMsg = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:inc="http://<redacted>.service-now.com/alm_hardware_list">' & @CRLF _
                    & '   <soapenv:Header/>' & @CRLF _
                    & '   <soapenv:Body>' & @CRLF _
                    & '	  <inc:getRecords>' & @CRLF _
                    & '		<__encoded_query>'&$sQuery&'</__encoded_query>' & @CRLF _
                    & '	  </inc:getRecords>' & @CRLF _
                    & '   </soapenv:Body>' & @CRLF _
                    & '</soapenv:Envelope>'
    $sQuery=_SoapQuery("alm_hardware_list",$SoapMsg)
    $oXml = ObjCreate("Msxml2.DOMDocument.3.0")
    $oXml.loadXML($sQuery)
    $oEnvelope=_SoapGetNode($oXml,"SOAP-ENV:Envelope")
    $oBody=_SoapGetNode($oEnvelope,"SOAP-ENV:Body")
    $oRecordsResponse=_SoapGetNode($oBody,"getRecordsResponse")
    $oRecordsResults=_SoapGetNode($oRecordsResponse,"getRecordsResult")
    Return $oRecordsResults
EndFunc


Func _SoapQuery($Database, $SoapMsg, $Environment="", $User="<redacted>", $Password="<redacted>")
	$objHTTP = ObjCreate("Microsoft.XMLHTTP")
	$objReturn = ObjCreate("Msxml2.DOMDocument.3.0")
	Select
		Case $Environment = "Cloud"
			$objHTTP.open ("post", "https://<redacted>.service-now.com/"& $Database & ".do?SOAP&displayvalue=true", False, $User, $Password)
		Case $Environment = "Test"
			$objHTTP.open ("post", "https://<redacted>.service-now.com/"& $Database & ".do?SOAP&displayvalue=true", False, $User, $Password)
		Case $Environment = "Live"
			$objHTTP.open ("post", "https://<redacted>.service-now.com/"& $Database & ".do?SOAP&displayvalue=true", False, $User, $Password)
		Case Else
			$objHTTP.open ("post", "https://<redacted>.service-now.com/"& $Database & ".do?SOAP&displayvalue=true", False, $User, $Password)
	EndSelect
		If @error Then Return(-1)
	$objHTTP.setRequestHeader ("Content-Type", "text/xml")
		If @error Then Return(-2)
	$objHTTP.send ($SoapMsg)
		If @error Then Return(-3)
	$strReturn = $objHTTP.responseText
		If @error Then Return(-4)
	$objReturn.loadXML ($strReturn)
		If @error Then Return(-5)
	$Soap = $objReturn.XML
		If @error Then Return(-6)
	Return($Soap)
EndFunc

Func _SoapGetAttr($vObj,$sValue="")
    If Not IsObj($vObj) Then Return SetError(1,0,False)
    If $sValue="" Then
        Local $aRet[]=[0]
        For $oNode In $vObj.childnodes
            $aRet[0]=UBound($aRet,1)
            ReDim $aRet[$aRet[0]+1]
            $aRet[$aRet[0]]=$oNode.nodename
        Next
        Return $aRet
    EndIf
    For $oNode In $vObj.childnodes
        If $oNode.nodename<>$sValue Then ContinueLoop
        Return $oNode.text
    Next
EndFunc

Func _SoapGetRecordsArray($vObj,$bIsArray=False)
    Local $iErr=0, $iExt=0
    If IsArray($vObj) Then
        Local $aRet[1][1]
        $aRet[0][0]=0
        For $i=1 To $vObj[0]
            $aRes=_SoapGetRecordsArray($vObj[$i],True)
            If @error Then
                $iErr=1
                $iExt+=1
                ContinueLoop
            EndIf
            If UBound($aRet,2)<>@extended Then ReDim $aRet[UBound($aRet,1)][@extended]
            $aRet[$aRet[0][0]][0]=UBound($aRet,1)
            ReDim $aRet[$aRet[0][0]+1][@extended]
            For $i=0 To UBound($aRes,1)-1
                $aRet[$aRet[0][0]][$i]=$aRes[$i]
            Next
        Next
        Return SetError($iErr,$iExt,$aRet)
    EndIf
    If Not IsObj($vObj) Then Return SetError(1,0,False)
    $iAttrMax=_SoapGetAttrCount($vObj)
    Local $aRet[$iAttrMax]
    $iAttr=0
    For $oNode In $vObj.childnodes
        $aRet[$iAttr]=$oNode.text
        $iAttr+=1
    Next
    Return SetError(0,$iAttrMax,$aRet)
EndFunc

Func _SoapGetAttrCount($vObj)
    If Not IsObj($vObj) Then Return SetError(1,0,False)
    Local $iRet=0
    For $oNode In $vObj.childnodes
        $iRet+=1
    Next
    Return $iRet
EndFunc

Func _SoapGetNode($oObj,$sName)
    Local $oaRet[1]=[0]
    $iMax=0
    If Not IsObj($oObj) Then Return SetError(1,0,False)
    For $oI In $oObj.childNodes
        If $oI.nodename<>$sName Then ContinueLoop
        $oaRet[0]=UBound($oaRet,1)
        ReDim $oaRet[$oaRet[0]+1]
        $oaRet[$oaRet[0]]=$oI
    Next
    If $oaRet[0]==0 Then
        Return SetError(2,0,False)
    ElseIf $oaRet[0]>1 Then
        Return $oaRet
    Else
        Return $oaRet[1]
    EndIf
EndFunc
