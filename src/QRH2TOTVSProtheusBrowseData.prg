#include "minigui.ch"
#include "tsbrowse.ch"

procedure QRH2TOTVSProtheusBrowseData(oRecordSet,cTitle)

    local cCodePage as character

    local oQRH2TotvsBrowseData
    local Font_QRH2TotvsBrowseData
    local Form_QRH2TOTVSBrwoseData

    local nWinWidth as numeric  := getdesktopwidth()
    local nWinHeight as numeric := getdesktopheight()
    local nBrwWidth as numeric := nWinWidth-30
    local nBrwHeight as numeric := nWinHeight-60

    IF (!_IsControlDefined ("Font_QRH2TotvsBrowseData","Main"))
        DEFINE FONT Font_QRH2TotvsBrowseData FONTNAME "Arial" SIZE 10
    endif

    DEFINE WINDOW Form_QRH2TOTVSBrwoseData AT 0,0 ;
    WIDTH nWinWidth HEIGHT nWinHeight ;
    TITLE cTitle;
    ICON GetStartupFolder()+"\QRH2TOTVSProtheus.ico";
    CHILD;
    NOMAXIMIZE NOSIZE
    ON INIT  oQRH2TotvsBrowseData:SetFocus()

     @  10,  10 TBROWSE oQRH2TotvsBrowseData RECORDSET oRecordSet  EDITABLE AUTOCOLS SELECTOR .T. ;
        WIDTH nBrwWidth HEIGHT nBrwHeight  ;
        FONT Font_QRH2TotvsBrowseData ;
        COLORS CLR_BLACK, CLR_WHITE, CLR_BLACK, { CLR_WHITE, COLOR_GRID }, CLR_BLACK, -CLR_HRED  ;

    oQRH2TotvsBrowseData:aColumns[ 1 ]:lEdit := .F.
    oQRH2TotvsBrowseData:nClrLine := COLOR_GRID
    
    if (oQRH2TotvsBrowseData:lDrawSpecHd)
        oQRH2TotvsBrowseData:nClrSpcHdBack := oQRH2TotvsBrowseData:nClrHeadBack
    endif
    
   ON KEY ESCAPE ACTION ThisWindow.Release

   END WINDOW

    cCodePage:=Hb_SetCodePage('PTISO')

   ACTIVATE WINDOW Form_QRH2TOTVSBrwoseData

   RELEASE FONT Font_QRH2TotvsBrowseData

   Hb_SetCodePage(cCodePage)
   
return
