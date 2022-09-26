procedure QRH2TOTVSProtheusBrowseData(oRecordSet,cTitle,lExcel,cExecTitle,bExec)

    local cCodePage as character

    local bGetFil as codeblock
    local bGetMat as codeblock

    local oQRH2TotvsBrowseData
    local Font_QRH2TotvsBrowseData
    local Form_QRH2TOTVSBrwoseData

    local nWinWidth as numeric  := getdesktopwidth()
    local nWinHeight as numeric := getdesktopheight()
    local nBrwWidth as numeric := nWinWidth-30
    local nBrwHeight as numeric := nWinHeight-60

    hb_default(@lExcel,.F.)

    if (!_IsControlDefined ("Font_QRH2TotvsBrowseData","Main"))
        DEFINE FONT Font_QRH2TotvsBrowseData FONTNAME "Arial" SIZE 10
    endif

    DEFINE WINDOW Form_QRH2TOTVSBrwoseData AT 0,0 ;
    WIDTH nWinWidth HEIGHT nWinHeight ;
    TITLE cTitle;
    ICON GetStartupFolder()+"\rc\QRH2TOTVSProtheus.ico";
    CHILD;
    NOMAXIMIZE NOSIZE
    ON INIT  oQRH2TotvsBrowseData:SetFocus()
    
    DEFINE MAIN MENU
        POPUP "Options"
            if (valType(bExec)=="B")
                DEFINE POPUP '&'+cExecTitle
                    ITEM cExecTitle ACTION Eval(bExec,Eval(bGetFil),Eval(bGetMat))
                END POPUP
            endif
            if (lExcel)
                DEFINE POPUP "&Excel"
                    ITEM "Export Browse to &Excel" ACTION fExcel(oQRH2TotvsBrowseData,cTitle+".xls",cTitle)
                END POPUP
            endif
            SEPARATOR
            ITEM '&About' ACTION MsgInfo("Connecti :: Quarta RH To TOTVS Microsiga Protheus ")
            ITEM 'Exit' ACTION ThisWindow.Release
        END POPUP
    END MENU

    @  10,10 TBROWSE oQRH2TotvsBrowseData RECORDSET oRecordSet  EDITABLE AUTOCOLS SELECTOR .T. ;
    WIDTH nBrwWidth HEIGHT nBrwHeight  ;
    FONT Font_QRH2TotvsBrowseData ;
    COLORS CLR_BLACK, CLR_WHITE, CLR_BLACK, { CLR_WHITE, COLOR_GRID }, CLR_BLACK, -CLR_HRED  ;

    oQRH2TotvsBrowseData:aColumns[ 1 ]:lEdit := .F.
    oQRH2TotvsBrowseData:nClrLine := COLOR_GRID

    bGetFil:=oQRH2TotvsBrowseData:GetColumn(1):bData
    bGetMat:=oQRH2TotvsBrowseData:GetColumn(3):bData
    
    if (oQRH2TotvsBrowseData:lDrawSpecHd)
    oQRH2TotvsBrowseData:nClrSpcHdBack := oQRH2TotvsBrowseData:nClrHeadBack
    endif

    ON KEY ESCAPE ACTION ThisWindow.Release

    END WINDOW

    cCodePage:=Hb_SetCodePage('PTISO')

    ACTIVATE WINDOW Form_QRH2TOTVSBrwoseData

    RELEASE FONT Font_QRH2TotvsBrowseData
    if (IsWindowDefined( Form_QRH2TOTVSBrwoseData ))
        RELEASE WINDOW Form_QRH2TOTVSBrwoseData
    endif

    Hb_SetCodePage(cCodePage)

return

procedure QRH2TOTVSProtheusBrowseData2(oRecordSet,cTitle,lExcel)

    local cCodePage as character

    local oQRH2TotvsBrowseData
    local Font_QRH2TotvsBrowseData2
    local Form_QRH2TOTVSBrwoseData2

    local nWinWidth as numeric  := getdesktopwidth()
    local nWinHeight as numeric := getdesktopheight()
    local nBrwWidth as numeric := nWinWidth-30
    local nBrwHeight as numeric := nWinHeight-60

    hb_default(@lExcel,.F.)

    if (!_IsControlDefined ("Font_QRH2TotvsBrowseData2","Main"))
        DEFINE FONT Font_QRH2TotvsBrowseData2 FONTNAME "Arial" SIZE 10
    endif

    DEFINE WINDOW Form_QRH2TOTVSBrwoseData2 AT 0,0 ;
    WIDTH nWinWidth HEIGHT nWinHeight ;
    TITLE cTitle;
    ICON GetStartupFolder()+"\rc\QRH2TOTVSProtheus.ico";
    CHILD;
    NOMAXIMIZE NOSIZE
    ON INIT  oQRH2TotvsBrowseData:SetFocus()
    
    DEFINE MAIN MENU
        POPUP "Options"
            if (lExcel)
                DEFINE POPUP "&Excel"
                    ITEM "Export Browse to &Excel" ACTION fExcel(oQRH2TotvsBrowseData,cTitle+".xls",cTitle)
                END POPUP
            endif
            SEPARATOR
            ITEM '&About' ACTION MsgInfo("Connecti :: Quarta RH To TOTVS Microsiga Protheus ")
            ITEM 'Exit' ACTION ThisWindow.Release
        END POPUP
    END MENU

    @  10,10 TBROWSE oQRH2TotvsBrowseData RECORDSET oRecordSet  EDITABLE AUTOCOLS SELECTOR .T. ;
    WIDTH nBrwWidth HEIGHT nBrwHeight  ;
    FONT Font_QRH2TotvsBrowseData2 ;
    COLORS CLR_BLACK, CLR_WHITE, CLR_BLACK, { CLR_WHITE, COLOR_GRID }, CLR_BLACK, -CLR_HRED  ;

    oQRH2TotvsBrowseData:aColumns[ 1 ]:lEdit := .F.
    oQRH2TotvsBrowseData:nClrLine := COLOR_GRID

    if (oQRH2TotvsBrowseData:lDrawSpecHd)
    oQRH2TotvsBrowseData:nClrSpcHdBack := oQRH2TotvsBrowseData:nClrHeadBack
    endif

    ON KEY ESCAPE ACTION ThisWindow.Release

    END WINDOW

    cCodePage:=Hb_SetCodePage('PTISO')

    ACTIVATE WINDOW Form_QRH2TOTVSBrwoseData2

    RELEASE FONT Font_QRH2TotvsBrowseData2
    if (IsWindowDefined( Form_QRH2TOTVSBrwoseData2 ))
        RELEASE WINDOW Form_QRH2TOTVSBrwoseData2
    endif

    Hb_SetCodePage(cCodePage)

return

Function fExcel(oQRH2TotvsBrowseData,cFile,cTitle)

   Local lActivate, lSave, hFont, nVer:=1

   Default cFile  := Padr( "NoName.xls", 255 ), ;
           cTitle := "TSBrowse/Excel Conectivity"

   lActivate := .T.
   lSave     := .F.
   cTitle    := PadR( cTitle, 255 )

   IF ! _IsControlDefined ("cFont1","Main")
      DEFINE FONT cFont1 FONTNAME "MS Sans Serif" SIZE 11 BOLD
   endif
      hFont := GetFontHandle( "cFont1" )
    IF !IsWIndowDefined ("Form_Excel" )

      DEFINE WINDOW Form_Excel At 150, 150 WIDTH 380 HEIGHT 240 ;
         TITLE cTitle CHILD TOPMOST


      @ 22,12 LABEL Lb1 VALUE "File: "  WIDTH 36

      @ 22 ,62 BTNTEXTBOX BtnTxt1 ;
         HEIGHT 18 ;
         WIDTH 282 ;
         VALUE cFile ;
         ACTION {||Form_Excel.BtnTxt1.Value:= PadR( fSaveFile(cFile), 255 ) }

      @ 54, 12 LABEL Lb2 VALUE "Title "   WIDTH 38

      @ 54 ,62 GETBOX Get_1 ;
         HEIGHT 18;
         WIDTH 282;
         VALUE cTitle;
         ON LOSTFOCUS { || cTitle := Form_Excel.Get_1.Value }

      @ 86 ,62 CHECKBOX Chk_1;
         CAPTION "Open Excel"  ;
         WIDTH 100 HEIGHT 32 ;
         VALUE lActivate

      @ 78 ,164 RADIOGROUP Radio_1 ;
         OPTIONS { "Excel 2", "Excel Ole"};
         VALUE nVer;
         WIDTH 100 ;
         SPACING 22 ;
         ON CHANGE { || nVer := Form_Excel.Radio_1.Value }

      @ 86 ,286 CHECKBOX Chk_2;
         CAPTION "Save File"  ;
         WIDTH 100 HEIGHT 32 ;
         VALUE lSave

         @ 132 ,72 BUTTON btn_Report;
            CAPTION "&Accept" ;
            ACTION {|| lSave := Form_Excel.Chk_2.Value, lActivate := Form_Excel.Chk_1.Value, ;
               If( nVer == 2, ;
               /*oQRH2TotvsBrowseData:ExcelOle( Form_Excel.BtnTxt1.Value, lActivate,GetControlHandle ( "Progress_1", "Form_Excel" ) , { cTitle, hFont }, hFont, lSave )*/;
               oQRH2TotvsBrowseData:Excel2( Form_Excel.BtnTxt1.Value, lActivate,GetControlHandle ( "Progress_1", "Form_Excel" ) , cTitle , lSave ),;
               oQRH2TotvsBrowseData:Excel2( Form_Excel.BtnTxt1.Value, lActivate,GetControlHandle ( "Progress_1", "Form_Excel" ) , cTitle , lSave ) ),;
               Form_Excel.Release };
            WIDTH 76 HEIGHT 24

         @ 132 ,198 BUTTON btn_Excel;
            CAPTION"&Exit"  ;
            ACTION Form_Excel.Release ;
            WIDTH 76 HEIGHT 24

         @ 172 ,12 PROGRESSBAR Progress_1;
            RANGE 1 , 100;
            VALUE 0;
            WIDTH 336 HEIGHT 24


      END WINDOW

      ACTIVATE WINDOW Form_Excel
      RELEASE FONT cFont1

   endif

  oQRH2TotvsBrowseData:GoTop()
  oQRH2TotvsBrowseData:Refresh( .T. )

Return Nil

Static Function fSaveFile(cFile)
RETURN PutFile({{"Excel Book (*.xls)","*.xls"}},"Select the file",,.T.,cFile)
