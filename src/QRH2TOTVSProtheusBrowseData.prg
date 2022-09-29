static procedure QRH2TOTVSProtheusBrowseData(oRecordSet,cTitle,lExcel,cExecTitle,bExec)

    local cCodePage as character

    local bGetFil as codeblock
    local bGetMat as codeblock

    local oQRH2TotvsBrowseData
    local Font_QRH2TotvsBrowseData
    local Form_QRH2TOTVSBrwoseData

    local nGetFil as numeric
    local nGetMat as numeric

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
        POPUP hb_OemToAnsi(hb_UTF8ToStr("&Opções"))
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

    nGetFil:=aScan(oQRH2TotvsBrowseData:aColumns,{|oCol|Upper(allTrim(oCol:cHeading))=="RA_FILIAL"})
    bGetFil:=oQRH2TotvsBrowseData:GetColumn(nGetFil):bData
    
    nGetMat:=aScan(oQRH2TotvsBrowseData:aColumns,{|oCol|Upper(allTrim(oCol:cHeading))=="RA_MAT"})
    bGetMat:=oQRH2TotvsBrowseData:GetColumn(nGetMat):bData
    
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

static procedure QRH2TOTVSProtheusBrowseData2(oRecordSet,cTitle,lExcel)

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
        POPUP hb_UTF8ToStr("&Opções")
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

static procedure QRH2TOTVSProtheusQRHTables(hINI as hash,lExcel as logical)
    
    local cTitle as character
    local cSource as character
    local cFiliais as character := ""

    local hOleConn as hash

    try

        hOleConn:=QRHGetProviders(hINI,1)

        with object hOleConn["SourceConnection"]
            if (:State==adStateOpen )
                hOleConn["QRH"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["QRH"]
                    #pragma __cstream|cSource:=%s
                        SELECT MSysObjects.Name AS table_name
                        FROM MSysObjects
                        WHERE (((Left([Name],1))<>"~") 
                                AND ((Left([Name],4))<>"MSys") 
                                AND ((MSysObjects.Type) In (1,4,6))
                                AND ((MSysObjects.Flags)=0))
                        order by MSysObjects.Name
                    #pragma __endtext
                    cTitle:=hb_OemToAnsi(hb_UTF8ToStr("QRHTables..."))
                    WAIT WINDOW cTitle NOWAIT
                        QRHOpenRecordSet(hOleConn["QRH"],hOleConn["SourceConnection"],cSource,"table_name")
                    WAIT CLEAR
                    if (:eof())
                        MsgInfo(hb_UTF8ToStr("Não Existem Dados para esta consulta"))
                    else
                        QRH2TOTVSProtheusBrowseQRHData(hINI,hOleConn["QRH"],cTitle,lExcel)
                    endif
                    :Close()
                end with
            endif
            :Close()
        end with

    catch 
    
        MsgInfo(hb_UTF8ToStr("Opção Indisponível para o usuário Atual"))
    
    end

return

static procedure QRH2TOTVSProtheusBrowseQRHData(hINI,oRecordSet,cTitle,lExcel)

    local cCodePage as character

    local oQRH2TOTVSProtheusBrowseQRHData
    local Font_QRH2TOTVSProtheusBrowseQRHData
    local Form_QRH2TOTVSProtheusBrowseQRHData

    local nQRHTable as numeric
    local bQRHTable as codeblock

    local nWinWidth as numeric  := getdesktopwidth()
    local nWinHeight as numeric := getdesktopheight()
    local nBrwWidth as numeric := nWinWidth-30
    local nBrwHeight as numeric := nWinHeight-60

    hb_default(@lExcel,.F.)

    if (!_IsControlDefined ("Font_QRH2TOTVSProtheusBrowseQRHData","Main"))
        DEFINE FONT Font_QRH2TOTVSProtheusBrowseQRHData FONTNAME "Arial" SIZE 10
    endif

    DEFINE WINDOW Form_QRH2TOTVSProtheusBrowseQRHData AT 0,0 ;
    WIDTH nWinWidth HEIGHT nWinHeight ;
    TITLE cTitle;
    ICON GetStartupFolder()+"\rc\QRH2TOTVSProtheus.ico";
    CHILD;
    NOMAXIMIZE NOSIZE
    ON INIT  oQRH2TOTVSProtheusBrowseQRHData:SetFocus()
    
    DEFINE MAIN MENU
        POPUP hb_OemToAnsi(hb_UTF8ToStr("&Opções"))
            DEFINE POPUP '&QRHTable'
                ITEM '&QRHTable' ACTION QRH2TOTVSProtheusQRHTableBrowseData(hINI,lExcel,Eval(bQRHTable))
            END POPUP
            if (lExcel)
                DEFINE POPUP "&Excel"
                    ITEM "Export Browse to &Excel" ACTION fExcel(oQRH2TOTVSProtheusBrowseQRHData,cTitle+".xls",cTitle)
                END POPUP
            endif
            SEPARATOR
            ITEM '&About' ACTION MsgInfo("Connecti :: Quarta RH To TOTVS Microsiga Protheus ")
            ITEM 'Exit' ACTION ThisWindow.Release
        END POPUP
    END MENU

    @  10,10 TBROWSE oQRH2TOTVSProtheusBrowseQRHData RECORDSET oRecordSet  EDITABLE AUTOCOLS SELECTOR .T. ;
    WIDTH nBrwWidth HEIGHT nBrwHeight  ;
    FONT Font_QRH2TOTVSProtheusBrowseQRHData ;
    COLORS CLR_BLACK, CLR_WHITE, CLR_BLACK, { CLR_WHITE, COLOR_GRID }, CLR_BLACK, -CLR_HRED  ;

    oQRH2TOTVSProtheusBrowseQRHData:aColumns[ 1 ]:lEdit := .F.
    oQRH2TOTVSProtheusBrowseQRHData:nClrLine := COLOR_GRID
    
    nQRHTable:=aScan(oQRH2TOTVSProtheusBrowseQRHData:aColumns,{|oCol|Upper(allTrim(oCol:cHeading))=="TABLE_NAME"})
    bQRHTable:=oQRH2TOTVSProtheusBrowseQRHData:GetColumn(nQRHTable):bData
    
    if (oQRH2TOTVSProtheusBrowseQRHData:lDrawSpecHd)
    oQRH2TOTVSProtheusBrowseQRHData:nClrSpcHdBack := oQRH2TOTVSProtheusBrowseQRHData:nClrHeadBack
    endif

    ON KEY ESCAPE ACTION ThisWindow.Release

    END WINDOW

    cCodePage:=Hb_SetCodePage('PTISO')

    ACTIVATE WINDOW Form_QRH2TOTVSProtheusBrowseQRHData

    RELEASE FONT Font_QRH2TOTVSProtheusBrowseQRHData
    if (IsWindowDefined( Form_QRH2TOTVSProtheusBrowseQRHData ))
        RELEASE WINDOW Form_QRH2TOTVSProtheusBrowseQRHData
    endif

    Hb_SetCodePage(cCodePage)

return

static procedure QRH2TOTVSProtheusQRHTableBrowseData(hINI as hash,lExcel as logical,cTable)
    
    local cTitle as character
    local cSource as character
    local cFiliais as character := ""

    local hOleConn as hash

    hOleConn:=QRHGetProviders(hINI,1)

    with object hOleConn["SourceConnection"]
        if (:State==adStateOpen )
            hOleConn[cTable]:=TOleAuto():New("ADODB.RecordSet")
            with object hOleConn[cTable]
                #pragma __cstream|cSource:=%s
                    SELECT *
                    FROM cTable
                #pragma __endtext
                cSource:=hb_StrReplace(cSource,{"cTable"=>cTable})
                cTitle:=hb_OemToAnsi(hb_UTF8ToStr(cTable))
                WAIT WINDOW cTitle NOWAIT
                    QRHOpenRecordSet(hOleConn[cTable],hOleConn["SourceConnection"],cSource)
                WAIT CLEAR
                if (:eof())
                    MsgInfo(hb_UTF8ToStr("Não Existem Dados para esta consulta"))
                else
                    QRH2TOTVSProtheusBrowseData2(hOleConn[cTable],cTitle,lExcel)
                endif
                :Close()
            end with
        endif
        :Close()
    end with

return

Function fExcel(oQRH2TotvsBrowseData,cFile,cTitle)

    Local lActivate, lSave, hFont, nVer:=1

    Default cFile  := Padr( "NoName.xls", 255 ), ;
           cTitle := "TSBrowse/Excel Conectivity"

    lActivate := .T.
    lSave     := .F.
    cTitle    := PadR( cTitle, 255 )
   
    Eval(oQRH2TotvsBrowseData:bGoTop)

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
               (ExcelOle(oQRH2TotvsBrowseData,Form_Excel.BtnTxt1.Value,lActivate,GetControlHandle("Progress_1","Form_Excel"),{cTitle,oQRH2TotvsBrowseData:hFont},oQRH2TotvsBrowseData:hFont,lSave,nil,{""}),Eval(oQRH2TotvsBrowseData:bGoTop)),;
               (oQRH2TotvsBrowseData:Excel2(Form_Excel.BtnTxt1.Value,lActivate,GetControlHandle("Progress_1","Form_Excel"),cTitle,lSave),Eval(oQRH2TotvsBrowseData:bGoTop))),;
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
  Eval(oQRH2TotvsBrowseData:bGoTop)
  oQRH2TotvsBrowseData:Refresh( .T. )

Return Nil

static function ExcelOle(oTSB,cXlsFile,lActivate,hProgress,cTitle,hFont,lSave,bExtern,aColSel,bPrintRow)
    local oError
    try
        oTSB:ExcelOle(cXlsFile,lActivate,hProgress,cTitle,hFont,lSave,bExtern,aColSel,bPrintRow)
    catch oError
        oTSB:Excel2(cXlsFile,lActivate,hProgress,cTitle,lSave,bPrintRow)
    end
return

Static Function fSaveFile(cFile)
RETURN PutFile({{"Excel Book (*.xls)","*.xls"}},"Select the file",,.T.,cFile)