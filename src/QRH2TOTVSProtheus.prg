/*
 * MINIGUI - Harbour Win32 GUI Quatra RH 2 TOTVS Protheus
 *
 * (c) 2022 Marinaldo de Jesus <marinaldo.jesus@gmail.com>
 */

#include "minigui.ch"
#include "c:\minigui\source\adordd\adordd.ch"

REQUEST HB_CODEPAGE_PTISO
REQUEST HB_CODEPAGE_UTF8EX

DECLARE WINDOW Form_QRH2Protheus

procedure main

    local hINI as hash

	SET CENTURY ON

    SET DEFAULT Icon TO GetStartupFolder() + "\QRH2TOTVSProtheus.ico"

    DEFINE WINDOW Form_MainQRH2Protheus ;
        AT 0, 0 ;
        WIDTH 600 HEIGHT 400 ;
        TITLE "Connecti :: Quarta RH To TOTVS Microsiga Protheus " ;
        MAIN ;
        ON INIT hINI:=hb_iniRead("QRH2TOTVSProtheus.ini")
        DEFINE MAIN MENU
            DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("&Opções"))
                DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("&Importação"))
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Funcionários")) ACTION QRHFuncionarios(hINI)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Dependentes")) ACTION QRHFuncionariosDependentes(hINI)
                END POPUP
                SEPARATOR
                DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("&Consulta"))
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Funcionários")) ACTION QRHFuncionariosBrowse(hINI)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Dependentes")) ACTION QRHFuncionariosDependentesBrowse(hINI)
                END POPUP
                SEPARATOR
                DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("Confi&gurações"))
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Show")) ACTION QRH2TOTVSProtheusViewIni(".\QRH2TOTVSProtheus.ini")
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Reload")) ACTION (hINI:=hb_iniRead("QRH2TOTVSProtheus.ini"))
                END POPUP
                SEPARATOR
                ITEM 'E&xit' ACTION Form_MainQRH2Protheus.Release()
            END POPUP
            DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("&Sobre"))
                ITEM  "&About" ACTION About()
            END POPUP
        END MENU
		DEFINE STATUSBAR FONT "MS Sans serif" SIZE 8
			STATUSITEM "Connecti :: Quarta RH To TOTVS Microsiga Protheus "
		END STATUSBAR
    ON KEY ESCAPE ACTION ThisWindow.Release
    END WINDOW

    CENTER WINDOW Form_MainQRH2Protheus
    MAXIMIZE WINDOW Form_MainQRH2Protheus
    ACTIVATE WINDOW Form_MainQRH2Protheus

return

static function CreateProgressBar( cTitle )

   DEFINE WINDOW Form_QRH2Protheus ;
      ROW 0 COL 0 ;
      WIDTH 428 HEIGHT 200 ;
      TITLE cTitle ;
      WINDOWTYPE MODAL ;
      NOSIZE ;
      FONT 'Tahoma' SIZE 11

   @ 10, 80 ANIMATEBOX Avi_1 ;
      WIDTH 260 HEIGHT 40 ;
      FILE 'filecopy.avi' ;
      AUTOPLAY TRANSPARENT NOBORDER

   @ 75, 10 LABEL Label_1 ;
      WIDTH 400 HEIGHT 20 ;
      VALUE ''            ;
      CENTERALIGN VCENTERALIGN

   @ 105, 20 PROGRESSBAR PrgBar_1 ;
      RANGE 0, 100 ;
      VALUE 0      ;
      WIDTH 380 HEIGHT 34

   END WINDOW

   Form_QRH2Protheus.Center
   Form_QRH2Protheus.Closable:=.F.

   Activate Window Form_QRH2Protheus NoWait

return NIL

static function CloseProgressBar()

   IF IsWindowDefined( Form_QRH2Protheus )
      Form_QRH2Protheus.Closable:=.T.
      Form_QRH2Protheus.Release
   ENDIF

   DO MESSAGE LOOP

return NIL

static function About()
    local cAbout as character
    local cCopyRight as character
    cAbout:="QRH2TOTVSProtheus :: "
    cAbout+=MiniGuiVersion()
    cCopyRight:="QRH2TOTVSProtheus"
    cCopyRight+=" 1.0"
    cCopyRight+=CRLF
    cCopyRight+=Chr(169)
    cCopyRight+=" marinaldo.jesus@gmail.com"
return(ShellAbout(cAbout,cCopyRight,LoadTrayIcon(GetInstance(),"MAINICON",50,50)))

function QRHGetProviders(hINI)

    local hOleConn as hash := {=>}

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Funcionários Quarta RH...")) NOWAIT

        //"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=databaseName;User ID=MyUserID;Password=MyPassword;"
        hOleConn["SourceProvider"]:="Provider="+hINI["QRHConnection"]["Provider"]
        hOleConn["SourceProvider"]+=";"
        hOleConn["SourceProvider"]+="Data Source="+hINI["QRHConnection"]["DataSource"]
        hOleConn["SourceProvider"]+=";"
        if ((hb_HHasKey(hINI["QRHConnection"],"UserID")).and.(!Empty(hINI["QRHConnection"]["UserID"])))
            hOleConn["SourceProvider"]+="User ID="+hINI["QRHConnection"]["UserID"]
            hOleConn["SourceProvider"]+=";"
            hOleConn["SourceProvider"]+="Password="+hINI["QRHConnection"]["Password"]
            hOleConn["SourceProvider"]+=";"
        endif

        hOleConn["SourceConnection"]:=TOleAuto():new("ADODB.connection")
        with object hOleConn["SourceConnection"]
            :Mode:=3
            :CursorLocation:=adUseClient
            :ConnectionString:=hOleConn["SourceProvider"]
            :Open()
        end with

    WAIT CLEAR

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Funcionários TOTVS Microsiga Protheus...")) NOWAIT

        //"Provider=SQLOLEDB;Data Source=serverName;Initial Catalog=databaseName;User ID=MyUserID;Password=MyPassword;"
        hOleConn["TargetProvider"]:="Provider="+hINI["TOTVSConnection"]["Provider"]
        hOleConn["TargetProvider"]+=";"
        hOleConn["TargetProvider"]+="Data Source="+hINI["TOTVSConnection"]["DataSource"]
        hOleConn["TargetProvider"]+=";"
        hOleConn["TargetProvider"]+="Initial Catalog="+hINI["TOTVSConnection"]["InitialCatalog"]
        hOleConn["TargetProvider"]+=";"
        hOleConn["TargetProvider"]+="User ID="+hINI["TOTVSConnection"]["UserID"]
        hOleConn["TargetProvider"]+=";"
        hOleConn["TargetProvider"]+="Password="+hINI["TOTVSConnection"]["Password"]
        hOleConn["TargetProvider"]+=";"

        hOleConn["TargetConnection"]:=TOleAuto():new("ADODB.connection")
        with object hOleConn["TargetConnection"]
            :Mode:=3
            :CursorLocation:=adUseClient
            :ConnectionString:=hOleConn["TargetProvider"]
            :Open()
        end with

    WAIT CLEAR

return(hOleConn) as hash

procedure QRHOpenRecordSet(oRecordSet,oProvider,cSource,cSort)

    local cStrFindReplace as character

    cSource:=allTrim(cSource)
    
    cStrFindReplace:=hb_eol()
    while cStrFindReplace$cSource
        cSource:=strTran(cSource,cStrFindReplace,"")
    end while

    cStrFindReplace:=chr(10)
    while cStrFindReplace$cSource
        cSource:=strTran(cSource,cStrFindReplace,"")
    end while

    cStrFindReplace:=chr(13)
    while cStrFindReplace$cSource
        cSource:=strTran(cSource,cStrFindReplace,"")
    end while
    
    cSource:=allTrim(cSource)

    cStrFindReplace:="  "
    while cStrFindReplace$cSource
        cSource:=strTran(cSource,cStrFindReplace," ")
    end while

    with object oProvider
        if (:State==adStateOpen )
            with object oRecordSet
                :CacheSize:=100
                :CursorLocation:=adUseClient
                :CursorType:=adOpenDynamic
                :LockType:=adLockOptimistic
                :ActiveConnection:=oProvider
                :Source:=cSource
                :Open()
                :Sort:=cSort
            end with
        endif
    end with

return

function TruncateName(cName as character,nMaxChar as numeric,lRemoveSpace as logical,lFirst as logical)

    local aSplitName as array

    local cString as character
    local cTruncateName as character

    local lTrucateName as logical

    local nString as numeric
    local nSplitName as numeric
    local nTruncateName as numeric

    begin sequence

        hb_default(@nMaxChar,40)
        hb_default(@lRemoveSpace,.T.)
        hb_default(@lFirst,.F.)

        cTruncateName:=allTrim(cName)
        if (len(cTruncateName)<=nMaxChar)
            break
        endif

        cTruncateName:=hb_StrReplace(cTruncateName,{'.'=>" ",':'=>" ",','=>" ",';'=>" ",'-'=>" ",'_'=>" "})

        aSplitName:=hb_aTokens(cTruncateName," ")
        nSplitName:=Len(aSplitName)

        for nString:=1 to nSplitName
            cString:=Upper(aSplitName[nString])
            switch cString
            case "E"
            case "DA"
            case "DE"
            case "DO"
            case "DAS"
            case "DOS"
                aSplitName[nString]:=""
            end switch
        next nString

        while ((nTruncateName:=aScan(aSplitName,{|e|empty(e)}))>0)
            aDel(aSplitName,nTruncateName)
            aSize(aSplitName,--nSplitName)
        end while

        cTruncateName:=""
        aEval(aSplitName,{|e|cTruncateName+=(e+" ")})

        nTruncateName:=0
        cTruncateName:=allTrim(cTruncateName)
        while (len(cTruncateName)>nMaxChar)
            nTruncateName++
            if (nTruncateName>nSplitName)
                exit
            endif
            cTruncateName:=""
            lTrucateName:=.T.
            for nString:=1 to nSplitName
                if ((if(!lFirst,nString>1,lFirst)).and.(nString<nSplitName))
                    if ((len(aSplitName[nString])>1).and.(lTrucateName))
                        lTrucateName:=.F.
                        aSplitName[nString]:=Left(aSplitName[nString],1)
                        cTruncateName+=aSplitName[nString]
                    else
                        cTruncateName+=aSplitName[nString]
                    endif
                else
                    cTruncateName+=aSplitName[nString]
                endif
                cTruncateName+=" "
            next nString
            cTruncateName:=allTrim(cTruncateName)
        end while

        if ((lRemoveSpace).and.(len(cTruncateName)>nMaxChar))
            cTruncateName:=""
            for nString:=1 to nSplitName
                if (nString>1)
                    if (len(aSplitName[nString])==1)
                        cTruncateName+=aSplitName[nString]
                    else
                        cTruncateName+=" "
                        cTruncateName+=aSplitName[nString]
                    endif
                else
                    cTruncateName+=aSplitName[nString]
                    cTruncateName+=" "
                endif
            next nString
            cTruncateName:=allTrim(cTruncateName)
        endif

    end sequence

return(cTruncateName) as character

function FindInTable(hINI as hash,cTable as character,xValue)

    switch valType(xValue)
      case "N"
        xValue:=hb_NToS(xValue)
        exit
    otherwise
        xValue:=cValToChar(xValue)
    end switch

    if (hb_HHasKey(hINI,cTable))
        if (hb_HHasKey(hINI[cTable],xValue))
            xValue:=hINI[cTable][xValue]
        endif
    endif

return(xValue)

function Concatenate(hIni as hash,hTable as hash,cTable as character,cToken as character,...)

    local aParams as array := hb_aParams()

    local cField as character
    local cConcatenate as character

    local hParam as hash

    HB_SYMBOL_UNUSED(hIni)

    hb_default(@cToken,"")

    cConcatenate:=""
    for each hParam in aParams
        if (hParam:__enumIndex<=4)
            loop
        endif
        cField:=hParam:__enumValue
        with object hTable[cTable]
            if (!:eof())
                cConcatenate+=:Fields(cField):Value
                cConcatenate+=cToken
            endif
        end with
    next each

    cConcatenate:=subStr(cConcatenate,1,Len(cConcatenate)-Len(cToken))

return(cConcatenate) as character

function ImgToFile(hIni as hash,hTable as hash,cTable as character,cField as character,cFilial as character,cMatricula as character)

    local cFileIMG as character
    local cFileEXT as character
    local cRABitMap as character

    cRABitMap:=cFilial
    cRABitMap+=cMatricula

    if (hb_HHasKey(hINI,"FuncionariosFoto"))
        if (hb_HHasKey(hINI["FuncionariosFoto"],"path"))
            cFileIMG:=hINI["FuncionariosFoto"]["path"]
        endif
        if (hb_HHasKey(hINI["FuncionariosFoto"],"extension"))
            cFileEXT:=hINI["FuncionariosFoto"]["extension"]
        endif
    endif

    if (empty(cFileIMG))
        cFileIMG:=".\images\"
    endif

    MakeDir(cFileIMG)

    cFileIMG+=cRABitMap
    if (empty(cFileEXT))
        cFileEXT:=".bmp"
    endif
    cFileIMG+=cFileEXT

    with object hTable[cTable]
        if (!:eof())
            if (:Fields(cField):ActualSize>0)
                hb_memoWrit(cFileIMG,:Fields(cField):GetChunk(:Fields(cField):ActualSize))
            endif
        endif
    end with

return(cRABitMap) as character

function GetDataField(hIni as hash,hTable as hash,cTable as character,cField as character)

    local xValue

    HB_SYMBOL_UNUSED(hIni)

    with object hTable[cTable]
        if (!:eof())
            xValue:=:Fields(cField):Value
        endif
    end with

return(xValue) as date

function DateAddDay(dDate as date,nDays as numeric)

    local dNewDate as date

    hb_default(@nDays,0)
    nDays:=Max(nDays,0)
    dNewDate:=dDate+nDays

return(dNewDate) as date

function getTargetFieldValue(hINI as hash,cTargetField as character,hFields as hash,hOleConn as hash,cFilial as character,cMatricula as character,cEmpresa as character,lLoop as logical,cIndexField,nIndexValue)

    local aTable as array

    local bTransform as codeblock

    local cTransform as character

    local cSourceTable as character
    local cSourceField as character

    local lTable as logical
    local lTransform as logical
    local lFindInTable as logical

    local xValue

    begin sequence

        if (cTargetField==cIndexField)
            xValue:=nIndexValue
        else
            xValue:=hFields[cTargetField]
        endif

        lTransform:=hb_HHasKey(hINI,cTargetField)
        if (lTransform)
            cTransform:=if(hb_HHasKey(hINI[cTargetField],"Transform"),hINI[cTargetField]["Transform"],"")
            lTransform:=(!empty(cTransform))
            if (lTransform)
                lFindInTable:=("FindInTable"==cTransform)
                if (lFindInTable)
                    cTransform:=if(hb_HHasKey(hINI[cTargetField],cTransform),hINI[cTargetField][cTransform],"")
                endif
                bTransform:=&(cTransform)
            endif
        endif

        lLoop:=.F.

        lTable:=((valType(xValue)=="C").and.("."$xValue))

        if (lTable)
            aTable:=hb_ATokens(xValue,".")
            lTable:=(len(aTable)>=2)
            if (lTable)
                cSourceTable:=aTable[1]
                cSourceField:=aTable[2]
            endif
        elseif (empty(xValue).and.(!lTransform))
            lLoop:=.T.
            break
        endif

        if (lTable)
            with object hOleConn[cSourceTable]
                lLoop:=(:eof())
                if (lLoop)
                    break
                endif
                xValue:=:Fields(cSourceField):Value
                if (lTransform)
                    xValue:=Eval(bTransform,xValue,hINI,hOleConn,cFilial,cMatricula,cEmpresa)
                endif
            end with
            break
        endif

        if (lTransform)
            xValue:=Eval(bTransform,xValue,hINI,hOleConn,cFilial,cMatricula,cEmpresa)
        endif

    end sequence

return(xValue)

static function QRHFuncionariosBrowse(hINI as hash)

    local cSource as character
    local cTitle as character :=hb_OemToAnsi(hb_UTF8ToStr("Funcionários TOTVS Microsiga Protheus..."))
    local hOleConn as hash := QRHGetProviders(hINI)

    with object hOleConn["TargetConnection"]
        if (:State==adStateOpen )
            hOleConn["SRA"]:=TOleAuto():New("ADODB.RecordSet")
            with object hOleConn["SRA"]
                #pragma __cstream|cSource:=%s
                    SELECT * FROM SRA010 SRA ORDER BY RA_CIC,RA_FILIAL
                #pragma __endtext
                WAIT WINDOW cTitle NOWAIT
                    QRHOpenRecordSet(hOleConn["SRA"],hOleConn["TargetConnection"],cSource,"RA_CIC,RA_FILIAL")
                WAIT CLEAR
                QRH2TOTVSProtheusBrowseData(hOleConn["SRA"],cTitle)
                :Close()
            end with
        endif
        :Close()
    end with

return

static function QRHFuncionariosDependentesBrowse(hINI as hash)

    local cSource as character
    local cTitle as character:=hb_OemToAnsi(hb_UTF8ToStr("Funcionários/Dependentes TOTVS Microsiga Protheus..."))
    local hOleConn as hash := QRHGetProviders(hINI)

    with object hOleConn["TargetConnection"]
        if (:State==adStateOpen )
            hOleConn["SRB"]:=TOleAuto():New("ADODB.RecordSet")
            with object hOleConn["SRB"]
                #pragma __cstream|cSource:=%s
                    SELECT * FROM SRB010 SRA ORDER BY RB_NOME,RB_FILIAL
                #pragma __endtext
                WAIT WINDOW cTitle NOWAIT
                    QRHOpenRecordSet(hOleConn["SRB"],hOleConn["TargetConnection"],cSource,"RB_FILIAL,RB_MAT,RB_COD")
                WAIT CLEAR
                QRH2TOTVSProtheusBrowseData(hOleConn["SRB"],cTitle)
                :Close()
            end with
        endif
        :Close()
    end with

return

#include "QRHFuncionarios.prg"
#include "QRHFuncionariosDependentes.prg"

#include "QRH2TOTVSProtheusViewIni.prg"
#include "QRH2TOTVSProtheusBrowseData.prg"