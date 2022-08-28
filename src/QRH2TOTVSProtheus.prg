/*
 * MINIGUI - Harbour Win32 GUI Quatra RH 2 TOTVS Protheus
 *
 * (c) 2022 Marinaldo de Jesus <marinaldo.jesus@gmail.com>
 */

#include "fileio.ch"
#include "hbclass.ch"

#include "xhb.ch"
#include "minigui.ch"

#include "c:\minigui\source\adordd\adordd.ch"

REQUEST HB_CODEPAGE_PTISO
REQUEST HB_CODEPAGE_UTF8EX

DECLARE WINDOW Form_QRH2Protheus

static st_aTables

procedure main

    local cIni as character := "QRH2TOTVSProtheus.ini"
    
    local hINI as hash

	SET CENTURY ON

    SET DEFAULT Icon TO GetStartupFolder() + "\QRH2TOTVSProtheus.ico"

    DEFINE WINDOW Form_MainQRH2Protheus ;
        AT 0, 0 ;
        WIDTH 600 HEIGHT 400 ;
        TITLE "Connecti :: Quarta RH To TOTVS Microsiga Protheus " ;
        MAIN ;
        ON INIT hINI:=hb_iniRead(cIni)
        DEFINE MAIN MENU
            DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("&Opções"))
                DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("&Importação"))
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Funcionários")) ACTION QRHFuncionarios(hINI)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Dependentes")) ACTION QRHFuncionariosDependentes(hINI)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Afastamentos")) ACTION QRHFuncionariosAfastamentos(hINI)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Histórico Salários")) ACTION QRHFuncionariosHistCargosSalarios(hINI)
                END POPUP
                SEPARATOR
                DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("&Consulta"))
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SRA) &Funcionários ")) ACTION QRHFuncionariosBrowse(hINI)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SRB) &Dependentes")) ACTION QRHFuncionariosDependentesBrowse(hINI)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SR8) &Afastamentos")) ACTION QRHFuncionariosAfastamentosBrowse(hINI)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SR&3) &Hist Salários")) ACTION QRHFuncionariosHistCargosSalariosSR3Browse(hINI)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SR&7) &Hist Salários")) ACTION QRHFuncionariosHistCargosSalariosSR7Browse(hINI)
                END POPUP
                SEPARATOR
                DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("Confi&gurações"))
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Show")) ACTION QRH2TOTVSProtheusViewIni(".\"+cIni)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Reload")) ACTION (hINI:=hb_iniRead(cIni))
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

function QRH2TotvsProtheusGetEmpresa(hINI)

    local cTOTVSEmpresa as character

    if (hb_HHasKey(hINI,"TOTVSConnection"))
        if (hb_HHasKey(hINI["TOTVSConnection"],"TOTVSEmpresa"))
            cTOTVSEmpresa:=hINI["TOTVSConnection"]["TOTVSEmpresa"]
        endif
    endif

return(cTOTVSEmpresa) as character

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

function FindInTable(hINI as hash,cTable as character,xValue,lTIniFile as logical)

    local hTable as hash
    
    local nTIniFile as numeric
    
    local oTIniFile as object
    
    local xTmp

    switch valType(xValue)
      case "C"
        exit
      case "N"
        xValue:=hb_NToS(xValue)
        exit
    otherwise
        xValue:=cValToChar(xValue)
    end switch

    if (hb_HHasKey(hINI,cTable))
        if (hb_HHasKey(hINI[cTable],"FindInTableFile"))
            hb_default(@lTIniFile,.F.)
            if (lTIniFile)
                hb_default(@st_aTables,array(0))
                nTIniFile:=aScan(nTIniFile,{|x|x[1]==cTable})
                if (nTIniFile==0)
                    oTIniFile:=TIniFile():New(hINI[cTable]["FindInTableFile"])
                    aAdd(st_aTables,{cTable,oTIniFile})
                else
                    oTIniFile:=st_aTables[nTIniFile][2]
                endif
                xTmp:=oTIniFile:ReadString(cTable,xValue,"")
                if (empty(xTmp))
                    xTmp:=oTIniFile:ReadString(cTable,"__DFV__","")
                endif
                xValue:=xTmp
            else
                hTable:=hb_iniRead(hINI[cTable]["FindInTableFile"])
                xValue:=FindInTable(hTable,cTable,xValue)
            endif
        elseif (hb_HHasKey(hINI[cTable],xValue))
            xValue:=hINI[cTable][xValue]
        elseif (hb_HHasKey(hINI[cTable],"__DFV__"))
            xValue:=hINI[cTable]["__DFV__"]
        endif
    endif

return(xValue)

function Concatenate(hINI as hash,hTable as hash,cTable as character,cToken as character,...)

    local aParams as array := hb_aParams()

    local cField as character
    local cConcatenate as character

    local hParam as hash

    HB_SYMBOL_UNUSED(hINI)

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

function ImgToFile(hINI as hash,hTable as hash,cTable as character,cField as character,cFilial as character,cMatricula as character)

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

function GetDataField(hINI as hash,hTable as hash,cTable as character,cField as character)

    local xValue

    HB_SYMBOL_UNUSED(hINI)

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

    local cTOTVSEmpresa as character := QRH2TotvsProtheusGetEmpresa(hINI)

    local cTitle as character
    local cSource as character

    local hOleConn as hash

    if (empty(cTOTVSEmpresa))
        MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Empresa Inválida")))
        return
    endif

    hOleConn:=QRHGetProviders(hINI)

    with object hOleConn["TargetConnection"]
        if (:State==adStateOpen )
            hOleConn["SRA"]:=TOleAuto():New("ADODB.RecordSet")
            with object hOleConn["SRA"]
                #pragma __cstream|cSource:=%s
                    SELECT * FROM SRA010 SRA ORDER BY RA_CIC,RA_FILIAL
                #pragma __endtext
                cSource:=hb_StrReplace(cSource,{"SRA010"=>"SRA"+cTOTVSEmpresa+"0"})
                cTitle:=hb_OemToAnsi(hb_UTF8ToStr("Funcionários TOTVS Microsiga Protheus..."))
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

    local cTOTVSEmpresa as character := QRH2TotvsProtheusGetEmpresa(hINI)

    local cTitle as character
    local cSource as character

    local hOleConn as hash

    if (empty(cTOTVSEmpresa))
        MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Empresa Inválida")))
        return
    endif

    hOleConn:=QRHGetProviders(hINI)

    with object hOleConn["TargetConnection"]
        if (:State==adStateOpen )
            hOleConn["SRB"]:=TOleAuto():New("ADODB.RecordSet")
            with object hOleConn["SRB"]
                #pragma __cstream|cSource:=%s
                    SELECT * FROM SRB010 SRB ORDER BY RB_NOME,RB_FILIAL,RB_MAT,RB_COD
                #pragma __endtext
                cSource:=hb_StrReplace(cSource,{"SRB010"=>"SRB"+cTOTVSEmpresa+"0"})
                cTitle:=hb_OemToAnsi(hb_UTF8ToStr("Funcionários/Dependentes TOTVS Microsiga Protheus..."))
                WAIT WINDOW cTitle NOWAIT
                    QRHOpenRecordSet(hOleConn["SRB"],hOleConn["TargetConnection"],cSource,"RB_NOME,RB_FILIAL,RB_MAT,RB_COD")
                WAIT CLEAR
                QRH2TOTVSProtheusBrowseData(hOleConn["SRB"],cTitle)
                :Close()
            end with
        endif
        :Close()
    end with

return

static function QRHFuncionariosAfastamentosBrowse(hINI as hash)

    local cTOTVSEmpresa as character := QRH2TotvsProtheusGetEmpresa(hINI)

    local cTitle as character
    local cSource as character

    local hOleConn as hash

    if (empty(cTOTVSEmpresa))
        MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Empresa Inválida")))
        return
    endif

    hOleConn:=QRHGetProviders(hINI)

    with object hOleConn["TargetConnection"]
        if (:State==adStateOpen )
            hOleConn["SR8"]:=TOleAuto():New("ADODB.RecordSet")
            with object hOleConn["SR8"]
                #pragma __cstream|cSource:=%s
                    SELECT * FROM SR8010 SR8 ORDER BY R8_FILIAL,R8_MAT,R8_SEQ
                #pragma __endtext
                cSource:=hb_StrReplace(cSource,{"SR8010"=>"SR8"+cTOTVSEmpresa+"0"})
                cTitle:=hb_OemToAnsi(hb_UTF8ToStr("Funcionários/Afastamentos TOTVS Microsiga Protheus..."))
                WAIT WINDOW cTitle NOWAIT
                    QRHOpenRecordSet(hOleConn["SR8"],hOleConn["TargetConnection"],cSource,"R8_FILIAL,R8_MAT,R8_SEQ")
                WAIT CLEAR
                QRH2TOTVSProtheusBrowseData(hOleConn["SR8"],cTitle)
                :Close()
            end with
        endif
        :Close()
    end with

return

static function QRHFuncionariosHistCargosSalariosSR3Browse(hINI as hash)

    local cTOTVSEmpresa as character := QRH2TotvsProtheusGetEmpresa(hINI)

    local cTitle as character
    local cSource as character

    local hOleConn as hash

    if (empty(cTOTVSEmpresa))
        MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Empresa Inválida")))
        return
    endif

    hOleConn:=QRHGetProviders(hINI)

    with object hOleConn["TargetConnection"]
        if (:State==adStateOpen )
            hOleConn["SR3"]:=TOleAuto():New("ADODB.RecordSet")
            with object hOleConn["SR3"]
                #pragma __cstream|cSource:=%s
                    SELECT * 
                      FROM SR3010 SR3
                     ORDER BY SR3.R3_FILIAL
                              ,SR3.R3_MAT
                              ,SR3.R3_DATA
                              ,SR3.R3_SEQ
                              ,SR3.R3_TIPO
                              ,SR3.R3_PD

                #pragma __endtext
                cSource:=hb_StrReplace(cSource,{"SR3010"=>"SR3"+cTOTVSEmpresa+"0"})
                cTitle:=hb_OemToAnsi(hb_UTF8ToStr("Funcionários/Histórico Salário SR3 TOTVS Microsiga Protheus..."))
                WAIT WINDOW cTitle NOWAIT
                    QRHOpenRecordSet(hOleConn["SR3"],hOleConn["TargetConnection"],cSource,"R3_FILIAL,R3_MAT,R3_DATA,R3_SEQ,R3_TIPO,R3_PD")
                WAIT CLEAR
                QRH2TOTVSProtheusBrowseData(hOleConn["SR3"],cTitle)
                :Close()
            end with
        endif
        :Close()
    end with

return

static function QRHFuncionariosHistCargosSalariosSR7Browse(hINI as hash)

    local cTOTVSEmpresa as character := QRH2TotvsProtheusGetEmpresa(hINI)

    local cTitle as character
    local cSource as character

    local hOleConn as hash

    if (empty(cTOTVSEmpresa))
        MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Empresa Inválida")))
        return
    endif

    hOleConn:=QRHGetProviders(hINI)

    with object hOleConn["TargetConnection"]
        if (:State==adStateOpen )
            hOleConn["SR7"]:=TOleAuto():New("ADODB.RecordSet")
            with object hOleConn["SR7"]
                #pragma __cstream|cSource:=%s
                    SELECT *
                          ,ISNULL(CAST(CAST(SR7.R7_DESCA AS VARBINARY(MAX)) AS NVARCHAR(MAX)),'') AS [R7_DESCA_MEMOFIELD]
                      FROM SR7010 SR7
                     ORDER BY SR7.R7_FILIAL
                             ,SR7.R7_MAT
                             ,SR7.R7_DATA
                             ,SR7.R7_SEQ
                             ,SR7.R7_TIPO
                #pragma __endtext
                cSource:=hb_StrReplace(cSource,{"SR7010"=>"SR7"+cTOTVSEmpresa+"0"})
                cTitle:=hb_OemToAnsi(hb_UTF8ToStr("Funcionários/Histórico Salário SR7 TOTVS Microsiga Protheus..."))
                WAIT WINDOW cTitle NOWAIT
                    QRHOpenRecordSet(hOleConn["SR7"],hOleConn["TargetConnection"],cSource,"R7_FILIAL,R7_MAT,R7_DATA,R7_SEQ,R7_TIPO")
                WAIT CLEAR
                QRH2TOTVSProtheusBrowseData(hOleConn["SR7"],cTitle)
                :Close()
            end with
        endif
        :Close()
    end with

return

CREATE CLASS TIniFile

   VAR FileName
   VAR Contents

   METHOD New( cFileName )
   METHOD ReadString( cSection, cIdent, cDefault )
   METHOD WriteString( cSection, cIdent, cString )
   METHOD ReadNumber( cSection, cIdent, nDefault )
   METHOD WriteNumber( cSection, cIdent, nNumber )
   METHOD ReadDate( cSection, cIdent, dDefault )
   METHOD WriteDate( cSection, cIdent, dDate )
   METHOD ReadBool( cSection, cIdent, lDefault )
   METHOD WriteBool( cSection, cIdent, lBool )
   METHOD DeleteKey( cSection, cIdent )
   METHOD EraseSection( cSection )
   METHOD ReadSection( cSection )
   METHOD ReadSections()
   METHOD UpdateFile()

END CLASS

METHOD New( cFileName ) CLASS TIniFile

   LOCAL lDone, hFile, cFile, cLine, cIdent, nPos
   LOCAL CurrArray

   IF Empty( cFileName )
      // raise an error?
      OutErr( "No filename passed to TIniFile():New()" )
      RETURN NIL

   ELSE
      ::FileName := cFilename
      ::Contents := {}
      CurrArray := ::Contents

      IF hb_FileExists( cFileName )
         hFile := FOpen( cFilename, FO_READ )
      ELSE
         hFile := FCreate( cFilename )
      ENDIF

      cLine := ""
      lDone := .F.
      DO WHILE ! lDone
         cFile := Space( 256 )
         lDone := ( FRead( hFile, @cFile, 256 ) <= 0 )

         cFile := StrTran( cFile, Chr( 13 ) ) // so we can just search for Chr( 10 )

         // prepend last read
         cFile := cLine + cFile
         DO WHILE ! Empty( cFile )
            IF ( nPos := At( Chr( 10 ), cFile ) ) > 0
               cLine := Left( cFile, nPos - 1 )
               cFile := SubStr( cFile, nPos + 1 )

               IF ! Empty( cLine )
                  IF Left( cLine, 1 ) == "[" // new section
                     IF ( nPos := At( "]", cLine ) ) > 1
                        cLine := SubStr( cLine, 2, nPos - 2 )
                     ELSE
                        cLine := SubStr( cLine, 2 )
                     ENDIF

                     AAdd( ::Contents, { cLine, { /* this will be CurrArray */ } } )
                     CurrArray := ::Contents[ Len( ::Contents ) ][ 2 ]

                  ELSEIF Left( cLine, 1 ) == ";" // preserve comments
                     AAdd( CurrArray, { NIL, cLine } )

                  ELSE
                     IF ( nPos := At( "=", cLine ) ) > 0
                        cIdent := Left( cLine, nPos - 1 )
                        cLine := SubStr( cLine, nPos + 1 )

                        AAdd( CurrArray, { cIdent, cLine } )

                     ELSE
                        AAdd( CurrArray, { cLine, "" } )
                     ENDIF
                  ENDIF
                  cLine := "" // to stop prepend later on
               ENDIF

            ELSE
               cLine := cFile
               cFile := ""
            ENDIF
         ENDDO
      ENDDO

      FClose( hFile )
   ENDIF

   RETURN Self

METHOD ReadString( cSection, cIdent, cDefault ) CLASS TIniFile

   LOCAL cResult := cDefault
   LOCAL i, j, cFind

   IF Empty( cSection )
      cFind := Lower( cIdent )
      j := AScan( ::Contents, {| x | HB_ISSTRING( x[ 1 ] ) .AND. Lower( x[ 1 ] ) == cFind .AND. HB_ISSTRING( x[ 2 ] ) } )

      IF j > 0
         cResult := ::Contents[ j ][ 2 ]
      ENDIF

   ELSE
      cFind := Lower( cSection )
      i := AScan( ::Contents, {| x | HB_ISSTRING( x[ 1 ] ) .AND. Lower( x[ 1 ] ) == cFind } )

      IF i > 0
         cFind := Lower( cIdent )
         j := AScan( ::Contents[ i ][ 2 ], {| x | HB_ISSTRING( x[ 1 ] ) .AND. Lower( x[ 1 ] ) == cFind } )

         IF j > 0
            cResult := ::Contents[ i ][ 2 ][ j ][ 2 ]
         ENDIF
      ENDIF
   ENDIF

   RETURN cResult

METHOD PROCEDURE WriteString( cSection, cIdent, cString ) CLASS TIniFile

   LOCAL i, j, cFind

   IF Empty( cIdent )
      OutErr( "Must specify an identifier" )

   ELSEIF Empty( cSection )
      cFind := Lower( cIdent )
      j := AScan( ::Contents, {| x | HB_ISSTRING( x[ 1 ] ) .AND. Lower( x[ 1 ] ) == cFind .AND. HB_ISSTRING( x[ 2 ] ) } )

      IF j > 0
         ::Contents[ j ][ 2 ] := cString
      ELSE
         AAdd( ::Contents, NIL )
         AIns( ::Contents, 1 )
         ::Contents[ 1 ] := { cIdent, cString }
      ENDIF

   ELSE
      cFind := Lower( cSection )
      IF ( i := AScan( ::Contents, {| x | HB_ISSTRING( x[ 1 ] ) .AND. Lower( x[ 1 ] ) == cFind .AND. HB_ISARRAY( x[ 2 ] ) } ) ) > 0
         cFind := Lower( cIdent )
         j := AScan( ::Contents[ i ][ 2 ], {| x | HB_ISSTRING( x[ 1 ] ) .AND. Lower( x[ 1 ] ) == cFind } )

         IF j > 0
            ::Contents[ i ][ 2 ][ j ][ 2 ] := cString
         ELSE
            AAdd( ::Contents[ i ][ 2 ], { cIdent, cString } )
         ENDIF

      ELSE
         AAdd( ::Contents, { cSection, { { cIdent, cString } } } )
      ENDIF
   ENDIF

   RETURN

METHOD ReadNumber( cSection, cIdent, nDefault ) CLASS TIniFile

   RETURN Val( ::ReadString( cSection, cIdent, Str( nDefault ) ) )

METHOD PROCEDURE WriteNumber( cSection, cIdent, nNumber ) CLASS TIniFile

   ::WriteString( cSection, cIdent, hb_ntos( nNumber ) )

   RETURN

METHOD ReadDate( cSection, cIdent, dDefault ) CLASS TIniFile

   RETURN hb_SToD( ::ReadString( cSection, cIdent, DToS( dDefault ) ) )

METHOD PROCEDURE WriteDate( cSection, cIdent, dDate ) CLASS TIniFile

   ::WriteString( cSection, cIdent, DToS( dDate ) )

   RETURN

METHOD ReadBool( cSection, cIdent, lDefault ) CLASS TIniFile

   LOCAL cDefault := iif( lDefault, ".T.", ".F." )

   RETURN ::ReadString( cSection, cIdent, cDefault ) == ".T."

METHOD PROCEDURE WriteBool( cSection, cIdent, lBool ) CLASS TIniFile

   ::WriteString( cSection, cIdent, iif( lBool, ".T.", ".F." ) )

   RETURN

METHOD PROCEDURE DeleteKey( cSection, cIdent ) CLASS TIniFile

   LOCAL i, j

   cSection := Lower( cSection )
   i := AScan( ::Contents, {| x | HB_ISSTRING( x[ 1 ] ) .AND. Lower( x[ 1 ] ) == cSection } )

   IF i > 0
      cIdent := Lower( cIdent )
      j := AScan( ::Contents[ i ][ 2 ], {| x | HB_ISSTRING( x[ 1 ] ) .AND. Lower( x[ 1 ] ) == cIdent } )

      hb_ADel( ::Contents[ i ][ 2 ], j, .T. )
   ENDIF

   RETURN

METHOD PROCEDURE EraseSection( cSection ) CLASS TIniFile

   LOCAL i

   IF Empty( cSection )
      DO WHILE ( i := AScan( ::Contents, {| x | HB_ISSTRING( x[ 1 ] ) .AND. HB_ISSTRING( x[ 2 ] ) } ) ) > 0
         hb_ADel( ::Contents, i, .T. )
      ENDDO

   ELSE
      cSection := Lower( cSection )
      IF ( i := AScan( ::Contents, {| x | HB_ISSTRING( x[ 1 ] ) .AND. Lower( x[ 1 ] ) == cSection .AND. HB_ISARRAY( x[ 2 ] ) } ) ) > 0
         hb_ADel( ::Contents, i, .T. )
      ENDIF
   ENDIF

   RETURN

METHOD ReadSection( cSection ) CLASS TIniFile

   LOCAL i, j, aSection := {}

   IF Empty( cSection )
      FOR i := 1 TO Len( ::Contents )
         IF HB_ISSTRING( ::Contents[ i ][ 1 ] ) .AND. HB_ISSTRING( ::Contents[ i ][ 2 ] )
            AAdd( aSection, ::Contents[ i ][ 1 ] )
         ENDIF
      NEXT

   ELSE
      cSection := Lower( cSection )
      IF ( i := AScan( ::Contents, {| x | HB_ISSTRING( x[ 1 ] ) .AND. x[ 1 ] == cSection .AND. HB_ISARRAY( x[ 2 ] ) } ) ) > 0

         FOR j := 1 TO Len( ::Contents[ i ][ 2 ] )

            IF ::Contents[ i ][ 2 ][ j ][ 1 ] != NIL
               AAdd( aSection, ::Contents[ i ][ 2 ][ j ][ 1 ] )
            ENDIF
         NEXT
      ENDIF
   ENDIF

   RETURN aSection

METHOD ReadSections() CLASS TIniFile

   LOCAL i, aSections := {}

   FOR i := 1 TO Len( ::Contents )

      IF HB_ISARRAY( ::Contents[ i ][ 2 ] )
         AAdd( aSections, ::Contents[ i ][ 1 ] )
      ENDIF
   NEXT

   RETURN aSections

METHOD PROCEDURE UpdateFile() CLASS TIniFile

   LOCAL i, j

   LOCAL hFile := FCreate( ::Filename )

   FOR i := 1 TO Len( ::Contents )
      IF ::Contents[ i ][ 1 ] == NIL
         FWrite( hFile, ::Contents[ i ][ 2 ] + hb_eol() )

      ELSEIF HB_ISARRAY( ::Contents[ i ][ 2 ] )
         FWrite( hFile, "[" + ::Contents[ i ][ 1 ] + "]" + hb_eol() )
         FOR j := 1 TO Len( ::Contents[ i ][ 2 ] )

            IF ::Contents[ i ][ 2 ][ j ][ 1 ] == NIL
               FWrite( hFile, ::Contents[ i ][ 2 ][ j ][ 2 ] + hb_eol() )
            ELSE
               FWrite( hFile, ::Contents[ i ][ 2 ][ j ][ 1 ] + "=" + ::Contents[ i ][ 2 ][ j ][ 2 ] + hb_eol() )
            ENDIF
         NEXT
         FWrite( hFile, hb_eol() )

      ELSEIF HB_ISSTRING( ::Contents[ i ][ 2 ] )
         FWrite( hFile, ::Contents[ i ][ 1 ] + "=" + ::Contents[ i ][ 2 ] + hb_eol() )

      ENDIF
   NEXT

   FClose( hFile )

   RETURN

#include "QRHFuncionarios.prg"
#include "QRHFuncionariosDependentes.prg"
#include "QRHFuncionariosAfastamentos.prg"
#include "QRHFuncionariosHistCargosSalarios.prg"

#include "QRH2TOTVSProtheusViewIni.prg"
#include "QRH2TOTVSProtheusBrowseData.prg"