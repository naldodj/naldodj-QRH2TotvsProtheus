/*
 * MINIGUI - Harbour Win32 GUI Quatra RH 2 TOTVS Protheus
 *
 * (c) 2022 Marinaldo de Jesus <marinaldo.jesus@gmail.com>
 */

#include "xhb.ch"
#include "minigui.ch"
#include "tsbrowse.ch"

#include "c:\minigui\harbour\extras\rddado\adordd.ch"

REQUEST HB_CODEPAGE_PTISO
REQUEST HB_CODEPAGE_UTF8EX

DECLARE WINDOW Form_QRH2Protheus

static st_hTables

procedure main

    local cIni as character := ".\ini\QRH2TOTVSProtheus.ini"

    local hINI as hash

    SET CENTURY ON

    SET DEFAULT Icon TO GetStartupFolder() + "\rc\QRH2TOTVSProtheus.ico"

    DEFINE WINDOW Form_MainQRH2Protheus ;
        AT 0, 0 ;
        WIDTH 600 HEIGHT 400 ;
        TITLE "Connecti :: Quarta RH To TOTVS Microsiga Protheus " ;
        MAIN ;
        ON INIT hINI:=hb_iniRead(cIni)
        DEFINE MAIN MENU
            DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("&Opções"))
                DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("&Importação"))
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Funcionários (SRA)")) ACTION QRHFuncionarios(hINI)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Dependentes (SRB)")) ACTION QRHFuncionariosDependentes(hINI)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Afastamentos (SR8)")) ACTION QRHFuncionariosAfastamentos(hINI)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Histórico Salários (SR3/SR7)")) ACTION QRHFuncionariosHistCargosSalarios(hINI)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Programação de Férias (SRF)")) ACTION QRHFuncionariosHistFeriasSRF(hINI)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Histórico de Férias (SRH)")) ACTION QRHFuncionariosHistFeriasSRH(hINI)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("Afastamento de &Férias (SR8)")) ACTION QRHFuncionariosHistFeriasSR8(hINI)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("Histórico de Folha (SRD)")) ACTION QRHFuncionariosHistFolhaSRD(hINI)
                END POPUP
                SEPARATOR
                DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("&Consulta"))
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SRA) &Funcionários ")) ACTION QRHFuncionariosBrowse(hINI,.F.)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SRB) &Dependentes")) ACTION QRHFuncionariosBrowse(hINI,.F.,hb_OemToAnsi(hb_UTF8ToStr("(SRB) &Dependentes")),{|cFil,cMat|QRHFuncionariosDependentesBrowse(hINI,.F.,cFil,cMat)})
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SR8) &Afastamentos")) ACTION QRHFuncionariosBrowse(hINI,.F.,hb_OemToAnsi(hb_UTF8ToStr("(SR8) &Afastamentos")),{|cFil,cMat|QRHFuncionariosAfastamentosBrowse(hINI,.F.,cFil,cMat)})
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SR&3) Hist Salários")) ACTION QRHFuncionariosBrowse(hINI,.F.,hb_OemToAnsi(hb_UTF8ToStr("(SR&3) Hist Salários")),{|cFil,cMat|QRHFuncionariosHistCargosSalariosSR3Browse(hINI,.F.,cFil,cMat)})
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SR&7) Hist Salários")) ACTION QRHFuncionariosBrowse(hINI,.F.,hb_OemToAnsi(hb_UTF8ToStr("(SR&7) Hist Salários")),{|cFil,cMat|QRHFuncionariosHistCargosSalariosSR7Browse(hINI,.F.,cFil,cMat)})
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SRF) &Programação de Férias")) ACTION QRHFuncionariosBrowse(hINI,.F.,hb_OemToAnsi(hb_UTF8ToStr("(SRF) &Programação de Férias")),{|cFil,cMat|QRHFuncionariosHistFeriasSRFBrowse(hINI,.F.,cFil,cMat)})
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SRH) &Histórico de Férias")) ACTION QRHFuncionariosBrowse(hINI,.F.,hb_OemToAnsi(hb_UTF8ToStr("(SRH) &Histórico de Férias")),{|cFil,cMat|QRHFuncionariosHistFeriasSRHBrowse(hINI,.F.,cFil,cMat)})
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SR8) &Afastamento de &Férias")) ACTION QRHFuncionariosBrowse(hINI,.F.,hb_OemToAnsi(hb_UTF8ToStr("(SR8) &Afastamento de &Férias")),{|cFil,cMat|QRHFuncionariosHistFeriasSR8Browse(hINI,.F.,cFil,cMat)})
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SRD) &Acumulados")) ACTION QRHFuncionariosBrowse(hINI,.F.,hb_OemToAnsi(hb_UTF8ToStr("(SRD) &Acumulados")),{|cFil,cMat|QRHFuncionariosHistFolhaSRDBrowse(hINI,.F.,cFil,cMat)})
                END POPUP
                DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("Consulta &Excel"))
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SRA) &Funcionários ")) ACTION QRHFuncionariosBrowse(hINI,.T.)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SRB) &Dependentes")) ACTION QRHFuncionariosBrowse(hINI,.T.,hb_OemToAnsi(hb_UTF8ToStr("(SRB) &Dependentes")),{|cFil,cMat|QRHFuncionariosDependentesBrowse(hINI,.T.,cFil,cMat)})
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SR8) &Afastamentos")) ACTION QRHFuncionariosBrowse(hINI,.T.,hb_OemToAnsi(hb_UTF8ToStr("(SR8) &Afastamentos")),{|cFil,cMat|QRHFuncionariosAfastamentosBrowse(hINI,.T.,cFil,cMat)})
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SR&3) Hist Salários")) ACTION QRHFuncionariosBrowse(hINI,.T.,hb_OemToAnsi(hb_UTF8ToStr("(SR&3) Hist Salários")),{|cFil,cMat|QRHFuncionariosHistCargosSalariosSR3Browse(hINI,.T.,cFil,cMat)})
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SR&7) Hist Salários")) ACTION QRHFuncionariosBrowse(hINI,.T.,hb_OemToAnsi(hb_UTF8ToStr("(SR&7) Hist Salários")),{|cFil,cMat|QRHFuncionariosHistCargosSalariosSR7Browse(hINI,.T.,cFil,cMat)})
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SRF) &Programação de Férias")) ACTION QRHFuncionariosBrowse(hINI,.T.,hb_OemToAnsi(hb_UTF8ToStr("(SRF) &Programação de Férias")),{|cFil,cMat|QRHFuncionariosHistFeriasSRFBrowse(hINI,.T.,cFil,cMat)})
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SRH) &Histórico de Férias")) ACTION QRHFuncionariosBrowse(hINI,.T.,hb_OemToAnsi(hb_UTF8ToStr("(SRH) &Histórico de Férias")),{|cFil,cMat|QRHFuncionariosHistFeriasSRHBrowse(hINI,.T.,cFil,cMat)})
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SR8) &Afastamento de &Férias")) ACTION QRHFuncionariosBrowse(hINI,.T.,hb_OemToAnsi(hb_UTF8ToStr("(SR8) &Afastamento de &Férias")),{|cFil,cMat|QRHFuncionariosHistFeriasSR8Browse(hINI,.T.,cFil,cMat)})
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("(SRD) &Acumulados")) ACTION QRHFuncionariosBrowse(hINI,.T.,hb_OemToAnsi(hb_UTF8ToStr("(SRD) &Acumulados")),{|cFil,cMat|QRHFuncionariosHistFolhaSRDBrowse(hINI,.T.,cFil,cMat)})
                END POPUP
                SEPARATOR
                DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("Confi&gurações"))
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Show")) ACTION QRH2TOTVSProtheusViewIni(".\"+cIni)
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Reload")) ACTION (hINI:=hb_iniRead(cIni),st_hTables:={=>})
                END POPUP
                SEPARATOR
                DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("QRHTables"))
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&QRHTables")) ACTION QRH2TOTVSProtheusQRHTables(hINI,.T.)
                END POPUP
                SEPARATOR
                ITEM 'E&xit' ACTION Form_MainQRH2Protheus.Release()
            END POPUP
            DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("&Sobre"))
                ITEM  "&About" ACTION About()
            END POPUP
        END MENU
        DEFINE STATUSBAR FONT "MS Sans serif" SIZE 8
            STATUSITEM allTrim((hb_CurDrive()+hb_osDriveSeparator()+hb_ps()+CurDir()))+" :: QRH2TOTVSMicrosigaProtheus"
            STATUSITEM "Connecti"
            STATUSITEM " "
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

   @ 10, 160 ANIMATEBOX Avi_1 ;
      WIDTH 500 HEIGHT 40 ;
      FILE 'QRH2TOTVSProtheus.avi' ;
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

function QRHGetProviders(hINI,nProviders)

    local hOleConn as hash := {=>}
    
    hb_default(@nProviders,3)

    if (nProviders==1).or.(nProviders==3)

        WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Quarta RH Provider...")) NOWAIT

            //"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=databaseName;Jet OLEDB:System Database=system.mdw;User ID=MyUserID;Password=MyPassword;"
            hOleConn["SourceProvider"]:="Provider="+hINI["QRHConnection"]["Provider"]
            hOleConn["SourceProvider"]+=";"
            hOleConn["SourceProvider"]+="Data Source="+hINI["QRHConnection"]["DataSource"]
            hOleConn["SourceProvider"]+=";"
            hOleConn["SourceProviderHasSystemDatabase"]:=(file(GetEnv("APPDATA")+"\Microsoft\Access\System.mdw"))
            if (hOleConn["SourceProviderHasSystemDatabase"])
                hOleConn["SourceProvider"]+="Jet OLEDB:System Database="+GetEnv("APPDATA")+"\Microsoft\Access\System.mdw"
                hOleConn["SourceProvider"]+=";"
            endif
            if ((hb_HHasKey(hINI["QRHConnection"],"UserID")).and.(!Empty(hINI["QRHConnection"]["UserID"])))
                hOleConn["SourceProvider"]+="User ID="+hINI["QRHConnection"]["UserID"]
                hOleConn["SourceProvider"]+=";"
                if ((hb_HHasKey(hINI["QRHConnection"],"Password")).and.(!Empty(hINI["QRHConnection"]["Password"])))
                    hOleConn["SourceProvider"]+="Password="+hINI["QRHConnection"]["Password"]
                    hOleConn["SourceProvider"]+=";"
                endif
            endif

            hOleConn["SourceConnection"]:=TOleAuto():new("ADODB.connection")
            with object hOleConn["SourceConnection"]
                :Mode:=3
                :CursorLocation:=adUseClient
                :ConnectionString:=hOleConn["SourceProvider"]
                :Open()
                if (hOleConn["SourceProviderHasSystemDatabase"])
                    :Execute("GRANT SELECT ON TABLE MSysObjects TO PUBLIC;")
                endif
            end with

        WAIT CLEAR

    endif

    if (nProviders==2).or.(nProviders==3)

        WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("TOTVS Microsiga Protheus Provider...")) NOWAIT

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

    endif

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

function QRH2TotvsProtheusGetFiliais(hINI)

    local cFilial as character
    
    local aFiliais as array := array(0)
    
    local hFiliais as hash

    if (hb_HHasKey(hINI,"TabelaFiliais"))
        hFiliais:=hINI["TabelaFiliais"]
        for each cFilial in hb_HKeys(hFiliais)
            if (aScan(aFiliais,{|cFil|cFil==hFiliais[cFilial]})==0)
                aAdd(aFiliais,hFiliais[cFilial])
            endif
        next each
    endif

return(aFiliais) as array

function TruncateName(cName as character,nMaxChar as numeric,lRemoveSpace as logical,lFirst as logical,lRemoveVowel as logical,hTruncateName as hash)

    local aSplitName as array

    local cString as character
    local cVowels as character := "aAeEiIoOuU"
    local cTruncateName as character

    local lTrucateName as logical

    local nVowel as numeric
    local nString as numeric
    local nSplitName as numeric
    local nTruncateName as numeric

    begin sequence

        hb_default(@nMaxChar,40)
        hb_default(@lRemoveSpace,.T.)
        hb_default(@lFirst,.F.)
        hb_default(@lRemoveVowel,.F.)
        hb_default(@hTruncateName,{=>})

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

        if (len(cTruncateName)>nMaxChar)
            cTruncateName:=hb_strReplace(cTruncateName,hTruncateName)
            aSplitName:=hb_aTokens(cTruncateName," ")
            nSplitName:=Len(aSplitName)
        endif

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
                        if (lRemoveVowel)
                            for nVowel:=1 to Len(cVowels)
                                aSplitName[nString]:=strTran(aSplitName[nString],cVowels[nVowel],"")
                            next nVowel
                        else
                            aSplitName[nString]:=Left(aSplitName[nString],1)
                        endif
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
                hb_default(@st_hTables,{=>})
                if (!hb_HHasKey(st_hTables,cTable))
                    oTIniFile:=TIniFile():New(hINI[cTable]["FindInTableFile"])
                    st_hTables[cTable]:=oTIniFile
                else
                    oTIniFile:=st_hTables[cTable]
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

function LoadFromJSONFile(cJSONFile as character)

    local hJSON as hash

    if (hb_FileExists(cJSONFile))
        if (!hb_HHasKey(st_hTables,cJSONFile))
            hJSON:=hb_JSONDecode(MemoRead(cJSONFile))
            st_hTables[cJSONFile]:=hJSON
        else
            hJSON:=st_hTables[cJSONFile]
        endif
    else
        hJSON:={=>}
    endif

return(hJSON) as hash

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

    if (hb_HHasKey(hTable,cTable))
        with object hTable[cTable]
            if (:State==adStateOpen)
                if (!:eof())
                    xValue:=:Fields(cField):Value
                endif
            endif
        end with
    endif

return(xValue) as date

function DateAddDay(dDate as date,nDays as numeric)

    local dNewDate as date

    hb_default(@nDays,0)
    nDays:=Max(nDays,0)
    dNewDate:=dDate+nDays

return(dNewDate) as date

function getATString(cString,nAT,nMax)
    local nIndex:=(nAT-nMax)
    nIndex:=Min(nIndex,Len(cString))
return(cString[nIndex])

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

        if ((lTable).and.(hb_HHasKey(hOleConn,cSourceTable)))
            with object hOleConn[cSourceTable]
                lLoop:=(:eof())
                if (lLoop)
                    break
                endif
                xValue:=:Fields(cSourceField):Value
                if (lTransform)
                    xValue:=Eval(bTransform,xValue,hINI,hOleConn,cFilial,cMatricula,cEmpresa,hFields)
                endif
            end with
            break
        endif

        if (lTransform)
            xValue:=Eval(bTransform,xValue,hINI,hOleConn,cFilial,cMatricula,cEmpresa,hFields)
        endif

    end sequence

return(xValue)

static procedure QRHFuncionariosBrowse(hINI as hash,lExcel as logical,cExecTitle,bExec)

    local aFiliais as array := QRH2TotvsProtheusGetFiliais(hINI)
    local cTOTVSEmpresa as character := QRH2TotvsProtheusGetEmpresa(hINI)
    
    local cTitle as character
    local cSource as character
    local cFiliais as character := ""

    local hOleConn as hash

    if (empty(cTOTVSEmpresa))
        MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Empresa Inválida")))
        return
    endif
    
    aEval(aFiliais,{|cFil|cFiliais+="'"+cFil+"',"})
    if (Right(cFiliais,1)==",")
        cFiliais:=subStr(cFiliais,1,Len(cFiliais)-1)
    endif

    hOleConn:=QRHGetProviders(hINI)

    with object hOleConn["TargetConnection"]
        if (:State==adStateOpen )
            hOleConn["SRA"]:=TOleAuto():New("ADODB.RecordSet")
            with object hOleConn["SRA"]
                #pragma __cstream|cSource:=%s
                    SELECT *
                      FROM SRA010 SRA
                     WHERE (1=1)
                       AND SRA.D_E_L_E_T_=' '
                       AND SRA.RA_FILIAL IN (cFiliais)
                  ORDER BY RA_FILIAL,RA_MAT
                #pragma __endtext
                cSource:=hb_StrReplace(cSource,{"SRA010"=>"SRA"+cTOTVSEmpresa+"0","cFiliais"=>cFiliais})
                cTitle:=hb_OemToAnsi(hb_UTF8ToStr("Funcionários TOTVS Microsiga Protheus..."))
                WAIT WINDOW cTitle NOWAIT
                    QRHOpenRecordSet(hOleConn["SRA"],hOleConn["TargetConnection"],cSource,"RA_FILIAL,RA_MAT")
                WAIT CLEAR
                QRH2TOTVSProtheusBrowseData(hOleConn["SRA"],cTitle,lExcel,cExecTitle,bExec)
                :Close()
            end with
        endif
        :Close()
    end with

return

static procedure QRHFuncionariosDependentesBrowse(hINI as hash,lExcel as logical,cFil as character,cMat as character)

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
                    SELECT *
                      FROM SRB010 SRB
                      WHERE (1=1)
                        AND SRB.RB_FILIAL='cFil'
                        AND SRB.RB_MAT='cMat'
                        AND SRB.D_E_L_E_T_=' '
                  ORDER BY RB_NOME,RB_FILIAL,RB_MAT,RB_COD
                #pragma __endtext
                cSource:=hb_StrReplace(cSource,{"SRB010"=>"SRB"+cTOTVSEmpresa+"0","cFil"=>cFil,"cMat"=>cMat})
                cTitle:=hb_OemToAnsi(hb_UTF8ToStr("Funcionários/Dependentes TOTVS Microsiga Protheus..."))
                WAIT WINDOW cTitle NOWAIT
                    QRHOpenRecordSet(hOleConn["SRB"],hOleConn["TargetConnection"],cSource,"RB_NOME,RB_FILIAL,RB_MAT,RB_COD")
                WAIT CLEAR
                if (:eof())
                    MsgInfo(hb_UTF8ToStr("Não Existem Dados para esta consulta"))
                else
                    QRH2TOTVSProtheusBrowseData2(hOleConn["SRB"],cTitle,lExcel)
                endif
                :Close()
            end with
        endif
        :Close()
    end with

return

static procedure QRHFuncionariosAfastamentosBrowse(hINI as hash,lExcel as logical,cFil as character,cMat as character)

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
                    SELECT *
                      FROM SR8010 SR8
                     WHERE (1=1)
                       AND SR8.R8_FILIAL='cFil'
                       AND SR8.R8_MAT='cMat'
                       AND SR8.D_E_L_E_T_=' '
                  ORDER BY SR8.R8_FILIAL
                          ,SR8.R8_MAT
                          ,SR8.R8_DATAINI
                          ,SR8.R8_TIPO
                          ,SR8.R8_TIPOAFA
                #pragma __endtext
                cSource:=hb_StrReplace(cSource,{"SR8010"=>"SR8"+cTOTVSEmpresa+"0","cFil"=>cFil,"cMat"=>cMat})
                cTitle:=hb_OemToAnsi(hb_UTF8ToStr("Funcionários/Afastamentos TOTVS Microsiga Protheus..."))
                WAIT WINDOW cTitle NOWAIT
                    QRHOpenRecordSet(hOleConn["SR8"],hOleConn["TargetConnection"],cSource,"R8_FILIAL,R8_MAT,R8_DATAINI,R8_TIPO,R8_TIPOAFA")
                WAIT CLEAR
                if (:eof())
                    MsgInfo(hb_UTF8ToStr("Não Existem Dados para esta consulta"))
                else
                    QRH2TOTVSProtheusBrowseData2(hOleConn["SR8"],cTitle,lExcel)
                endif
                :Close()
            end with
        endif
        :Close()
    end with

return

static procedure QRHFuncionariosHistCargosSalariosSR3Browse(hINI as hash,lExcel as logical,cFil as character,cMat as character)

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
                     WHERE (1=1)
                       AND SR3.R3_FILIAL='cFil'
                       AND SR3.R3_MAT='cMat'
                       AND SR3.D_E_L_E_T_=' '
                     ORDER BY SR3.R3_FILIAL
                              ,SR3.R3_MAT
                              ,SR3.R3_DATA
                              ,SR3.R3_SEQ
                              ,SR3.R3_TIPO
                              ,SR3.R3_PD

                #pragma __endtext
                cSource:=hb_StrReplace(cSource,{"SR3010"=>"SR3"+cTOTVSEmpresa+"0","cFil"=>cFil,"cMat"=>cMat})
                cTitle:=hb_OemToAnsi(hb_UTF8ToStr("Funcionários/Hist.Salário SR3 TOTVS Microsiga Protheus"))
                WAIT WINDOW cTitle NOWAIT
                    QRHOpenRecordSet(hOleConn["SR3"],hOleConn["TargetConnection"],cSource,"R3_FILIAL,R3_MAT,R3_DATA,R3_SEQ,R3_TIPO,R3_PD")
                WAIT CLEAR
                if (:eof())
                    MsgInfo(hb_UTF8ToStr("Não Existem Dados para esta consulta"))
                else
                    QRH2TOTVSProtheusBrowseData2(hOleConn["SR3"],cTitle,lExcel)
                endif
                :Close()
            end with
        endif
        :Close()
    end with

return

static procedure QRHFuncionariosHistCargosSalariosSR7Browse(hINI as hash,lExcel as logical,cFil as character,cMat as character)

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
                     WHERE (1=1)
                       AND SR7.R7_FILIAL='cFil'
                       AND SR7.R7_MAT='cMat'
                       AND SR7.D_E_L_E_T_=' '
                     ORDER BY SR7.R7_FILIAL
                             ,SR7.R7_MAT
                             ,SR7.R7_DATA
                             ,SR7.R7_SEQ
                             ,SR7.R7_TIPO
                #pragma __endtext
                cSource:=hb_StrReplace(cSource,{"SR7010"=>"SR7"+cTOTVSEmpresa+"0","cFil"=>cFil,"cMat"=>cMat})
                cTitle:=hb_OemToAnsi(hb_UTF8ToStr("Funcionários/Hist.Salário SR7 TOTVS Microsiga Protheus"))
                WAIT WINDOW cTitle NOWAIT
                    QRHOpenRecordSet(hOleConn["SR7"],hOleConn["TargetConnection"],cSource,"R7_FILIAL,R7_MAT,R7_DATA,R7_SEQ,R7_TIPO")
                WAIT CLEAR
                if (:eof())
                    MsgInfo(hb_UTF8ToStr("Não Existem Dados para esta consulta"))
                else
                    QRH2TOTVSProtheusBrowseData2(hOleConn["SR7"],cTitle,lExcel)
                endif
                :Close()
            end with
        endif
        :Close()
    end with

return

procedure QRHFuncionariosHistFeriasSRFBrowse(hINI as hash,lExcel as logical,cFil as character,cMat as character)

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
            hOleConn["SRF"]:=TOleAuto():New("ADODB.RecordSet")
            with object hOleConn["SRF"]
                #pragma __cstream|cSource:=%s
                    SELECT *
                      FROM SRF010 SRF
                     WHERE (1=1)
                       AND SRF.RF_FILIAL='cFil'
                       AND SRF.RF_MAT='cMat'
                       AND SRF.D_E_L_E_T_=' '
                  ORDER BY SRF.RF_FILIAL
                          ,SRF.RF_MAT
                          ,SRF.RF_DATABAS
                          ,SRF.RF_PD
                #pragma __endtext
                cSource:=hb_StrReplace(cSource,{"SRF010"=>"SRF"+cTOTVSEmpresa+"0","cFil"=>cFil,"cMat"=>cMat})
                cTitle:=hb_OemToAnsi(hb_UTF8ToStr("Funcionários/Hist.Férias SRF TOTVS Microsiga Protheus"))
                WAIT WINDOW cTitle NOWAIT
                    QRHOpenRecordSet(hOleConn["SRF"],hOleConn["TargetConnection"],cSource,"RF_FILIAL,RF_MAT,RF_DATABAS,RF_PD")
                WAIT CLEAR
                if (:eof())
                    MsgInfo(hb_UTF8ToStr("Não Existem Dados para esta consulta"))
                else
                    QRH2TOTVSProtheusBrowseData2(hOleConn["SRF"],cTitle,lExcel)
                endif
                :Close()
            end with
        endif
        :Close()
    end with

return

procedure QRHFuncionariosHistFeriasSRHBrowse(hINI as hash,lExcel as logical,cFil as character,cMat as character)

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
            hOleConn["SRH"]:=TOleAuto():New("ADODB.RecordSet")
            with object hOleConn["SRH"]
                #pragma __cstream|cSource:=%s
                    SELECT *
                      FROM SRH010 SRH
                     WHERE (1=1)
                       AND SRH.RH_FILIAL='cFil'
                       AND SRH.RH_MAT='cMat'
                       AND SRH.D_E_L_E_T_=' '
                  ORDER BY SRH.RH_FILIAL
                          ,SRH.RH_MAT
                          ,SRH.RH_DATABAS
                          ,SRH.RH_DATAINI
                #pragma __endtext
                cSource:=hb_StrReplace(cSource,{"SRH010"=>"SRH"+cTOTVSEmpresa+"0","cFil"=>cFil,"cMat"=>cMat})
                cTitle:=hb_OemToAnsi(hb_UTF8ToStr("Funcionários/Hist.Férias SRH TOTVS Microsiga Protheus"))
                WAIT WINDOW cTitle NOWAIT
                    QRHOpenRecordSet(hOleConn["SRH"],hOleConn["TargetConnection"],cSource,"RH_FILIAL,RH_MAT,RH_DATABAS,RH_DATAINI")
                WAIT CLEAR
                if (:eof())
                    MsgInfo(hb_UTF8ToStr("Não Existem Dados para esta consulta"))
                else
                    QRH2TOTVSProtheusBrowseData2(hOleConn["SRH"],cTitle,lExcel)
                endif
                :Close()
            end with
        endif
        :Close()
    end with

return

procedure QRHFuncionariosHistFeriasSR8Browse(hINI as hash,lExcel as logical,cFil as character,cMat as character)

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
                    SELECT *
                      FROM SR8010 SR8
                     WHERE (1=1)
                       AND SR8.R8_FILIAL='cFil'
                       AND SR8.R8_MAT='cMat'
                       AND SR8.D_E_L_E_T_=' '
                       AND SR8.R8_TIPO='F'
                  ORDER BY SR8.R8_FILIAL
                          ,SR8.R8_MAT
                          ,SR8.R8_DATAINI
                          ,SR8.R8_TIPO
                          ,SR8.R8_TIPOAFA
                #pragma __endtext
                cSource:=hb_StrReplace(cSource,{"SR8010"=>"SR8"+cTOTVSEmpresa+"0","cFil"=>cFil,"cMat"=>cMat})
                cTitle:=hb_OemToAnsi(hb_UTF8ToStr("Funcionários/Hist.Férias SR8 TOTVS Microsiga Protheus"))
                WAIT WINDOW cTitle NOWAIT
                    QRHOpenRecordSet(hOleConn["SR8"],hOleConn["TargetConnection"],cSource,"R8_FILIAL,R8_MAT,R8_DATAINI,R8_TIPO,R8_TIPOAFA")
                WAIT CLEAR
                if (:eof())
                    MsgInfo(hb_UTF8ToStr("Não Existem Dados para esta consulta"))
                else
                    QRH2TOTVSProtheusBrowseData2(hOleConn["SR8"],cTitle,lExcel)
                endif
                :Close()
            end with
        endif
        :Close()
    end with

return

procedure QRHFuncionariosHistFolhaSRDBrowse(hINI as hash,lExcel as logical,cFil as character,cMat as character)

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
            hOleConn["SRD"]:=TOleAuto():New("ADODB.RecordSet")
            with object hOleConn["SRD"]
                #pragma __cstream|cSource:=%s
                    SELECT *
                      FROM SRD010 SRD
                     WHERE (1=1)
                       AND SRD.RD_FILIAL='cFil'
                       AND SRD.RD_MAT='cMat'
                       AND SRD.D_E_L_E_T_=' '
                  ORDER BY SRD.RD_FILIAL
                          ,SRD.RD_MAT
                          ,SRD.RD_DATARQ
                          ,SRD.RD_PD
                          ,SRD.RD_CC
                          ,SRD.RD_SEQ
                #pragma __endtext
                cSource:=hb_StrReplace(cSource,{"SRD010"=>"SRD"+cTOTVSEmpresa+"0","cFil"=>cFil,"cMat"=>cMat})
                cTitle:=hb_OemToAnsi(hb_UTF8ToStr("Funcionários/Hist.Folha SRD TOTVS Microsiga Protheus"))
                WAIT WINDOW cTitle NOWAIT
                    QRHOpenRecordSet(hOleConn["SRD"],hOleConn["TargetConnection"],cSource,"RD_FILIAL,RD_MAT,RD_DATARQ,RD_PD,RD_CC,RD_SEQ")
                WAIT CLEAR
                if (:eof())
                    MsgInfo(hb_UTF8ToStr("Não Existem Dados para esta consulta"))
                else
                    QRH2TOTVSProtheusBrowseData2(hOleConn["SRD"],cTitle,lExcel)
                endif
                :Close()
            end with
        endif
        :Close()
    end with

return

#include "QRHFuncionarios.prg"
#include "QRHFuncionariosDependentes.prg"
#include "QRHFuncionariosAfastamentos.prg"

#include "QRHFuncionariosHistFeriasSRF.prg"
#include "QRHFuncionariosHistFeriasSRH.prg"
#include "QRHFuncionariosHistFeriasSR8.prg"
#include "QRHFuncionariosHistCargosSalarios.prg"

#include "QRHFuncionariosHistFolhaSRD.prg"

#include "QRH2TOTVSProtheusViewIni.prg"
#include "QRH2TOTVSProtheusTIniFile.prg"
#include "QRH2TOTVSProtheusBrowseData.prg"