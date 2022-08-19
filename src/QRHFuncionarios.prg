#include "minigui.ch"

procedure QRHFuncionarios(hINI as hash)

    local aTable as array

    local bTransform as codeblock

    local cErrorMsg as character

    local cFilial as character
    local cEmpresa as character
    local cMatricula as character
    local cFuncionarioID as character

    local cSourceField as character
    local cTargetField as character

    local cSourceTable as character
    local cTargetTable as character

    local cTransform as character
    local cCommonFindKey as character

    local hFields as hash := hINI["Funcionarios"]
    local hOleConn as hash := {=>}

    local lTable as logical
    local lAddNew as logical
    local lTransform as logical
    local lFindInTable as logical

    local nEmpresa as numeric
    local nMatricula as numeric
    local nFuncionarioID as numeric

    local nSRARecNo as numeric

    local nRow as numeric
    local nComplete as numeric

    local xValue

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
            :ConnectionString:=hOleConn["SourceProvider"]
            :Open()
            if (:State==adStateOpen )
                hOleConn["Funcionarios"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["Funcionarios"]
                    :CursorLocation:=adUseClient
                    :CursorType:=adOpenDynamic
                    :LockType:=adLockOptimistic
                    :ActiveConnection:=hOleConn["SourceConnection"]
                    #pragma __cstream|:Source:=%s
                        SELECT * FROM Funcionarios ORDER BY Empresa,Matricula,FuncionarioID
                    #pragma __endtext
                        :Open()
                        :Sort:="Empresa,Matricula,FuncionarioID"
                    WAIT CLEAR
                end with
                hOleConn["FuncionarioEndereco"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["FuncionarioEndereco"]
                    :CursorLocation:=adUseClient
                    :CursorType:=adOpenDynamic
                    :LockType:=adLockOptimistic
                    :ActiveConnection:=hOleConn["SourceConnection"]
                    #pragma __cstream|:Source:=%s
                        SELECT * FROM FuncionarioEndereco ORDER BY Empresa,Matricula,FuncionarioID
                    #pragma __endtext
                    :Open()
                    :Sort:="Empresa,Matricula,FuncionarioID"
                end with
                hOleConn["FuncionarioDocumentos"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["FuncionarioDocumentos"]
                    :CursorLocation:=adUseClient
                    :CursorType:=adOpenDynamic
                    :LockType:=adLockOptimistic
                    :ActiveConnection:=hOleConn["SourceConnection"]
                    #pragma __cstream|:Source:=%s
                        SELECT * FROM FuncionarioDocumentos ORDER BY Empresa,Matricula,FuncionarioID
                    #pragma __endtext
                    :Open()
                    :Sort:="Empresa,Matricula,FuncionarioID"
                end with
                hOleConn["FuncionarioFoto"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["FuncionarioFoto"]
                    :CursorLocation:=adUseClient
                    :CursorType:=adOpenDynamic
                    :LockType:=adLockOptimistic
                    :ActiveConnection:=hOleConn["SourceConnection"]
                    #pragma __cstream|:Source:=%s
                        SELECT * FROM FuncionarioFoto ORDER BY Empresa,Matricula,FuncionarioID
                    #pragma __endtext
                    :Open()
                    :Sort:="Empresa,Matricula,FuncionarioID"
                end with
                hOleConn["FuncionarioLancamentos"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["FuncionarioLancamentos"]
                    :CursorLocation:=adUseClient
                    :CursorType:=adOpenDynamic
                    :LockType:=adLockOptimistic
                    :ActiveConnection:=hOleConn["SourceConnection"]
                    #pragma __cstream|:Source:=%s
                        SELECT * FROM FuncionarioLancamentos ORDER BY Empresa,Matricula,FuncionarioID,Tipo
                    #pragma __endtext
                    :Open()
                    :Sort:="Empresa,Matricula,FuncionarioID,Tipo"
                end with
                hOleConn["PontoFuncionarios"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["PontoFuncionarios"]
                    :CursorLocation:=adUseClient
                    :CursorType:=adOpenDynamic
                    :LockType:=adLockOptimistic
                    :ActiveConnection:=hOleConn["SourceConnection"]
                    #pragma __cstream|:Source:=%s
                        SELECT * FROM PontoFuncionarios ORDER BY Empresa,Matricula,FuncionarioID
                    #pragma __endtext
                    :Open()
                    :Sort:="Empresa,Matricula,FuncionarioID"
                end with
            endif
        end

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
            :ConnectionString:=hOleConn["TargetProvider"]
            :Open()
            if (:State==adStateOpen )
                hOleConn["SRA"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["SRA"]
                    :CursorLocation:=adUseClient
                    :CursorType:=adOpenDynamic
                    :LockType:=adLockOptimistic
                    :ActiveConnection:=hOleConn["TargetConnection"]
                    #pragma __cstream|:Source:=%s
                        SELECT (MAX(SRA.R_E_C_N_O_)+1) SRARECNO
                          FROM SRA010 SRA
                    #pragma __endtext
                    :Open()
                    if (:eof())
                        nSRARecNo:=1
                    else
                        nSRARecNo:=:Fields("SRARECNO"):Value
                    endif
                    :Close()
                end with
            endif
        end with

    WAIT CLEAR

    with object hOleConn["SourceConnection"]
        if (:State==adStateOpen )
            with object hOleConn["TargetConnection"]
                if (:State==adStateOpen )
                    CreateProgressBar("Importando "+hb_OemToAnsi(hb_UTF8ToStr("Funcionários"))+"...")
                    with object hOleConn["Funcionarios"]
                        nRow:=0
                        :MoveFirst()
                        while (!:eof())
                            nRow++
                            nEmpresa:=:Fields("Empresa"):Value
                            cEmpresa:=hb_NToS(nEmpresa)
                            nMatricula:=:Fields("Matricula"):Value
                            cMatricula:=hb_NToS(nMatricula)
                            nFuncionarioID:=:Fields("FuncionarioID"):Value
                            cFuncionarioID:=hb_NToS(nFuncionarioID)
                            cCommonFindKey:="Empresa="+cEmpresa
                            cCommonFindKey+=" AND "
                            cCommonFindKey+="Matricula="+cMatricula
                            cCommonFindKey+=" AND "
                            cCommonFindKey:="FuncionarioID="+cFuncionarioID
                            with object hOleConn["FuncionarioEndereco"]
                                :MoveFirst()
                                :Find(cCommonFindKey,0,1)
                            end with
                            with object hOleConn["FuncionarioDocumentos"]
                                :MoveFirst()
                                :Find(cCommonFindKey,0,1)
                            end with
                            with object hOleConn["FuncionarioFoto"]
                                :MoveFirst()
                                :Find(cCommonFindKey,0,1)
                            end with
                            with object hOleConn["FuncionarioLancamentos"]
                                :MoveFirst()
                                :Find(cCommonFindKey,0,1)
                                while (!:eof())
                                    if !(;
                                            (nEmpresa==:Fields("Empresa"):Value);
                                            .and.;
                                            (nMatricula==:Fields("Matricula"):Value);
                                            .and.;
                                            (nFuncionarioID==:Fields("FuncionarioID"):Value);
                                        )
                                        exit
                                    endif
                                    if (:Fields("Tipo"):Value=="1")
                                        exit
                                    endif
                                    :MoveNext()
                                end while
                            end with
                            with object hOleConn["PontoFuncionarios"]
                                :MoveFirst()
                                :Find(cCommonFindKey,0,1)
                            end with
                            for each cTargetField in hb_HKeys(hFields)
                                switch cTargetField
                                  case "RA_FILIAL"
                                  case "RA_MAT"
                                    xValue:=hFields[cTargetField]
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
                                    lTable:=("."$xValue)
                                    if (lTable)
                                        aTable:=hb_ATokens(xValue,".")
                                        lTable:=(len(aTable)>=2)
                                        if (lTable)
                                            cSourceTable:=aTable[1]
                                            cSourceField:=aTable[2]
                                        endif
                                    elseif (empty(xValue).and.(!lTransform))
                                        loop
                                    endif
                                    if (lTable)
                                        with object hOleConn[cSourceTable]
                                            xValue:=:Fields(cSourceField):Value
                                            if (lTransform)
                                                xValue:=Eval(bTransform,xValue,hINI,hOleConn,cFilial,cMatricula,cEmpresa)
                                            endif
                                        end with
                                    elseif (lTransform)
                                        xValue:=Eval(bTransform,xValue,hINI,hOleConn,cFilial,cMatricula,cEmpresa)
                                    endif
                                    if (cTargetField=="RA_MAT")
                                        cMatricula:=xValue
                                    else
                                        cFilial:=xValue
                                    endif
                                end switch
                            next each
                            with object hOleConn["SRA"]
                                :CursorLocation:=adUseClient
                                :CursorType:=adOpenDynamic
                                :LockType:=adLockOptimistic
                                :ActiveConnection:=hOleConn["TargetConnection"]
                                #pragma __cstream|:Source:=%s
                                    SELECT *
                                      FROM SRA010 SRA
                                     WHERE SRA.D_E_L_E_T_=' '
                                       AND SRA.RA_FILIAL='Filial'
                                       AND SRA.RA_MAT='Matricula'
                                     ORDER BY RA_FILIAL
                                             ,RA_MAT
                                #pragma __endtext
                                :Source:=hb_StrReplace(:Source,{'Filial'=>cFilial,'Matricula'=>cMatricula})
                                :Open()
                                :Sort:="RA_FILIAL,RA_MAT"
                                :Find("RA_MAT='"+cMatricula+"'",0,1)
                                lAddNew:=:Eof()
                                if (lAddNew)
                                    :AddNew()
                                endif
                                for each cTargetField in hb_HKeys(hFields)
                                    xValue:=hFields[cTargetField]
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
                                    lTable:=("."$xValue)
                                    if (lTable)
                                        aTable:=hb_ATokens(xValue,".")
                                        lTable:=(len(aTable)>=2)
                                        if (lTable)
                                            cSourceTable:=aTable[1]
                                            cSourceField:=aTable[2]
                                        endif
                                    elseif (empty(xValue).and.(!lTransform))
                                        loop
                                    endif
                                    switch cTargetField
                                    case "R_E_C_N_O_"
                                        if (lAddNew)
                                            xValue:=nSRARecNo++
                                            if (lTransform)
                                                xValue:=Eval(bTransform,xValue,hINI,hOleConn,cFilial,cMatricula,cEmpresa)
                                            endif
                                            :Fields(cTargetField):Value:=xValue
                                        endif
                                        exit
                                    otherwise
                                        if (lTable)
                                            with object hOleConn[cSourceTable]
                                                if (:eof())
                                                    loop
                                                endif
                                                xValue:=:Fields(cSourceField):Value
                                                if (lTransform)
                                                    xValue:=Eval(bTransform,xValue,hINI,hOleConn,cFilial,cMatricula,cEmpresa)
                                                endif
                                            end with
                                        elseif (lTransform)
                                            xValue:=Eval(bTransform,xValue,hINI,hOleConn,cFilial,cMatricula,cEmpresa)
                                        endif
                                        try
                                            :Fields(cTargetField):Value:=xValue
                                        catch e
                                            cErrorMsg:="TargetField='cTargetField';Value='xValue';Error='Description'"
                                            MsgInfo(hb_StrReplace(cErrorMsg,{;
                                                'cTargetField'=>cTargetField,;
                                                'xValue'=>cValToChar(xValue),;
                                                'Description'=>e:Description,;
                                                ";"=>hb_eol();
                                            }))
                                        end
                                    end switch
                                next each
                                :Update()
                                :Close()
                            end whith
                            nComplete:=Int((nRow/:RecordCount)*100)
                            if ((nComplete%10)==0)
                                if (IsWindowDefined(Form_QRH2Protheus))
                                    Form_QRH2Protheus.PrgBar_1.Value:=nComplete
                                    Form_QRH2Protheus.Label_1.Value:="Completed "+hb_NToS(nComplete)+"%"
                                else
                                    exit
                                endif
                                // refreshing
                                InkeyGui()
                            endif
                            :MoveNext()
                        end while
                        :Close()
                    end whith
                    hOleConn["TargetConnection"]:Close()
                    // final waiting
                    InkeyGui( 800 )
                    CloseProgressBar()
                endif
            end whith
            hOleConn["SourceConnection"]:Close()
        endif
    end whith

    MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Importação Funcionários Finalizada")))

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

function GetFieldData(hIni as hash,hTable as hash,cTable as character,cField as character)

    local xValue

    HB_SYMBOL_UNUSED(hIni)

    with object hTable[cTable]
        if (!:eof())
            xValue:=:Fields(cField):Value
        endif
    end with

return(xValue) as date

function AddDaysToDate(hIni as hash,hTable as hash,cTable as character,cFieldDate as character,nDays as numeric)

    local dDate as date
    local dNewDate as date

    HB_SYMBOL_UNUSED(hIni)

    with object hTable[cTable]
        if (!:eof())
            dDate:=:Fields(cFieldDate):Value
            dNewDate:=dDate+nDays
        endif
    end with
    
    hb_default(@dNewDate,CToD("//"))

return(dNewDate) as date
