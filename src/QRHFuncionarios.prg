#include "minigui.ch"

procedure QRHFuncionarios(hIni as hash)

    local aTable as array
    
    local bTransform as codeblock
    
    local cField as character
    local cFilial as character
    local cEmpresa as character
    local cMatricula as character
    local cFuncionarioID as character

    local cTransform as character
    local cCommonFindKey as character

    local hFields as hash := hIni["Funcionarios"]
    local hOleConn as hash := {=>}

    local lTable as logical
    local lAddNew as logical
    local lTRansform as logical

    local nEmpresa as numeric
    local nMatricula as numeric
    local nFuncionarioID as numeric
    
    local nSRARecNo as numeric

    local nRow as numeric
    local nComplete as numeric
    
    local xValue

    WAIT WINDOW "Funcionarios Solotica..." NOWAIT

        //"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=databaseName;User ID=MyUserID;Password=MyPassword;"
        hOleConn["SourceProvider"]:="Provider="+hIni["QRHConnection"]["Provider"]
        hOleConn["SourceProvider"]+=";"
        hOleConn["SourceProvider"]+="Data Source="+hIni["QRHConnection"]["DataSource"]
        hOleConn["SourceProvider"]+=";"
        if ((hb_HHasKey(hIni["QRHConnection"],"UserID")).and.(!Empty(hIni["QRHConnection"]["UserID"])))
            hOleConn["SourceProvider"]+="User ID="+hIni["QRHConnection"]["UserID"]
            hOleConn["SourceProvider"]+=";"
            hOleConn["SourceProvider"]+="Password="+hIni["QRHConnection"]["Password"]
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
            endif
        end

    WAIT CLEAR                                        

    WAIT WINDOW "Funcionarios Protheus..." NOWAIT

        //"Provider=SQLOLEDB;Data Source=serverName;Initial Catalog=databaseName;User ID=MyUserID;Password=MyPassword;"
        hOleConn["TargetProvider"]:="Provider="+hIni["TOTVSConnection"]["Provider"]
        hOleConn["TargetProvider"]+=";"
        hOleConn["TargetProvider"]+="Data Source="+hIni["TOTVSConnection"]["DataSource"]
        hOleConn["TargetProvider"]+=";"
        hOleConn["TargetProvider"]+="Initial Catalog="+hIni["TOTVSConnection"]["InitialCatalog"]
        hOleConn["TargetProvider"]+=";"
        hOleConn["TargetProvider"]+="User ID="+hIni["TOTVSConnection"]["UserID"]
        hOleConn["TargetProvider"]+=";"
        hOleConn["TargetProvider"]+="Password="+hIni["TOTVSConnection"]["Password"]
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
                            nComplete:=Int((nRow/:RecordCount)*100)
                            if ((nComplete%10)==0)
                                if (IsWindowDefined(Form_QRH2Protheus))
                                    Form_QRH2Protheus.PrgBar_1.Value:=nComplete
                                    Form_QRH2Protheus.Label_1.Value:="Completed " + hb_ntos( nComplete ) + "%"
                                else
                                    exit
                                endif
                                // refreshing
                                InkeyGui()
                            endif
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
                                :Source:=hb_StrReplace(:Source,{'Filial'=>cFilial,'Matricula'=>cMatricula} )
                                :Open()
                                lAddNew:=:Eof()
                                if (lAddNew)
                                    :AddNew()
                                endif
                                for each cField in hb_HKeys(hFields)
                                    xValue:=hFields[cField]
                                    lTRansform:=hb_HHasKey(hIni,cField)
                                    if (lTRansform)
                                        cTransform:=if(hb_HHasKey(hIni[cField],"Transform"),hIni[cField]["Transform"],"")
                                        lTRansform:=(!empty(cTransform))
                                        if (lTRansform)
                                            bTransform:=&(cTransform)
                                        endif
                                    endif
                                    lTable:=("."$xValue)
                                    if (lTable)
                                        aTable:=hb_ATokens(xValue,".")
                                        lTable:=(len(aTable)>=2)
                                    elseif (empty(xValue).and.(!lTRansform))
                                        loop
                                    endif
                                    if (hb_HHasKey(hIni,cField))
                                        cTransform:=if(hb_HHasKey(hIni[cField],"Transform"),hIni[cField]["Transform"],"")
                                        if (!empty(cTransform))
                                            bTransform:=&(cTransform)
                                            lTRansform:=.T.
                                        else
                                            lTRansform:=.F.
                                        endif
                                    endif
                                    switch cField
                                    case "R_E_C_N_O_"
                                        if (lAddNew)
                                            if (lTRansform)
                                                :Fields(cField):Value:=Eval(bTransform,nSRARecNo++)
                                            else    
                                                :Fields(cField):Value:=nSRARecNo++
                                            endif
                                        endif
                                        exit
                                    case "RA_BITMAP"
                                        exit
                                    otherwise
                                        if (lTable)
                                            with object hOleConn[aTable[1]]
                                                if (lTRansform)
                                                    xValue:=Eval(bTransform,:Fields(aTable[2]):Value)
                                                else    
                                                    xValue:=:Fields(aTable[2]):Value
                                                endif
                                            end with 
                                        elseif (lTRansform)
                                            xValue:=Eval(bTransform,xValue)
                                        endif
                                        try 
                                            :Fields(cField):Value:=xValue
                                        catch
                                            msginfo(cField)
                                        end
                                    end switch
                                 next each
                                :Update()
                                :Close()
                            end whith
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

function TruncateName(cName as character,nMaxChar as numeric,lRemoveSpace as character)

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
      
        cTruncateName:=allTrim(cName)
        if (len(cTruncateName)<=nMaxChar)
            break
        endif

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
                if ((nString>1).and.(nString<nSplitName))
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