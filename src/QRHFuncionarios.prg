#include "minigui.ch"

procedure QRHFuncionarios(hINI as hash)

    local cErrorMsg as character

    local cFilial as character
    local cEmpresa as character
    local cMatricula as character
    local cFuncionarioID as character

    local cTargetField as character

    local cCommonFindKey as character

    local hFields as hash := hINI["Funcionarios"]
    local hOleConn as hash := QRHGetProviders(hINI)

    local lLoop as logical
    local lRecNo as logical
    local lAddNew as logical

    local nEmpresa as numeric
    local nMatricula as numeric
    local nFuncionarioID as numeric

    local nSRARecNo as numeric

    local nRow as numeric
    local nComplete as numeric

    local oError as object

    local xValue

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Funcionários Quarta RH...")) NOWAIT

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
                                    xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa)
                                    if (cTargetField=="RA_FILIAL")
                                        cFilial:=xValue
                                    elseif (cTargetField=="RA_MAT")
                                        cMatricula:=xValue
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
                                lAddNew:=(:eof())
                                if (lAddNew)
                                    :AddNew()
                                endif
                                for each cTargetField in hb_HKeys(hFields)
                                    lLoop:=.F.
                                    lRecNo:=(cTargetField=="R_E_C_N_O_")
                                    if (lRecNo)
                                        if (lAddNew)
                                            xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa,@lLoop,"R_E_C_N_O_",nSRARecNo++)
                                        endif
                                    else
                                        xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa,@lLoop)
                                    endif
                                    if (lLoop)
                                        loop
                                    endif
                                    if (lRecNo)
                                        if (lAddNew)
                                            :Fields(cTargetField):Value:=xValue
                                        endif
                                    else
                                        try
                                            :Fields(cTargetField):Value:=xValue
                                        catch oError
                                            cErrorMsg:="TargetField='cTargetField';Value='xValue';Error='Description'"
                                            MsgInfo(hb_StrReplace(cErrorMsg,{;
                                                'cTargetField'=>cTargetField,;
                                                'xValue'=>cValToChar(xValue),;
                                                'Description'=>oError:Description,;
                                                ";"=>hb_eol();
                                            }))
                                        end try
                                    endif
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
