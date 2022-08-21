#include "minigui.ch"

procedure QRHFuncionariosDependentes(hINI as hash)

    local cErrorMsg as character

    local cRBCod as character
    local cFilial as character
    local cEmpresa as character
    local cMatricula as character
    local cFuncionarioID as character

    local cTargetField as character

    local cCommonFindKey as character

    local hFields as hash := hINI["FuncionariosDependentes"]
    local hOleConn as hash := QRHGetProviders(hINI)

    local lLoop as logical
    local lRecNo as logical
    local lAddNew as logical

    local nRBCod as numeric
    local nEmpresa as numeric
    local nMatricula as numeric
    local nFuncionarioID as numeric

    local nSRBRecNo as numeric

    local nRow as numeric
    local nComplete as numeric

    local oError as object

    local xValue

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Dependentes Quarta RH...")) NOWAIT

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
                hOleConn["FuncionarioDependentes"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["FuncionarioDependentes"]
                    :CursorLocation:=adUseClient
                    :CursorType:=adOpenDynamic
                    :LockType:=adLockOptimistic
                    :ActiveConnection:=hOleConn["SourceConnection"]
                    #pragma __cstream|:Source:=%s
                        SELECT * FROM FuncionarioDependentes ORDER BY Empresa,Matricula,FuncionarioID,FuncionarioDependenteID
                    #pragma __endtext
                    :Open()
                    :Sort:="Empresa,Matricula,FuncionarioID,FuncionarioDependenteID"
                end with
            endif
        end

    WAIT CLEAR

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Dependentes TOTVS Microsiga Protheus...")) NOWAIT

        with object hOleConn["TargetConnection"]
            :ConnectionString:=hOleConn["TargetProvider"]
            :Open()
            if (:State==adStateOpen )
                hOleConn["SRB"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["SRB"]
                    :CursorLocation:=adUseClient
                    :CursorType:=adOpenDynamic
                    :LockType:=adLockOptimistic
                    :ActiveConnection:=hOleConn["TargetConnection"]
                    #pragma __cstream|:Source:=%s
                        SELECT (MAX(SRB.R_E_C_N_O_)+1) SRBRECNO
                          FROM SRB010 SRB
                    #pragma __endtext
                    :Open()
                    if (:eof())
                        nSRBRecNo:=1
                    else
                        nSRBRecNo:=:Fields("SRBRECNO"):Value
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
                    CreateProgressBar("Importando "+hb_OemToAnsi(hb_UTF8ToStr("Dependentes"))+"...")
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
                            with object hOleConn["FuncionarioDependentes"]
                                :MoveFirst()
                                :Find(cCommonFindKey,0,1)
                                nRBCod:=0
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
                                    for each cTargetField in hb_HKeys(hFields)
                                        switch cTargetField
                                          case "RB_FILIAL"
                                          case "RB_MAT"
                                          case "RB_COD"
                                            if (cTargetField=="RB_COD")
                                                xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa,nil,"RB_COD",++nRBCod)
                                            else
                                                xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa)
                                            endif
                                            if (cTargetField=="RB_FILIAL")
                                                cFilial:=xValue    
                                            elseif (cTargetField=="RB_MAT")
                                                cMatricula:=xValue
                                            elseif (cTargetField=="RB_COD")
                                                cRBCod:=xValue
                                            endif
                                        end switch
                                    next each                                    
                                    with object hOleConn["SRB"]
                                        :CursorLocation:=adUseClient
                                        :CursorType:=adOpenDynamic
                                        :LockType:=adLockOptimistic
                                        :ActiveConnection:=hOleConn["TargetConnection"]
                                        #pragma __cstream|:Source:=%s
                                            SELECT *
                                              FROM SRB010 SRB
                                             WHERE SRB.D_E_L_E_T_=' '
                                               AND SRB.RB_FILIAL='Filial'
                                               AND SRB.RB_MAT='Matricula'
                                               AND SRB.RB_COD='Codigo'
                                             ORDER BY RB_FILIAL
                                                     ,RB_MAT
                                                     ,RB_COD
                                        #pragma __endtext
                                        :Source:=hb_StrReplace(:Source,{'Filial'=>cFilial,'Matricula'=>cMatricula,'Codigo'=>cRBCod})
                                        :Open()
                                        :Sort:="RB_FILIAL,RB_MAT,RB_COD"
                                        :Find("RB_MAT='"+cMatricula+"'",0,1)
                                        lAddNew:=(:eof())
                                        if (lAddNew)
                                            :AddNew()
                                        endif
                                        for each cTargetField in hb_HKeys(hFields)
                                            lLoop:=.F.
                                            lRecNo:=(cTargetField=="R_E_C_N_O_")
                                            if (lRecNo)
                                                if (lAddNew)
                                                    xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa,@lLoop,"R_E_C_N_O_",nSRBRecNo++)
                                                endif
                                            elseif (cTargetField=="RB_COD")
                                                xValue:=cRBCod
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
                                    :MoveNext()
                                end while
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

    MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Importação Dependentes Finalizada")))

return
