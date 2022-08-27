#include "minigui.ch"

procedure QRHFuncionariosDependentes(hINI as hash)

    local cErrorMsg as character

    local cRBCod as character
    local cFilial as character
    local cEmpresa as character
    local cMatricula as character
    local cFuncionarioID as character
    
    local cTOTVSEmpresa as character

    local cSource as character
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
    
    cTOTVSEmpresa:=QRH2TotvsProtheusGetEmpresa(hINI)

    if (empty(cTOTVSEmpresa))
        MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Empresa Inválida")))
        return
    endif

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Dependentes Quarta RH...")) NOWAIT

        with object hOleConn["SourceConnection"]
            if (:State==adStateOpen )
                hOleConn["Funcionarios"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["Funcionarios"]
                    #pragma __cstream|cSource:=%s
                        SELECT * FROM Funcionarios ORDER BY Empresa,Matricula,FuncionarioID
                    #pragma __endtext
                    QRHOpenRecordSet(hOleConn["Funcionarios"],hOleConn["SourceConnection"],cSource,"Empresa,Matricula,FuncionarioID")
                end with
                hOleConn["FuncionarioDependentes"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["FuncionarioDependentes"]
                    #pragma __cstream|cSource:=%s
                        SELECT * FROM FuncionarioDependentes ORDER BY Empresa,Matricula,FuncionarioID,FuncionarioDependenteID
                    #pragma __endtext
                    QRHOpenRecordSet(hOleConn["FuncionarioDependentes"],hOleConn["SourceConnection"],cSource,"Empresa,Matricula,FuncionarioID,FuncionarioDependenteID")
                end with
            endif
        end

    WAIT CLEAR

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Dependentes TOTVS Microsiga Protheus...")) NOWAIT

        with object hOleConn["TargetConnection"]
            if (:State==adStateOpen )
                hOleConn["SRA"]:=TOleAuto():New("ADODB.RecordSet")
                hOleConn["SRB"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["SRB"]
                    #pragma __cstream|cSource:=%s
                        SELECT (MAX(SRB.R_E_C_N_O_)+1) SRBRECNO
                          FROM SRB010 SRB
                    #pragma __endtext
                    cSource:=hb_StrReplace(cSource,{"SRB010"=>"SRB"+cTOTVSEmpresa+"0"})
                    QRHOpenRecordSet(hOleConn["SRB"],hOleConn["TargetConnection"],cSource,"SRBRECNO")
                    if (:eof())
                        nSRBRecNo:=1
                    else
                        nSRBRecNo:=:Fields("SRBRECNO"):Value
                    endif
                    :Close()
                end with
                hOleConn["SRB"]:=TOleAuto():New("ADODB.RecordSet")
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
                                    with object hOleConn["SRA"]
                                        #pragma __cstream|cSource:=%s
                                            SELECT *
                                              FROM SRA010 SRA
                                             WHERE SRA.D_E_L_E_T_=' '
                                               AND SRA.RA_FILIAL='Filial'
                                               AND SRA.RA_MAT='Matricula'
                                             ORDER BY SRA.RA_FILIAL
                                                     ,SRA.RA_MAT
                                        #pragma __endtext
                                        cSource:=hb_StrReplace(cSource,{"SRA010"=>"SRA"+cTOTVSEmpresa+"0","Filial"=>cFilial,"Matricula"=>cMatricula})
                                        QRHOpenRecordSet(hOleConn["SRA"],hOleConn["TargetConnection"],cSource,"RA_FILIAL,RA_MAT")
                                        :Find("RA_MAT='"+cMatricula+"'",0,1)
                                        if (!:eof())
                                            with object hOleConn["SRB"]
                                                #pragma __cstream|cSource:=%s
                                                    SELECT *
                                                      FROM SRB010 SRB
                                                     WHERE SRB.D_E_L_E_T_=' '
                                                       AND SRB.RB_FILIAL='Filial'
                                                       AND SRB.RB_MAT='Matricula'
                                                       AND SRB.RB_COD='Codigo'
                                                     ORDER BY SRB.RB_FILIAL
                                                             ,SRB.RB_MAT
                                                             ,SRB.RB_COD
                                                #pragma __endtext
                                                cSource:=hb_StrReplace(cSource,{"SRB010"=>"SRB"+cTOTVSEmpresa+"0","Filial"=>cFilial,"Matricula"=>cMatricula,"Codigo"=>cRBCod})
                                                QRHOpenRecordSet(hOleConn["SRB"],hOleConn["TargetConnection"],cSource,"RB_FILIAL,RB_MAT,RB_COD")
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
                                                            switch (:Fields(cTargetField):Type)
                                                              case adBinary
                                                              case adLongVarBinary
                                                                :Fields(cTargetField):AppendChunk(xValue)
                                                                exit
                                                            otherwise
                                                                :Fields(cTargetField):Value:=xValue
                                                            endswitch
                                                        catch oError
                                                            cErrorMsg:="TargetField='cTargetField';Value='xValue';Error='Description'"
                                                            MsgInfo(hb_StrReplace(cErrorMsg,{;
                                                                "cTargetField"=>cTargetField,;
                                                                "xValue"=>cValToChar(xValue),;
                                                                "Description"=>oError:Description,;                                                        
                                                                ";"=>hb_eol();
                                                            }))
                                                        end try
                                                    endif
                                                next each
                                                :Update()
                                                :Close()
                                            end whith
                                        endif
                                    end whith
                                    :MoveNext()
                                end while
                            end whith
                            nComplete:=Int((nRow/:RecordCount)*100)
                            if (Mod(nComplete,10)==0)
                                if (IsWindowDefined(Form_QRH2Protheus))
                                    Form_QRH2Protheus.PrgBar_1.Value:=nComplete
                                    Form_QRH2Protheus.Label_1.Value:=hb_StrReplace("Completed [nRow/:RecordCount]("+hb_NToS(nComplete)+")%",{"nRow"=>hb_NToS(nRow),":RecordCount"=>hb_NToS(:RecordCount)})
                                else
                                    exit
                                endif
                            endif
                            // refreshing
                            InkeyGui()
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
