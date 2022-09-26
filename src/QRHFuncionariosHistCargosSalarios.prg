#include "minigui.ch"

procedure QRHFuncionariosHistCargosSalarios(hINI as hash)


    local cErrorMsg as character

    local cSeq as character
    local cSeqs as character := "123456789ABCDEFGH1HKLMNOPKRSTUVWXYZ"

    local cData as character
    local cLastData as character

    local cTipo as character

    local cFilial as character
    local cEmpresa as character
    local cMatricula as character
    local cFuncionarioID as character

    local cTOTVSEmpresa as character

    local cSource as character
    local cTargetField as character

    local cCommonFindKey as character

    local hFields as hash := hINI["QRHFuncionariosHistCargosSalarios"]
    local hOleConn as hash := QRHGetProviders(hINI)

    local lLoop as logical
    local lRecNo as logical
    local lAddNew as logical

    local nSeq as numeric

    local nEmpresa as numeric
    local nMatricula as numeric
    local nFuncionarioID as numeric

    local nSR3RecNo as numeric
    local nSR7RecNo as numeric

    local nRow as numeric
    local nComplete as numeric

    local oError as object

    local xValue

    cTOTVSEmpresa:=QRH2TotvsProtheusGetEmpresa(hINI)

    if (empty(cTOTVSEmpresa))
        MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Empresa Inválida")))
        return
    endif

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Históricos de Salários Quarta RH...")) NOWAIT

        with object hOleConn["SourceConnection"]
            if (:State==adStateOpen )
                hOleConn["Funcionarios"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["Funcionarios"]
                    #pragma __cstream|cSource:=%s
                        SELECT * FROM Funcionarios ORDER BY Empresa,Matricula,FuncionarioID
                    #pragma __endtext
                    QRHOpenRecordSet(hOleConn["Funcionarios"],hOleConn["SourceConnection"],cSource,"Empresa,Matricula,FuncionarioID")
                end with
                hOleConn["HistCargosSalarios"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["HistCargosSalarios"]
                    #pragma __cstream|cSource:=%s
                        SELECT * FROM HistCargosSalarios where Salario>0 order by Empresa,FuncionarioID,Matricula,DataAlteracao,HistCargosSalariosID
                    #pragma __endtext
                    QRHOpenRecordSet(hOleConn["HistCargosSalarios"],hOleConn["SourceConnection"],cSource,"Empresa,FuncionarioID,Matricula,DataAlteracao,HistCargosSalariosID")
                end with
            endif
        end

    WAIT CLEAR

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Históricos de Salários TOTVS Microsiga Protheus...")) NOWAIT

        with object hOleConn["TargetConnection"]
            if (:State==adStateOpen )
                hOleConn["SR3RECNO"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["SR3RECNO"]
                    #pragma __cstream|cSource:=%s
                        SELECT (MAX(SR3.R_E_C_N_O_)+1) SR3RECNO
                          FROM SR3010 SR3
                    #pragma __endtext
                    cSource:=hb_StrReplace(cSource,{"SR3010"=>"SR3"+cTOTVSEmpresa+"0"})
                    QRHOpenRecordSet(hOleConn["SR3RECNO"],hOleConn["TargetConnection"],cSource,"SR3RECNO")
                    if (:eof())
                        nSR3RecNo:=1
                    else
                        nSR3RecNo:=:Fields("SR3RECNO"):Value
                    endif
                    :Close()
                end with
                hOleConn["SR7RECNO"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["SR7RECNO"]
                    #pragma __cstream|cSource:=%s
                        SELECT (MAX(SR7.R_E_C_N_O_)+1) SR7RECNO
                          FROM SR7010 SR7
                    #pragma __endtext
                    cSource:=hb_StrReplace(cSource,{"SR7010"=>"SR7"+cTOTVSEmpresa+"0"})
                    QRHOpenRecordSet(hOleConn["SR7RECNO"],hOleConn["TargetConnection"],cSource,"SR7RECNO")
                    if (:eof())
                        nSR7RecNo:=1
                    else
                        nSR7RecNo:=:Fields("SR7RECNO"):Value
                    endif
                    :Close()
                end with
                hOleConn["SR3"]:=TOleAuto():New("ADODB.RecordSet")
                hOleConn["SRA"]:=TOleAuto():New("ADODB.RecordSet")
                hOleConn["SR7"]:=TOleAuto():New("ADODB.RecordSet")
            endif
        end with

    WAIT CLEAR

    with object hOleConn["SourceConnection"]
        if (:State==adStateOpen )
            with object hOleConn["TargetConnection"]
                if (:State==adStateOpen )
                    CreateProgressBar("Importando "+hb_OemToAnsi(hb_UTF8ToStr("Históricos de Salários "))+"...")
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
                            with object hOleConn["HistCargosSalarios"]
                                :MoveFirst()
                                :Find(cCommonFindKey,0,1)
                                nSeq:=0
                                cLastData:=""
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
                                          case "R3_FILIAL"
                                          case "R3_MAT"
                                          case "R3_SEQ"
                                          case "R3_DATA"
                                          case "R3_TIPO"
                                            if (cTargetField=="R3_SEQ")
                                                xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa,nil,"R3_SEQ",nSeq)
                                            else
                                                xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa)
                                            endif
                                            if (cTargetField=="R3_FILIAL")
                                                cFilial:=xValue
                                            elseif (cTargetField=="R3_MAT")
                                                cMatricula:=xValue
                                            elseif (cTargetField=="R3_SEQ")
                                                cSeq:=if(valType(xValue)=="N",cSeqs[xValue],xValue)
                                            elseif (cTargetField=="R3_DATA")
                                                cData:=xValue
                                                if (cLastData!=cData)
                                                    cLastData:=cData
                                                    nSeq:=1
                                                else
                                                    ++nSeq
                                                endif
                                            elseif (cTargetField=="R3_TIPO")
                                                cTipo:=xValue
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
                                            with object hOleConn["SR3"]
                                                #pragma __cstream|cSource:=%s
                                                    SELECT *
                                                      FROM SR3010 SR3
                                                     WHERE SR3.D_E_L_E_T_=' '
                                                       AND SR3.R3_FILIAL='Filial'
                                                       AND SR3.R3_MAT='Matricula'
                                                       AND SR3.R3_DATA='Data'
                                                       AND SR3.R3_SEQ='Sequencia'
                                                       AND SR3.R3_TIPO='Tipo'
                                                       AND SR3.R3_PD='000'
                                                  ORDER BY SR3.R3_FILIAL
                                                          ,SR3.R3_MAT
                                                          ,SR3.R3_DATA
                                                          ,SR3.R3_SEQ
                                                          ,SR3.R3_TIPO
                                                          ,SR3.R3_PD
                                                #pragma __endtext
                                                cSource:=hb_StrReplace(cSource,{"SR3010"=>"SR3"+cTOTVSEmpresa+"0","Filial"=>cFilial,"Matricula"=>cMatricula,"Data"=>cData,"Sequencia"=>cSeq,"Tipo"=>cTipo})
                                                QRHOpenRecordSet(hOleConn["SR3"],hOleConn["TargetConnection"],cSource,"R3_FILIAL,R3_MAT,R3_DATA,R3_SEQ,R3_TIPO,R3_PD")
                                                :Find("R3_MAT='"+cMatricula+"'",0,1)
                                                lAddNew:=(:eof())
                                                if (lAddNew)
                                                    :AddNew()
                                                endif
                                                for each cTargetField in hb_HKeys(hFields)
                                                    try
                                                        lLoop:=empty(:Fields(cTargetField):name)
                                                    catch
                                                        lLoop:=.T.
                                                    end try
                                                    if (lLoop)
                                                        loop
                                                    endif
                                                    lLoop:=.F.
                                                    lRecNo:=(cTargetField=="R_E_C_N_O_")
                                                    if (lRecNo)
                                                        if (lAddNew)
                                                            xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa,@lLoop,"R_E_C_N_O_",nSR3RecNo++)
                                                        endif
                                                    elseif (cTargetField=="R3_SEQ")
                                                        xValue:=cSeq
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
                                            with object hOleConn["SR7"]
                                                #pragma __cstream|cSource:=%s
                                                    SELECT *
                                                      FROM SR7010 SR7
                                                     WHERE SR7.D_E_L_E_T_=' '
                                                       AND SR7.R7_FILIAL='Filial'
                                                       AND SR7.R7_MAT='Matricula'
                                                       AND SR7.R7_DATA='Data'
                                                       AND SR7.R7_SEQ='Sequencia'
                                                       AND SR7.R7_TIPO='Tipo'
                                                  ORDER BY SR7.R7_FILIAL
                                                          ,SR7.R7_MAT
                                                          ,SR7.R7_DATA
                                                          ,SR7.R7_SEQ
                                                          ,SR7.R7_TIPO
                                                #pragma __endtext
                                                cSource:=hb_StrReplace(cSource,{"SR7010"=>"SR7"+cTOTVSEmpresa+"0","Filial"=>cFilial,"Matricula"=>cMatricula,"Data"=>cData,"Sequencia"=>cSeq,"Tipo"=>cTipo})
                                                QRHOpenRecordSet(hOleConn["SR7"],hOleConn["TargetConnection"],cSource,"R7_FILIAL,R7_MAT,R7_DATA,R7_SEQ,R7_TIPO")
                                                :Find("R7_MAT='"+cMatricula+"'",0,1)
                                                lAddNew:=(:eof())
                                                if (lAddNew)
                                                    :AddNew()
                                                endif
                                                for each cTargetField in hb_HKeys(hFields)
                                                    try
                                                        lLoop:=empty(:Fields(cTargetField):name)
                                                    catch
                                                        lLoop:=.T.
                                                    end try
                                                    if (lLoop)
                                                        loop
                                                    endif
                                                    lLoop:=.F.
                                                    lRecNo:=(cTargetField=="R_E_C_N_O_")
                                                    if (lRecNo)
                                                        if (lAddNew)
                                                            xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa,@lLoop,"R_E_C_N_O_",nSR7RecNo++)
                                                        endif
                                                    elseif (cTargetField=="R7_SEQ")
                                                        xValue:=cSeq
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
                                        :Close()
                                    end whith
                                    :MoveNext()
                                    // refreshing
                                    InkeyGui()
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

    MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Importação Históricos de Salários  Finalizada")))

return