#include "minigui.ch"

procedure QRHFuncionarios(hINI as hash)

    local cErrorMsg as character

    local cFilial as character
    local cEmpresa as character
    local cMatricula as character
    local cFuncionarioID as character
    local cTOTVSEmpresa as character

    local cSource as string
    local cTargetField as character

    local cCommonFindKey as character

    local hFields as hash := hINI["QRHFuncionarios"]
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
    
    cTOTVSEmpresa:=QRH2TotvsProtheusGetEmpresa(hINI)

    if (empty(cTOTVSEmpresa))
        MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Empresa Inválida")))
        return
    endif

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Funcionários Quarta RH...")) NOWAIT

        with object hOleConn["SourceConnection"]
            if (:State==adStateOpen )
                hOleConn["Funcionarios"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["Funcionarios"]
                    #pragma __cstream|cSource:=%s
                        SELECT * FROM Funcionarios ORDER BY Empresa,Matricula,FuncionarioID
                    #pragma __endtext
                    QRHOpenRecordSet(hOleConn["Funcionarios"],hOleConn["SourceConnection"],cSource,"Empresa,Matricula,FuncionarioID")
                end with
                hOleConn["FuncionarioEndereco"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["FuncionarioEndereco"]
                    #pragma __cstream|cSource:=%s
                        SELECT * FROM FuncionarioEndereco ORDER BY Empresa,Matricula,FuncionarioID
                    #pragma __endtext
                    QRHOpenRecordSet(hOleConn["FuncionarioEndereco"],hOleConn["SourceConnection"],cSource,"Empresa,Matricula,FuncionarioID")
                end with
                hOleConn["FuncionarioDocumentos"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["FuncionarioDocumentos"]
                    #pragma __cstream|cSource:=%s
                        SELECT * FROM FuncionarioDocumentos ORDER BY Empresa,Matricula,FuncionarioID
                    #pragma __endtext
                    QRHOpenRecordSet(hOleConn["FuncionarioDocumentos"],hOleConn["SourceConnection"],cSource,"Empresa,Matricula,FuncionarioID")
                end with
                hOleConn["FuncionarioFoto"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["FuncionarioFoto"]
                    #pragma __cstream|cSource:=%s
                        SELECT * FROM FuncionarioFoto ORDER BY Empresa,Matricula,FuncionarioID
                    #pragma __endtext
                    QRHOpenRecordSet(hOleConn["FuncionarioFoto"],hOleConn["SourceConnection"],cSource,"Empresa,Matricula,FuncionarioID")
                end with
                hOleConn["FuncionarioLancamentos"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["FuncionarioLancamentos"]
                    #pragma __cstream|cSource:=%s
                        SELECT * FROM FuncionarioLancamentos WHERE Codigo=1 AND Tipo='P' ORDER BY Empresa,Matricula,FuncionarioID,Codigo,Tipo
                    #pragma __endtext
                    QRHOpenRecordSet(hOleConn["FuncionarioLancamentos"],hOleConn["SourceConnection"],cSource,"Empresa,Matricula,FuncionarioID,Codigo,Tipo")
                end with
                hOleConn["PontoFuncionarios"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["PontoFuncionarios"]
                    #pragma __cstream|cSource:=%s
                        SELECT * FROM PontoFuncionarios ORDER BY Empresa,Matricula,FuncionarioID
                    #pragma __endtext
                    QRHOpenRecordSet(hOleConn["PontoFuncionarios"],hOleConn["SourceConnection"],cSource,"Empresa,Matricula,FuncionarioID")
                end with
            endif
        end

    WAIT CLEAR

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Funcionários TOTVS Microsiga Protheus...")) NOWAIT

        with object hOleConn["TargetConnection"]
            if (:State==adStateOpen )
                hOleConn["SRARECNO"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["SRARECNO"]
                    #pragma __cstream|cSource:=%s
                        SELECT (MAX(SRA.R_E_C_N_O_)+1) SRARECNO
                          FROM SRA010 SRA
                    #pragma __endtext
                    cSource:=hb_StrReplace(cSource,{"SRA010"=>"SRA"+cTOTVSEmpresa+"0"})
                    QRHOpenRecordSet(hOleConn["SRARECNO"],hOleConn["TargetConnection"],cSource,"SRARECNO")
                    if (:eof())
                        nSRARecNo:=1
                    else
                        nSRARecNo:=:Fields("SRARECNO"):Value
                    endif
                    :Close()
                end with
                hOleConn["SRA"]:=TOleAuto():New("ADODB.RecordSet")
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
                                    if ((:Fields("Codigo")==1).and.(:Fields("Tipo"):Value=="P"))
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

    MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Importação Funcionários Finalizada")))

return