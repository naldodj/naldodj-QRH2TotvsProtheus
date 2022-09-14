#include "minigui.ch"

procedure QRHFuncionariosHistFeriasSR8(hINI as hash)

    local cErrorMsg as character

    local cR8SEQ as character
    local cR8Tipo as character
    local cR8TipoAfa as character
    local cR8DataIni as character

    local cFilial as character
    local cEmpresa as character
    local cMatricula as character
    local cFuncionarioID as character

    local cTOTVSEmpresa as character

    local cSource as character
    local cTargetField as character

    local cCommonFindKey as character

    local hFields as hash := hINI["QRHFuncionariosHistFeriasSR8"]
    local hOleConn as hash := QRHGetProviders(hINI)

    local lLoop as logical
    local lRecNo as logical
    local lAddNew as logical
    local lIncSeq as logical

    local nR8SEQ as numeric
    local nEmpresa as numeric
    local nMatricula as numeric
    local nFuncionarioID as numeric

    local nSR8RecNo as numeric

    local nRow as numeric
    local nComplete as numeric

    local oError as object

    local xValue

    cTOTVSEmpresa:=QRH2TotvsProtheusGetEmpresa(hINI)

    if (empty(cTOTVSEmpresa))
        MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Empresa Inválida")))
        return
    endif

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Afastamentos Quarta RH...")) NOWAIT

        with object hOleConn["SourceConnection"]
            if (:State==adStateOpen )
                hOleConn["Funcionarios"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["Funcionarios"]
                    #pragma __cstream|cSource:=%s
                        SELECT * FROM Funcionarios ORDER BY Empresa,Matricula,FuncionarioID
                    #pragma __endtext
                    QRHOpenRecordSet(hOleConn["Funcionarios"],hOleConn["SourceConnection"],cSource,"Empresa,Matricula,FuncionarioID")
                end with
                hOleConn["HistFerias"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["HistFerias"]
                    #pragma __cstream|cSource:=%s
                        SELECT *
                         FROM
                               [HistFerias]
                        WHERE [HistFerias].[DiasFerias]>0
                          AND (
                                    SELECT COUNT(*)
                                      FROM [HistFerias] [HFAB]
                                     WHERE [HistFerias].[FuncionarioID]=[HFAB].[FuncionarioID]
                                       AND [HistFerias].[Empresa]=[HFAB].[Empresa]
                                       AND [HistFerias].[Matricula]=[HFAB].[Matricula]
                                       AND [HistFerias].[RefInicial]=[HFAB].[RefInicial]
                                       AND [HistFerias].[ConcedInicial]<>[HFAB].[ConcedInicial]
                                       AND [HFAB].[AbonoPecuniario]=0
                                       AND [HistFerias].[AbonoPecuniario]<>0
                              )=0
                        ORDER
                        BY Empresa,Matricula,FuncionarioID,RefInicial,ConcedInicial
                        ;
                    #pragma __endtext
                    QRHOpenRecordSet(hOleConn["HistFerias"],hOleConn["SourceConnection"],cSource,"Empresa,Matricula,FuncionarioID,RefInicial,ConcedInicial")
                end with
            endif
        end

    WAIT CLEAR

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Afastamentos TOTVS Microsiga Protheus...")) NOWAIT

        with object hOleConn["TargetConnection"]
            if (:State==adStateOpen )
                hOleConn["SRA"]:=TOleAuto():New("ADODB.RecordSet")
                hOleConn["SR8RECNO"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["SR8RECNO"]
                    #pragma __cstream|cSource:=%s
                        SELECT (MAX(SR8.R_E_C_N_O_)+1) SR8RECNO
                          FROM SR8010 SR8
                    #pragma __endtext
                    cSource:=hb_StrReplace(cSource,{"SR8010"=>"SR8"+cTOTVSEmpresa+"0"})
                    QRHOpenRecordSet(hOleConn["SR8RECNO"],hOleConn["TargetConnection"],cSource,"SR8RECNO")
                    if (:eof())
                        nSR8RecNo:=1
                    else
                        nSR8RecNo:=:Fields("SR8RECNO"):Value
                    endif
                    :Close()
                end with
                hOleConn["SR8"]:=TOleAuto():New("ADODB.RecordSet")
                hOleConn["SR8SEQ"]:=TOleAuto():New("ADODB.RecordSet")
            endif
        end with

    WAIT CLEAR

    with object hOleConn["SourceConnection"]
        if (:State==adStateOpen )
            with object hOleConn["TargetConnection"]
                if (:State==adStateOpen )
                    CreateProgressBar("Importando "+hb_OemToAnsi(hb_UTF8ToStr("Afastamentos de Férias"))+"...")
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
                            with object hOleConn["HistFerias"]
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
                                    for each cTargetField in hb_HKeys(hFields)
                                        switch cTargetField
                                          case "R8_FILIAL"
                                          case "R8_MAT"
                                          case "R8_DATAINI"
                                          case "R8_TIPO"
                                          case "R8_TIPOAFA"
                                            xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa)
                                            if (cTargetField=="R8_FILIAL")
                                                cFilial:=xValue
                                            elseif (cTargetField=="R8_MAT")
                                                cMatricula:=xValue
                                            elseif (cTargetField=="R8_DATAINI")
                                                cR8DataIni:=xValue
                                            elseif (cTargetField=="R8_TIPO")
                                                cR8Tipo:=xValue
                                            elseif (cTargetField=="R8_TIPOAFA")
                                                cR8TipoAfa:=xValue
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
                                            with object hOleConn["SR8SEQ"]
                                                #pragma __cstream|cSource:=%s
                                                    SELECT 'MAX' TYPESEQ
                                                          ,MAX(R8_SEQ) R8_SEQ
                                                      FROM SR8010 SR8
                                                     WHERE SR8.R8_FILIAL='Filial'
                                                       AND SR8.R8_MAT='Matricula'
                                                       AND SR8.R8_DATAINI='DataIni'
                                                     UNION
                                                    SELECT 'DEF' TYPESEQ
                                                          ,R8_SEQ
                                                      FROM SR8010 SR8
                                                     WHERE SR8.R8_FILIAL='Filial'
                                                       AND SR8.R8_MAT='Matricula'
                                                       AND SR8.R8_DATAINI='DataIni'
                                                       AND SR8.R8_TIPO='Tipo'
                                                       AND SR8.R8_TIPOAFA='TipoAfa'
                                                #pragma __endtext
                                                cSource:=hb_StrReplace(cSource,{"SR8010"=>"SR8"+cTOTVSEmpresa+"0","Filial"=>cFilial,"Matricula"=>cMatricula,"DataIni"=>cR8DataIni,"TipoAfa"=>cR8TipoAfa,"Tipo"=>cR8Tipo})
                                                QRHOpenRecordSet(hOleConn["SR8SEQ"],hOleConn["TargetConnection"],cSource,"TYPESEQ")
                                                lIncSeq:=.F.
                                                if (:eof())
                                                    cR8SEQ:="1"
                                                else
                                                    :Find("TYPESEQ='DEF'",0,1)
                                                    if (:eof())
                                                        :Find("TYPESEQ='MAX'",0,1)
                                                        if (:eof())
                                                            cR8SEQ:="1"
                                                        else
                                                            lIncSeq:=.T.
                                                            cR8SEQ:=:Fields("TYPESEQ"):Value
                                                        endif
                                                    else
                                                        cR8SEQ:=:Fields("TYPESEQ"):Value
                                                    endif
                                                endif
                                                nR8SEQ:=val(cR8SEQ)
                                                nR8SEQ+=if(lIncSeq,1,0)
                                                cR8SEQ:=getTargetFieldValue(hIni,"R8_SEQ",hFields,hOleConn,cFilial,cMatricula,cEmpresa,nil,"R8_SEQ",nR8SEQ)
                                                :Close()
                                            end whith
                                            with object hOleConn["SR8"]
                                                #pragma __cstream|cSource:=%s
                                                    SELECT *
                                                      FROM SR8010 SR8
                                                     WHERE SR8.R8_FILIAL='Filial'
                                                       AND SR8.R8_MAT='Matricula'
                                                       AND SR8.R8_DATAINI='DataIni'
                                                       AND SR8.R8_TIPO='Tipo'
                                                       AND SR8.R8_TIPOAFA='TipoAfa'
                                                  ORDER BY SR8.R8_FILIAL
                                                          ,SR8.R8_MAT
                                                          ,SR8.R8_DATAINI
                                                          ,SR8.R8_TIPO
                                                          ,SR8.R8_TIPOAFA
                                                #pragma __endtext
                                                cSource:=hb_StrReplace(cSource,{"SR8010"=>"SR8"+cTOTVSEmpresa+"0","Filial"=>cFilial,"Matricula"=>cMatricula,"DataIni"=>cR8DataIni,"TipoAfa"=>cR8TipoAfa,"Tipo"=>cR8Tipo})
                                                QRHOpenRecordSet(hOleConn["SR8"],hOleConn["TargetConnection"],cSource,"R8_FILIAL,R8_MAT,R8_DATAINI,R8_TIPO,R8_TIPOAFA")
                                                :Find("R8_MAT='"+cMatricula+"'",0,1)
                                                lAddNew:=(:eof())
                                                if (lAddNew)
                                                    :AddNew()
                                                endif
                                                for each cTargetField in hb_HKeys(hFields)
                                                    lLoop:=.F.
                                                    lRecNo:=(cTargetField=="R_E_C_N_O_")
                                                    if (lRecNo)
                                                        if (lAddNew)
                                                            xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa,@lLoop,"R_E_C_N_O_",nSR8RecNo++)
                                                        endif
                                                    elseif (cTargetField=="R8_SEQ")
                                                        xValue:=cR8SEQ
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

    MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Importação Afastamentos de Férias Finalizada")))

return
