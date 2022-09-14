#include "minigui.ch"

procedure QRHFuncionariosHistFeriasSRH(hINI as hash)

    local cErrorMsg as character

    local cFilial as character
    local cEmpresa as character
    local cDataIni as character
    local cDataBase as character
    local cMatricula as character
    local cFuncionarioID as character

    local cTOTVSEmpresa as character

    local cSource as character
    local cTargetField as character

    local cCommonFindKey as character

    local hFields as hash := hINI["QRHFuncionariosHistFeriasSRH"]
    local hOleConn as hash := QRHGetProviders(hINI)

    local lLoop as logical
    local lRecNo as logical
    local lAddNew as logical

    local nEmpresa as numeric
    local nMatricula as numeric
    local nFuncionarioID as numeric

    local nSRHRecNo as numeric

    local nRow as numeric
    local nComplete as numeric

    local oError as object

    local xValue

    cTOTVSEmpresa:=QRH2TotvsProtheusGetEmpresa(hINI)

    if (empty(cTOTVSEmpresa))
        MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Empresa Inválida")))
        return
    endif

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Férias Quarta RH...")) NOWAIT

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
                        SELECT
                               [HistFerias].[FuncionarioID]
                             , [HistFerias].[Empresa]
                             , [HistFerias].[Matricula]
                             , [HistFerias].[RefInicial]
                             , [HistFerias].[RefFinal]
                             , [HistFerias].[Faltas]
                             , [HistFerias].[ConcedInicial]
                             , [HistFerias].[ConcedFinal]
                             , [HistFerias].[DiasFerias]
                             , [HistFerias].[FeriasColetivas]
                             , (
                                    SELECT [HFAB].[AbonoPecuniario]
                                      FROM [HistFerias] [HFAB]
                                      WHERE [HistFerias].[FuncionarioID]=[HFAB].[FuncionarioID]
                                        AND [HistFerias].[RefInicial]=[HFAB].[RefInicial]
                                        AND [HistFerias].[ConcedInicial]<>[HFAB].[ConcedInicial]
                                        AND [HFAB].[AbonoPecuniario]<>0
                                        AND [HistFerias].[AbonoPecuniario]=0
                              ) AS AbonoPecuniario
                             , [HistFerias].[13Salario]
                             , [HistFerias].[Notas]
                             ,IIF(((((DateDiff("m",DateAdd("d",-1,[HistFerias].[RefInicial]),[HistFerias].[RefFinal])/30)*2.5))*30)>30,30,((((DateDiff("m",DateAdd("d",-1,[HistFerias].[RefInicial]),[HistFerias].[RefFinal])/30)*2.5))*30)) AS RH_DFERVEN
                             , (
                                    SELECT [HFAB].[DiasFerias]
                                      FROM [HistFerias] [HFAB]
                                      WHERE [HistFerias].[FuncionarioID]=[HFAB].[FuncionarioID]
                                        AND [HistFerias].[Empresa]=[HFAB].[Empresa]
                                        AND [HistFerias].[Matricula]=[HFAB].[Matricula]
                                        AND [HistFerias].[RefInicial]=[HFAB].[RefInicial]
                                        AND [HistFerias].[ConcedInicial]<>[HFAB].[ConcedInicial]
                                        AND [HFAB].[AbonoPecuniario]=0
                                        AND [HistFerias].[AbonoPecuniario]<>0
                              ) AS RH_DABONPE
                              ,IIF([HistFerias].[13Salario]<>0,50,0) AS RH_PERC13S
                              ,DateAdd("d",-30,[ConcedInicial]) AS RH_DTAVISO
                              ,DateAdd("d",-2,[ConcedInicial]) AS RH_DTRECIB
                              ,DateAdd("d",-2,[ConcedInicial]) AS RH_PERIODO
                              ,IIF([HistFerias].[AbonoPecuniario]<>0,'2','1') AS RH_ABOPEC
                              ,(
                                SELECT Max(Valor)
                                   FROM [HistFolha]
                                  WHERE [HistFolha].[FuncionarioID]=[HistFerias].[FuncionarioID]
                                    AND [HistFolha].[TipoFolha] IN (1,7)
                                    AND [HistFolha].[Codigo]=1
                                    AND (
                                            mid(format([HistFolha].[DataCalculo],'yyyymmdd'),1,6)<=mid(format(DateAdd("d",+30,[HistFerias].[ConcedInicial]),'yyyymmdd'),1,6)
                                        AND mid(format([HistFolha].[DataCalculo],'yyyymmdd'),1,6)>=mid(format(DateAdd("d",-30,[HistFerias].[ConcedInicial]),'yyyymmdd'),1,6)
                                   )
                             ) AS RH_SALARIO
                         FROM
                               [HistFerias]
                        WHERE [HistFerias].[DiasFerias]>0
                        ORDER
                        BY Empresa,Matricula,FuncionarioID,RefInicial,ConcedInicial
                        ;
                    #pragma __endtext
                    QRHOpenRecordSet(hOleConn["HistFerias"],hOleConn["SourceConnection"],cSource,"Empresa,Matricula,FuncionarioID,RefInicial,ConcedInicial")
                end with
            endif
        end

    WAIT CLEAR

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Férias TOTVS Microsiga Protheus...")) NOWAIT

        with object hOleConn["TargetConnection"]
            if (:State==adStateOpen )
                hOleConn["SRHRECNO"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["SRHRECNO"]
                    #pragma __cstream|cSource:=%s
                        SELECT (MAX(SRH.R_E_C_N_O_)+1) SRHRECNO
                          FROM SRH010 SRH
                    #pragma __endtext
                    cSource:=hb_StrReplace(cSource,{"SRH010"=>"SRH"+cTOTVSEmpresa+"0"})
                    QRHOpenRecordSet(hOleConn["SRHRECNO"],hOleConn["TargetConnection"],cSource,"SRHRECNO")
                    if (:eof())
                        nSRHRecNo:=1
                    else
                        nSRHRecNo:=:Fields("SRHRECNO"):Value
                    endif
                    :Close()
                end with
                hOleConn["SRA"]:=TOleAuto():New("ADODB.RecordSet")
                hOleConn["SRH"]:=TOleAuto():New("ADODB.RecordSet")
            endif
        end with

    WAIT CLEAR

    with object hOleConn["SourceConnection"]
        if (:State==adStateOpen )
            with object hOleConn["TargetConnection"]
                if (:State==adStateOpen )
                    CreateProgressBar("Importando "+hb_OemToAnsi(hb_UTF8ToStr("Histórico de Férias"))+"...")
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
                                          case "RH_FILIAL"
                                          case "RH_MAT"
                                          case "RH_DATABAS"
                                          case "RH_DATAINI"
                                            xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa)
                                            if (cTargetField=="RH_FILIAL")
                                                cFilial:=xValue
                                            elseif (cTargetField=="RH_MAT")
                                                cMatricula:=xValue
                                            elseif (cTargetField=="RH_DATABAS")
                                                cDataBase:=xValue
                                            elseif (cTargetField=="RH_DATAINI")
                                                cDataIni:=xValue
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
                                            with object hOleConn["SRH"]
                                                #pragma __cstream|cSource:=%s
                                                    SELECT *
                                                      FROM SRH010 SRH
                                                     WHERE SRH.D_E_L_E_T_=' '
                                                       AND SRH.RH_FILIAL='Filial'
                                                       AND SRH.RH_MAT='Matricula'
                                                       AND SRH.RH_DATABAS='DataBase'
                                                       AND SRH.RH_DATAINI='DataIni'
                                                     ORDER BY SRH.RH_FILIAL
                                                             ,SRH.RH_MAT
                                                             ,SRH.RH_DATABAS
                                                             ,SRH.RH_DATAINI
                                                #pragma __endtext
                                                cSource:=hb_StrReplace(cSource,{"SRH010"=>"SRH"+cTOTVSEmpresa+"0","Filial"=>cFilial,"Matricula"=>cMatricula,"DataBase"=>cDataBase,"DataIni"=>cDataIni})
                                                QRHOpenRecordSet(hOleConn["SRH"],hOleConn["TargetConnection"],cSource,"RH_FILIAL,RH_MAT,RH_DATABAS,RH_DATAINI")
                                                :Find("RH_MAT='"+cMatricula+"'",0,1)
                                                lAddNew:=(:eof())
                                                if (lAddNew)
                                                    :AddNew()
                                                endif
                                                for each cTargetField in hb_HKeys(hFields)
                                                    lLoop:=.F.
                                                    lRecNo:=(cTargetField=="R_E_C_N_O_")
                                                    if (lRecNo)
                                                        if (lAddNew)
                                                            xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa,@lLoop,"R_E_C_N_O_",nSRHRecNo++)
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

    MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Importação Histórico de Férias Finalizada")))

return
