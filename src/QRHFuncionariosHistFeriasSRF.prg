#include "minigui.ch"

procedure QRHFuncionariosHistFeriasSRF(hINI as hash)

    local cErrorMsg as character

    local cFilial as character
    local cEmpresa as character
    local cDataBase as character
    local cMatricula as character
    local cFuncionarioID as character

    local cTOTVSEmpresa as character

    local cSource as character
    local cTargetField as character

    local cCommonFindKey as character

    local hFields as hash := hINI["QRHFuncionariosHistFeriasSRF"]
    local hOleConn as hash := QRHGetProviders(hINI)

    local lLoop as logical
    local lRecNo as logical
    local lAddNew as logical

    local nEmpresa as numeric
    local nMatricula as numeric
    local nFuncionarioID as numeric

    local nSRFRecNo as numeric

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
                        SELECT IIF(DateDiff("m",DateAdd("d",-1,[HistFerias].[RefInicial]),DateSerial(Year(Date()),Month(Date())+1,0))>=12,30,0) AS [RF_DFERVAT]
                              ,IIF(DateDiff("m",DateAdd("d",-1,[HistFerias].[RefInicial]),DateSerial(Year(Date()),Month(Date())+1,0))>=12,0,((((DateDiff("m",DateAdd("d",-1,[HistFerias].[RefInicial]),DateSerial(Year(Date()),Month(Date())+1,0))/30)*2.5))*30)) AS [RF_DFERAAT]
                              ,1 AS [RF_STATUS]
                              ,[HistFerias].[FuncionarioID]
                              ,[HistFerias].[Empresa]
                              ,[HistFerias].[Matricula]
                              ,[HistFerias].[RefInicial]
                              ,[HistFerias].[RefFinal]
                              ,[HistFerias].[ConcedInicial]
                              ,[HistFerias].[ConcedFinal]
                              ,[HistFerias].[Faltas]
                              ,[HistFerias].[DiasFerias]
                              ,[HistFerias].[FeriasColetivas]
                              ,[HistFerias].[AbonoPecuniario]
                              ,[HistFerias].[13Salario]
                              ,[HistFerias].[Notas]
                          FROM [HistFerias]
                        WHERE (
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
                         UNION 
                         SELECT [FuncionarioFerias].[Ferias1Dias] AS [RF_DFERVAT]
                               ,0 AS [RF_DFERAAT]
                               ,0 AS [RF_STATUS]
                               ,[FuncionarioFerias].[FuncionarioID]
                               ,[Funcionarios].[Empresa]
                               ,[Funcionarios].[Matricula]
                               ,[FuncionarioFerias].[Ferias1AqIni] AS [RefInicial]
                               ,[FuncionarioFerias].[Ferias1AqFim] AS [RefFinal]
                               ,[FuncionarioFerias].[Ferias1Ini]   AS [ConcedInicial]
                               ,DateAdd('d',[FuncionarioFerias].[Ferias1Dias],[FuncionarioFerias].[Ferias1Ini]) AS [ConcedFinal]
                               ,0 AS [Faltas]
                               ,[FuncionarioFerias].[Ferias1Dias] AS [DiasFerias]
                               ,[FuncionarioFerias].[UltPerFerColetivas] AS [FeriasColetivas]
                               ,[FuncionarioFerias].[TeraAbonoNasFerias] AS [AbonoPecuniario]
                               ,[FuncionarioFerias].[Tera13NasFerias] AS [13Sarario]
                               ,0 AS [Notas]
                        FROM [FuncionarioFerias]
                        INNER JOIN [Funcionarios]
                            ON (Funcionarios.FuncionarioID=[FuncionarioFerias].[FuncionarioID])
                         UNION 
                         SELECT [FuncionarioFerias].[Ferias2Dias] AS [RF_DFERVAT]
                               ,0 AS [RF_DFERAAT]
                               ,0 AS [RF_STATUS]
                               ,[FuncionarioFerias].[FuncionarioID]
                               ,[Funcionarios].[Empresa]
                               ,[Funcionarios].[Matricula]
                               ,[FuncionarioFerias].[Ferias2AqIni] AS [RefInicial]
                               ,[FuncionarioFerias].[Ferias2AqFim] AS [RefFinal]
                               ,[FuncionarioFerias].[Ferias2Ini]   AS [ConcedInicial]
                               ,DateAdd('d',[FuncionarioFerias].[Ferias2Dias],[FuncionarioFerias].[Ferias2Ini]) AS [ConcedFinal]
                               ,0 AS [Faltas]
                               ,[FuncionarioFerias].[Ferias2Dias] AS [DiasFerias]
                               ,[FuncionarioFerias].[UltPerFerColetivas] AS [FeriasColetivas]
                               ,[FuncionarioFerias].[TeraAbonoNasFerias] AS [AbonoPecuniario]
                               ,[FuncionarioFerias].[Tera13NasFerias] AS [13Sarario]
                               ,0 AS [Notas]
                        FROM [FuncionarioFerias]
                        INNER JOIN [Funcionarios]
                            ON (Funcionarios.FuncionarioID=[FuncionarioFerias].[FuncionarioID])
                        ORDER
                           BY Empresa,Matricula,FuncionarioID,RefInicial
                    #pragma __endtext
                    QRHOpenRecordSet(hOleConn["HistFerias"],hOleConn["SourceConnection"],cSource,"Empresa,Matricula,FuncionarioID,RefInicial")
                end with
            endif
        end

    WAIT CLEAR

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Férias TOTVS Microsiga Protheus...")) NOWAIT

        with object hOleConn["TargetConnection"]
            if (:State==adStateOpen )
                hOleConn["SRFRECNO"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["SRFRECNO"]
                    #pragma __cstream|cSource:=%s
                        SELECT (MAX(SRF.R_E_C_N_O_)+1) SRFRECNO
                          FROM SRF010 SRF
                    #pragma __endtext
                    cSource:=hb_StrReplace(cSource,{"SRF010"=>"SRF"+cTOTVSEmpresa+"0"})
                    QRHOpenRecordSet(hOleConn["SRFRECNO"],hOleConn["TargetConnection"],cSource,"SRFRECNO")
                    if (:eof())
                        nSRFRecNo:=1
                    else
                        nSRFRecNo:=:Fields("SRFRECNO"):Value
                    endif
                    :Close()
                end with
                hOleConn["SRA"]:=TOleAuto():New("ADODB.RecordSet")
                hOleConn["SRF"]:=TOleAuto():New("ADODB.RecordSet")
            endif
        end with

    WAIT CLEAR

    with object hOleConn["SourceConnection"]
        if (:State==adStateOpen )
            with object hOleConn["TargetConnection"]
                if (:State==adStateOpen )
                    CreateProgressBar("Importando "+hb_OemToAnsi(hb_UTF8ToStr("Programação de Férias"))+"...")
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
                                          case "RF_FILIAL"
                                          case "RF_MAT"
                                          case "RF_DATABAS"
                                            xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa)
                                            if (cTargetField=="RF_FILIAL")
                                                cFilial:=xValue
                                            elseif (cTargetField=="RF_MAT")
                                                cMatricula:=xValue
                                            elseif (cTargetField=="RF_DATABAS")
                                                cDataBase:=xValue
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
                                            with object hOleConn["SRF"]
                                                #pragma __cstream|cSource:=%s
                                                    SELECT *
                                                      FROM SRF010 SRF
                                                     WHERE SRF.D_E_L_E_T_=' '
                                                       AND SRF.RF_FILIAL='Filial'
                                                       AND SRF.RF_MAT='Matricula'
                                                       AND SRF.RF_DATABAS='DataBase'
                                                       AND SRF.RF_PD='100'
                                                     ORDER BY SRF.RF_FILIAL
                                                             ,SRF.RF_MAT
                                                             ,SRF.RF_DATABAS
                                                             ,SRF.RF_PD
                                                #pragma __endtext
                                                cSource:=hb_StrReplace(cSource,{"SRF010"=>"SRF"+cTOTVSEmpresa+"0","Filial"=>cFilial,"Matricula"=>cMatricula,"DataBase"=>cDataBase})
                                                QRHOpenRecordSet(hOleConn["SRF"],hOleConn["TargetConnection"],cSource,"RF_FILIAL,RF_MAT,RF_DATABAS,RF_PD")
                                                :Find("RF_MAT='"+cMatricula+"'",0,1)
                                                lAddNew:=(:eof())
                                                if (lAddNew)
                                                    :AddNew()
                                                endif
                                                for each cTargetField in hb_HKeys(hFields)
                                                    lLoop:=.F.
                                                    lRecNo:=(cTargetField=="R_E_C_N_O_")
                                                    if (lRecNo)
                                                        if (lAddNew)
                                                            xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa,@lLoop,"R_E_C_N_O_",nSRFRecNo++)
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

    MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Importação Programação de Férias")))

return
