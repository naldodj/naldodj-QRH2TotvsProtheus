#include "minigui.ch"

#define SRDUPDATEDATA 500

procedure QRHFuncionariosHistFolhaSRD(hINI as hash)

    local aRDSEQ as array := Array(0)

    local cErrorMsg as character

    local cRDPD as character
    local cRDCC as character
    local cRDSEQ as character
    local cRDRoteir as character
    local cRDDatarq as character

    local cFilial as character
    local cEmpresa as character
    local cMatricula as character
    local cFuncionarioID as character

    local cTOTVSEmpresa as character

    local cSource as character
    local cTargetField as character

    local cCommonFindKey as character

    local hFields as hash := hINI["QRHFuncionariosHistFolhaSRD"]
    local hFieldsSRA as hash := hINI["QRHFuncionarios"]

    local hOleConn as hash := QRHGetProviders(hINI)

    local hSRDData as hash := {=>}

    local lLoop as logical
    local lRecNo as logical
    local lAddNew as logical
    local lIncSeq as logical

    local nRDSEQ as numeric := 0
    local nATRDSEQ as numeric
    local nEmpresa as numeric
    local nMatricula as numeric
    local nFuncionarioID as numeric

    local nSRDConn as numeric := 0
    local nSRDRecNo as numeric

    local nRow as numeric
    local nComplete as numeric

    local oError as object
    local oSRDOleData as object

    local xValue

    cTOTVSEmpresa:=QRH2TotvsProtheusGetEmpresa(hINI)

    if (empty(cTOTVSEmpresa))
        MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Empresa Inválida")))
        return
    endif

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("HistFolha Quarta RH...")) NOWAIT

        with object hOleConn["SourceConnection"]
            if (:State==adStateOpen )
                hOleConn["Funcionarios"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["Funcionarios"]
                    #pragma __cstream|cSource:=%s
                        SELECT * FROM Funcionarios ORDER BY Empresa,Matricula,FuncionarioID
                    #pragma __endtext
                    QRHOpenRecordSet(hOleConn["Funcionarios"],hOleConn["SourceConnection"],cSource,"Empresa,Matricula,FuncionarioID")
                end with
                hOleConn["HistFolha"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["HistFolha"]
                    #pragma __cstream|cSource:=%s
                        SELECT HistFolha.*
                              ,Eventos.Tipo
                          FROM HistFolha
                     LEFT JOIN Eventos ON (HistFolha.Codigo=Eventos.Codigo)
                      ORDER BY HistFolha.FuncionarioID,HistFolha.TipoFolha,HistFolha.DataCalculo,HistFolha.Codigo
                    #pragma __endtext
                    QRHOpenRecordSet(hOleConn["HistFolha"],hOleConn["SourceConnection"],cSource,"FuncionarioID,TipoFolha,DataCalculo,Codigo")
                end with
            endif
        end

    WAIT CLEAR

    WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("HistFolha TOTVS Microsiga Protheus...")) NOWAIT

        with object hOleConn["TargetConnection"]
            if (:State==adStateOpen )
                hOleConn["SRDRECNO"]:=TOleAuto():New("ADODB.RecordSet")
                with object hOleConn["SRDRECNO"]
                    #pragma __cstream|cSource:=%s
                        SELECT (MAX(SRD.R_E_C_N_O_)+1) SRDRECNO
                          FROM SRD010 SRD
                    #pragma __endtext
                    cSource:=hb_StrReplace(cSource,{"SRD010"=>"SRD"+cTOTVSEmpresa+"0"})
                    QRHOpenRecordSet(hOleConn["SRDRECNO"],hOleConn["TargetConnection"],cSource,"SRDRECNO")
                    if (:eof())
                        nSRDRecNo:=1
                    else
                        nSRDRecNo:=:Fields("SRDRECNO"):Value
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
                    CreateProgressBar("Importando "+hb_OemToAnsi(hb_UTF8ToStr("HistFolha"))+"...")
                    with object hOleConn["Funcionarios"]
                        nRow:=0
                        :MoveFirst()
                        while (!:eof())
                            nRow++
                            aSize(aRDSEQ,0)
                            nEmpresa:=:Fields("Empresa"):Value
                            cEmpresa:=hb_NToS(nEmpresa)
                            nMatricula:=:Fields("Matricula"):Value
                            cMatricula:=(nMatricula)
                            nFuncionarioID:=:Fields("FuncionarioID"):Value
                            cFuncionarioID:=hb_NToS(nFuncionarioID)
                            cCommonFindKey:="FuncionarioID="+cFuncionarioID
                            cFilial:=getTargetFieldValue(hIni,"RA_FILIAL",hFieldsSRA,hOleConn)
                            cMatricula:=getTargetFieldValue(hIni,"RA_MAT",hFieldsSRA,hOleConn)
                            with object hOleConn["HistFolha"]
                                :MoveFirst()
                                :Find(cCommonFindKey,0,1)
                                while (!:eof())
                                    if (nFuncionarioID!=:Fields("FuncionarioID"):Value)
                                        exit
                                    endif
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
                                        cSource:=hb_StrReplace(cSource,;
                                            {;
                                                "SRA010"=>"SRA"+cTOTVSEmpresa+"0",;
                                                "Filial"=>cFilial,;
                                                "Matricula"=>cMatricula;
                                            };
                                        )
                                        QRHOpenRecordSet(hOleConn["SRA"],hOleConn["TargetConnection"],cSource,"RA_FILIAL,RA_MAT")
                                        :Find("RA_MAT='"+cMatricula+"'",0,1)
                                        if (!:eof())
                                            for each cTargetField in hb_HKeys(hFields)
                                                switch cTargetField
                                                  case "RD_FILIAL"
                                                  case "RD_MAT"
                                                  case "RD_DATARQ"
                                                  case "RD_PD"
                                                  case "RD_CC"
                                                  case "RD_ROTEIR"
                                                    xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa)
                                                    if (cTargetField=="RD_FILIAL")
                                                        cFilial:=xValue
                                                    elseif (cTargetField=="RD_MAT")
                                                        cMatricula:=xValue
                                                    elseif (cTargetField=="RD_DATARQ")
                                                        cRDDatarq:=xValue
                                                    elseif (cTargetField=="RD_PD")
                                                        cRDPD:=xValue
                                                    elseif (cTargetField=="RD_CC")
                                                        cRDCC:=xValue
                                                    elseif (cTargetField=="RD_ROTEIR")
                                                        cRDRoteir:=xValue
                                                    endif
                                                end switch
                                            next each
                                            if (!Empty(cRDPD))
                                                nATRDSEQ:=aScan(aRDSEQ,;
                                                    {|x|;
                                                         (x[2]==cFilial);
                                                    .and.(x[3]==cMatricula);
                                                    .and.(x[4]==cRDDatarq);
                                                    .and.(x[5]==cRDPD);
                                                    .and.(x[6]==cRDCC);
                                                    .and.(x[7]=cRDRoteir);
                                                    };
                                                )
                                                if (nATRDSEQ==0)
                                                    aAdd(aRDSEQ,{1,cFilial,cMatricula,cRDDatarq,cRDPD,cRDCC,cRDRoteir})
                                                    nATRDSEQ:=Len(aRDSEQ)
                                                else
                                                    aRDSEQ[nATRDSEQ][1]++
                                                endif
                                                nRDSEQ:=aRDSEQ[nATRDSEQ][1]
                                                cRDSEQ:=getTargetFieldValue(hIni,"RD_SEQ",hFields,hOleConn,cFilial,cMatricula,cEmpresa,nil,"RD_SEQ",nRDSEQ)
                                                hOleConn["SRD"]:=TOleAuto():New("ADODB.RecordSet")
                                                with object hOleConn["SRD"]
                                                    #pragma __cstream|cSource:=%s
                                                        SELECT *
                                                          FROM SRD010 SRD
                                                         WHERE SRD.RD_FILIAL='Filial'
                                                           AND SRD.RD_MAT='Matricula'
                                                           AND SRD.RD_DATARQ='cRDDatarq'
                                                           AND SRD.RD_ROTEIR='cRDRoteir'
                                                           AND SRD.RD_PD='cRDPD'
                                                           AND SRD.RD_CC='cRDCC'
                                                           AND SRD.RD_SEQ='cRDSEQ'
                                                      ORDER BY SRD.RD_FILIAL
                                                              ,SRD.RD_MAT
                                                              ,SRD.RD_DATARQ
                                                              ,SRD.RD_ROTEIR
                                                              ,SRD.RD_PD
                                                              ,SRD.RD_CC
                                                              ,SRD.RD_SEQ
                                                    #pragma __endtext
                                                    cSource:=hb_StrReplace(cSource,;
                                                        {;
                                                            "SRD010"=>"SRD"+cTOTVSEmpresa+"0",;
                                                            "Filial"=>cFilial,;
                                                            "Matricula"=>cMatricula,;
                                                            "cRDDatarq"=>cRDDatarq,;
                                                            "cRDRoteir"=>cRDRoteir,;
                                                            "cRDPD"=>cRDPD,;
                                                            "cRDCC"=>cRDCC,;
                                                            "cRDSEQ"=>cRDSEQ;
                                                        };
                                                    )
                                                    QRHOpenRecordSet(hOleConn["SRD"],hOleConn["TargetConnection"],cSource,"RD_FILIAL,RD_MAT,RD_DATARQ,RD_ROTEIR,RD_PD,RD_CC,RD_SEQ")
                                                    :Find("RD_MAT='"+cMatricula+"'",0,1)
                                                    lAddNew:=(:eof())
                                                    if (lAddNew)
                                                        :AddNew()
                                                    endif
                                                    for each cTargetField in hb_HKeys(hFields)
                                                        lLoop:=.F.
                                                        lRecNo:=(cTargetField=="R_E_C_N_O_")
                                                        if (lRecNo)
                                                            if (lAddNew)
                                                                xValue:=getTargetFieldValue(hIni,cTargetField,hFields,hOleConn,cFilial,cMatricula,cEmpresa,@lLoop,"R_E_C_N_O_",nSRDRecNo++)
                                                            endif
                                                        elseif (cTargetField=="RD_SEQ")
                                                            xValue:=cRDSEQ
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
                                                    nSRDConn++
                                                    if (nSRDConn>SRDUPDATEDATA)
                                                        WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Update HistFolha TOTVS Protheus...")) NOWAIT
                                                            for each oSRDOleData in hSRDData
                                                                with object oSRDOleData
                                                                    if (:State==adStateOpen )
                                                                        :Update()
                                                                        :Close()
                                                                    endif
                                                                end whith
                                                            next each
                                                            for nSRDConn:=1 to Len(hSRDData)
                                                                if (hb_HHasKey(hSRDData,nSRDConn))
                                                                    hb_hDel(hSRDData,nSRDConn)
                                                                endif
                                                            next nSRDConn
                                                            nSRDConn:=1
                                                        WAIT CLEAR
                                                    endif
                                                    hSRDData[nSRDConn]:=hOleConn["SRD"]
                                                end whith
                                            endif
                                        endif
                                        :Close()
                                    end whith
                                    :MoveNext()
                                    // refreshing
                                    InkeyGui()
                                end while
                            end whith
                            if (nSRDConn>0)
                                WAIT WINDOW hb_OemToAnsi(hb_UTF8ToStr("Update HistFolha TOTVS Protheus...")) NOWAIT
                                    for each oSRDOleData in hSRDData
                                        with object oSRDOleData
                                            if (:State==adStateOpen )
                                                :Update()
                                                :Close()
                                            endif
                                        end whith
                                    next each
                                    for nSRDConn:=1 to Len(hSRDData)
                                        if (hb_HHasKey(hSRDData,nSRDConn))
                                            hb_hDel(hSRDData,nSRDConn)
                                        endif
                                    next nSRDConn
                                WAIT CLEAR
                                nSRDConn:=0
                            endif
                            nComplete:=Int((nRow/:RecordCount)*100)
                            if (Mod(nComplete,10)==0)
                                if (IsWindowDefined(Form_QRH2Protheus))
                                    Form_QRH2Protheus.PrgBar_1.Value:=nComplete
                                    Form_QRH2Protheus.Label_1.Value:=hb_StrReplace("Completed [nRow/:RecordCount]("+hb_NToS(nComplete)+")%",{"nRow"=>hb_NToS(nRow),":RecordCount"=>hb_NToS(:RecordCount)})
                                else
                                    exit
                                endif
                            endif
                            :MoveNext()
                            // refreshing
                            InkeyGui()
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

    MsgInfo(hb_OemToAnsi(hb_UTF8ToStr("Importação HistFolha Finalizada")))

return
