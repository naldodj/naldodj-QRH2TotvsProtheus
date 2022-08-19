/*
 * MINIGUI - Harbour Win32 GUI Quatra RH 2 TOTVS Protheus
 *
 * (c) 2022 Marinaldo de Jesus <marinaldo.jesus@gmail.com>
 */

#include "minigui.ch"
#include "c:\minigui\source\adordd\adordd.ch"

DECLARE WINDOW Form_QRH2Protheus

procedure main

    local hINI as hash

	SET CENTURY ON

    SET DEFAULT Icon TO GetStartupFolder() + "\QRH2TOTVSProtheus.ico"

    DEFINE WINDOW Form_MainQRH2Protheus ;
        AT 0, 0 ;
        WIDTH 600 HEIGHT 400 ;
        TITLE "Connecti :: Quarta RH To TOTVS Microsiga Protheus " ;
        MAIN ;
        ON INIT hINI:=hb_iniRead("QRH2TOTVSProtheus.ini")
        DEFINE MAIN MENU
            DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("&Opções"))
                DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("&Importação"))
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Funcionários")) ACTION QRHFuncionarios(hINI)
                END POPUP
                SEPARATOR
                DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("&Configurações"))
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Reload")) ACTION (hINI:=hb_iniRead("QRH2TOTVSProtheus.ini"))
                    MENUITEM hb_OemToAnsi(hb_UTF8ToStr("&Show")) ACTION QRH2TOTVSProtheusViewIni(".\QRH2TOTVSProtheus.ini")
                END POPUP
                SEPARATOR
                ITEM 'E&xit' ACTION Form_MainQRH2Protheus.Release()
            END POPUP
            DEFINE POPUP hb_OemToAnsi(hb_UTF8ToStr("&Sobre"))
                ITEM  "&About" ACTION About()
            END POPUP
        END MENU
		DEFINE STATUSBAR FONT "MS Sans serif" SIZE 9 BOLD
			STATUSITEM "Connecti :: Quarta RH To TOTVS Microsiga Protheus " 	 
		END STATUSBAR
    ON KEY ESCAPE ACTION ThisWindow.Release
    END WINDOW

    CENTER WINDOW Form_MainQRH2Protheus
    MAXIMIZE WINDOW Form_MainQRH2Protheus
    ACTIVATE WINDOW Form_MainQRH2Protheus

return

static function CreateProgressBar( cTitle )

   DEFINE WINDOW Form_QRH2Protheus ;
      ROW 0 COL 0 ;
      WIDTH 428 HEIGHT 200 ;
      TITLE cTitle ;
      WINDOWTYPE MODAL ;
      NOSIZE ;
      FONT 'Tahoma' SIZE 11

   @ 10, 80 ANIMATEBOX Avi_1 ;
      WIDTH 260 HEIGHT 40 ;
      FILE 'filecopy.avi' ;
      AUTOPLAY TRANSPARENT NOBORDER

   @ 75, 10 LABEL Label_1 ;
      WIDTH 400 HEIGHT 20 ;
      VALUE ''            ;
      CENTERALIGN VCENTERALIGN

   @ 105, 20 PROGRESSBAR PrgBar_1 ;
      RANGE 0, 100 ;
      VALUE 0      ;
      WIDTH 380 HEIGHT 34

   END WINDOW

   Form_QRH2Protheus.Center
   Form_QRH2Protheus.Closable:=.F.

   Activate Window Form_QRH2Protheus NoWait

return NIL

static function CloseProgressBar()

   IF IsWindowDefined( Form_QRH2Protheus )
      Form_QRH2Protheus.Closable:=.T.
      Form_QRH2Protheus.Release
   ENDIF

   DO MESSAGE LOOP

return NIL

static function About()
    local cAbout as character
    local cCopyRight as character
    cAbout:="QRH2TOTVSProtheus :: "
    cAbout+=MiniGuiVersion()
    cCopyRight:="QRH2TOTVSProtheus"
    cCopyRight+=" 1.0"
    cCopyRight+=CRLF
    cCopyRight+=Chr(169)
    cCopyRight+=" marinaldo.jesus@gmail.com"
return(ShellAbout(cAbout,cCopyRight,LoadTrayIcon(GetInstance(),"MAINICON",50,50)))

#include "QRHFuncionarios.prg"
#include "QRH2TOTVSProtheusViewIni.prg"