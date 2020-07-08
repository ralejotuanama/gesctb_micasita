VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.MDIForm frm_MnuPri_01 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7695
   ClientLeft      =   3000
   ClientTop       =   3660
   ClientWidth     =   16605
   Icon            =   "GesCtb_frm_002.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16605
      _Version        =   65536
      _ExtentX        =   29289
      _ExtentY        =   1138
      _StockProps     =   15
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.CommandButton cmd_CamCon 
         Height          =   585
         Left            =   630
         Picture         =   "GesCtb_frm_002.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cambio de Contraseña"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   30
         Picture         =   "GesCtb_frm_002.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir de Plataforma"
         Top             =   30
         Width           =   585
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   7305
      Width           =   16605
      _Version        =   65536
      _ExtentX        =   29289
      _ExtentY        =   688
      _StockProps     =   15
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   1
      Begin Threed.SSPanel pnl_EntDat 
         Height          =   315
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   3900
         _Version        =   65536
         _ExtentX        =   6879
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "lm_db_db1 - prod1"
         ForeColor       =   32768
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Font3D          =   2
      End
      Begin Threed.SSPanel pnl_NumVer 
         Height          =   315
         Left            =   3960
         TabIndex        =   5
         Top             =   30
         Width           =   2100
         _Version        =   65536
         _ExtentX        =   3704
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "rev. 008-1028.1"
         ForeColor       =   32768
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Font3D          =   2
      End
      Begin Threed.SSPanel pnl_TipCam 
         Height          =   315
         Left            =   6090
         TabIndex        =   6
         Top             =   30
         Width           =   5400
         _Version        =   65536
         _ExtentX        =   9525
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Tipo Cambio Operativo: Compra: S/. 2.00 - Venta: S/. 2.01"
         ForeColor       =   32768
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Font3D          =   2
      End
   End
   Begin VB.Menu mnuMnt 
      Caption         =   "&Mantenimientos"
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Tipo de Cambio"
         Index           =   1
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Períodos"
         Index           =   3
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Clase de Créditos"
         Index           =   5
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Tipo de Créditos"
         Index           =   6
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Naturaleza de Créditos"
         Index           =   7
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Clasificación por Clase de Crédito"
         Index           =   8
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Situación por Clase de Crédito"
         Index           =   9
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Clase de Garantías"
         Index           =   11
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Tipos de Garantías"
         Index           =   12
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Garantías"
         Index           =   13
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Sectores Económicos"
         Index           =   14
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "CIIU"
         Index           =   15
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Provisiones"
         Index           =   16
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Clase de Cuenta Contable"
         Index           =   17
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Tipo de Monedas"
         Index           =   18
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Parámetros por Empresa"
         Index           =   19
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Plan de Cuentas"
         Index           =   20
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "ITF"
         Index           =   21
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Libros Contables"
         Index           =   22
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Bancos"
         Index           =   23
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Empresas Supervisadas SBS"
         Index           =   24
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Créditos Comerciales"
         Index           =   25
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Capital y Reserva Legal / Patrimonio Efectivo"
         Index           =   26
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "EEFF - Estado de Ganancias y Perdidas"
         Index           =   27
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "EEFF - Balance General"
         Index           =   28
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "-"
         Index           =   29
      End
      Begin VB.Menu mnuMnt_Opcion 
         Caption         =   "Desembolso al Promotor"
         Index           =   30
      End
   End
   Begin VB.Menu mnuCon 
      Caption         =   "&Consultas"
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "Consulta de Tipo de Cambio"
         Index           =   1
      End
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "Consulta de Operaciones Financieras"
         Index           =   3
      End
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "Consulta de Crédito Hipotecario"
         Index           =   5
      End
   End
   Begin VB.Menu mnuPro 
      Caption         =   "&Procesos Cierre"
      Begin VB.Menu mnuPro_Opcion 
         Caption         =   "Carga Archivo RCC (1)"
         Index           =   1
      End
      Begin VB.Menu mnuPro_Opcion 
         Caption         =   "Asientos Diferidos (Cred. Hipot.) (2)"
         Index           =   2
      End
      Begin VB.Menu mnuPro_Opcion 
         Caption         =   "Cierre de Créditos Hipotecarios (3)"
         Index           =   3
      End
      Begin VB.Menu mnuPro_Opcion 
         Caption         =   "Asientos Devengados (Cred. Hipot.) (4)"
         Index           =   4
      End
      Begin VB.Menu mnuPro_Opcion 
         Caption         =   "Asientos Varios (5)"
         Index           =   5
      End
      Begin VB.Menu mnuPro_Opcion 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuPro_Opcion 
         Caption         =   "Balance General"
         Index           =   7
      End
      Begin VB.Menu mnuPro_Opcion 
         Caption         =   "Cierre del Ejercicio"
         Index           =   8
      End
      Begin VB.Menu mnuPro_Opcion 
         Caption         =   "Anexo 7-16-16b (Proceso)"
         Index           =   9
      End
      Begin VB.Menu mnuPro_Opcion 
         Caption         =   "Proceso de Limites Globales"
         Index           =   10
      End
      Begin VB.Menu mnuPro_Opcion 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuPro_Opcion 
         Caption         =   "Recalculo de Provisiones"
         Index           =   12
      End
      Begin VB.Menu mnuPro_Opcion 
         Caption         =   "Diferencia de Saldos"
         Index           =   13
      End
   End
   Begin VB.Menu mnuAsi 
      Caption         =   "&Asientos Contables"
      Begin VB.Menu mnuAsi_Opcion 
         Caption         =   "Asientos Manuales"
         Index           =   1
      End
      Begin VB.Menu mnuAsi_Opcion 
         Caption         =   "Consulta de Asiento Contable"
         Index           =   2
      End
   End
   Begin VB.Menu mnuMat 
      Caption         =   "Ma&trices Contables"
      Begin VB.Menu mnuMat_Opcion 
         Caption         =   "Dinámicas Contables"
         Index           =   1
      End
      Begin VB.Menu mnuMat_Opcion 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuMat_Opcion 
         Caption         =   "Conceptos Contables"
         Index           =   3
      End
      Begin VB.Menu mnuMat_Opcion 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuMat_Opcion 
         Caption         =   "Cuentas Contables para Productos"
         Index           =   5
      End
      Begin VB.Menu mnuMat_Opcion 
         Caption         =   "Cuentas Contables para Proyectos Hipotecarios"
         Index           =   6
      End
      Begin VB.Menu mnuMat_Opcion 
         Caption         =   "Cuentas Contables para Cuentas Bancarias"
         Index           =   7
      End
   End
   Begin VB.Menu mnuEdp 
      Caption         =   "Asientos &Edpymebank"
      Begin VB.Menu mnuEdp_Opcion 
         Caption         =   "Carga Masiva de Asientos Contables"
         Index           =   1
      End
      Begin VB.Menu mnuEdp_Opcion 
         Caption         =   "Nivelación Diferidos"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEdp_Opcion 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuEdp_Opcion 
         Caption         =   "Contabilización de Plan de Ahorros"
         Index           =   4
      End
      Begin VB.Menu mnuEdp_Opcion 
         Caption         =   "Contabilización de Pagos de Gastos de Cierre (Cred. Hipot.)"
         Index           =   5
      End
      Begin VB.Menu mnuEdp_Opcion 
         Caption         =   "Contabilización de Desembolsos (Cred. Hipot.)"
         Index           =   6
      End
      Begin VB.Menu mnuEdp_Opcion 
         Caption         =   "Contabilización de Pagos de Cuotas (Cred. Hipot.)"
         Index           =   7
      End
      Begin VB.Menu mnuEdp_Opcion 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuEdp_Opcion 
         Caption         =   "Contabilización de Pre-Pagos (Cred. Hipot.)"
         Index           =   9
      End
      Begin VB.Menu mnuEdp_Opcion 
         Caption         =   "Contabilizacion de Interes PBP (Cred. Hipot.)"
         Index           =   10
      End
      Begin VB.Menu mnuEdp_Opcion 
         Caption         =   "Contabilizacion de Asignación PBP (Cred. Hipot.)"
         Index           =   11
      End
      Begin VB.Menu mnuEdp_Opcion 
         Caption         =   "Contabilizacion de Provisiones (Cred. Hipot.)"
         Index           =   12
      End
      Begin VB.Menu mnuEdp_Opcion 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuEdp_Opcion 
         Caption         =   "Contabilizacion Desembolsos Cofide"
         Index           =   14
      End
      Begin VB.Menu mnuEdp_Opcion 
         Caption         =   "Contabilizacion Desembolsos BBP"
         Index           =   15
      End
      Begin VB.Menu mnuEdp_Opcion 
         Caption         =   "Contabilizacion Interes por Pagar Cofide"
         Index           =   16
      End
   End
   Begin VB.Menu mnuRop 
      Caption         =   "&Reportes Operativos"
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "Créditos Hipotecarios Desembolsados (Mensual)"
         Index           =   1
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "Saldos de Créditos Hipotecarios (Mensual)"
         Index           =   2
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "Generación de Balances"
         Index           =   4
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "Padrón de Deudores"
         Index           =   5
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "Cuentas por Pagar (COFIDE)"
         Index           =   7
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "Saldos por Pagar (COFIDE)"
         Index           =   8
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "Consolidado de Clasificación de Cartera"
         Index           =   10
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "Consolidado de Cartera en Riesgo"
         Index           =   11
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "Provisión de clientes Morosos y Alineados"
         Index           =   12
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "Resumen de Provisiones"
         Index           =   13
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "EEFF - Estados de Ganancias y Pérdidas"
         Index           =   14
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "EEFF - Balance General"
         Index           =   15
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "EEFF - Origen/Aplicacion de Balance"
         Index           =   16
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "Reporte de Cobranzas a Clientes Morosos y Alineados"
         Index           =   18
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "Reporte de Morosidad - Cartera Atrasada"
         Index           =   19
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "Reporte de Indicadores Varios"
         Index           =   21
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "Reporte de Cuotas por Cobrar y Pagar"
         Index           =   22
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "-"
         Index           =   23
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "Carga de archivos"
         Index           =   24
      End
      Begin VB.Menu mnuRop_Opcion 
         Caption         =   "Conciliación - Contable/Operativa"
         Index           =   25
      End
   End
   Begin VB.Menu mnuRsu 
      Caption         =   "Reportes &Sunat"
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "0695 Impuesto a las Transacciones Financieras (ITF)"
         Index           =   1
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "F01.2: ""LIBRO CAJA Y BANCOS"" - Detalle de los Movimientos de la Cuenta Corriente"
         Index           =   3
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "F03.4:   Cuentas por Cobrar a Accionistas (o socios) y Personal"
         Index           =   5
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "F03.5:   Cuentas por Cobrar Diversas"
         Index           =   6
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "F03.6:   Provisión para cuentas de cobranza dudosa"
         Index           =   7
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "F03.9:   Intangibles"
         Index           =   8
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "F03.11: Remuneraciones por Pagar"
         Index           =   9
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "F03.12: Proveedores"
         Index           =   10
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "F03.13: Cuentas por Pagar Diversas"
         Index           =   11
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "F03.14: Beneficios Sociales de los Trabajadores"
         Index           =   12
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "F03.15: Ganancias Diferidas"
         Index           =   13
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "F03.16: Capital"
         Index           =   14
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "F05.1: LIBRO DIARIO"
         Index           =   16
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "F06.1: LIBRO MAYOR"
         Index           =   17
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "F07.1: Registro de Activos Fijos - Detalle de los Activos Fijos"""
         Index           =   18
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "F08.1: Registro de Compras"
         Index           =   19
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "F14.1: Registro de Ventas e Ingresos"
         Index           =   20
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "-"
         Index           =   21
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "Facturación Electronica"
         Index           =   22
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "Facturador Manual - Registro"
         Index           =   23
      End
      Begin VB.Menu mnuRsu_Opcion 
         Caption         =   "Facturador Manual - Aprobación"
         Index           =   24
      End
   End
   Begin VB.Menu mnuCxP 
      Caption         =   "Cuentas por Pagar"
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Maestro de Proveedores"
         Index           =   1
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Registro de Compras"
         Index           =   2
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Registro de Ventas"
         Index           =   3
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Modulo Caja Chica"
         Index           =   5
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Modulo Entregas a Rendir"
         Index           =   6
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Modulo Tarjeta de Crédito"
         Index           =   7
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Modulo Cuentas por Pagar"
         Index           =   8
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Módulo Gestión Personal"
         Index           =   10
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Generación Data Planilla"
         Index           =   11
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Gestión de Vacaciones"
         Index           =   12
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Autorización de Vacaciones"
         Index           =   13
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Dinamica Contable DPF"
         Index           =   15
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Módulo Inversiones DPF"
         Index           =   16
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Módulo de Transferencias"
         Index           =   18
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Modulo de Carga Archivos de Recaudo"
         Index           =   19
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Modulo de Autorización"
         Index           =   21
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Modulo de Compensación"
         Index           =   22
      End
      Begin VB.Menu mnuCxP_Opcion 
         Caption         =   "Reporte _Ultimo"
         Index           =   23
      End
   End
End
Attribute VB_Name = "frm_MnuPri_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_CamCon_Click()
   frm_IdeUsu_02.Show 1
End Sub

Private Sub cmd_Salida_Click()
   If MsgBox("¿Está seguro de salir de la Plataforma?", vbExclamation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      Call gs_Desconecta_Servidor
      End
   End If
End Sub

Private Sub MDIForm_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
'   Call fs_HabSeg
   Screen.MousePointer = 0
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   If MsgBox("¿Está seguro de salir de la Plataforma?", vbExclamation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      Call gs_Desconecta_Servidor
      End
   Else
      Cancel = True
   End If
End Sub

Private Sub mnuCxP_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1:  frm_Ctb_RegCom_01.Show 1      'Registro de proveedores
      Case 2:  frm_Ctb_RegCom_03.Show 1      'Registro de compras
      Case 3:  frm_Ctb_RegVen_01.Show 1      'Registro de ventas
      Case 5:  frm_Ctb_CajChc_01.Show 1      'Registro de caja chica
      Case 6:  frm_Ctb_EntRen_01.Show 1      'Modulo de Entregas a Rendir
      Case 7:  frm_Ctb_TarCre_02.Show 1      'Modulo de Tarjetas de Credito
      Case 8:  frm_Ctb_CtaPag_01.Show 1      'Modulo de Cuentas por Pagar
      Case 10: moddat_g_int_TipRec = 1       'Gestión Personal
               frm_Ctb_GesPer_01.Show 1
      Case 11: frm_Ctb_RegCom_05.Show 1      'Modulo de Generacion de data para planillas
      Case 12: moddat_g_int_TipRec = 2       'Gestión Vacaciones
               frm_Ctb_GesPer_01.Show 1
      Case 13: moddat_g_int_TipRec = 3       'Autorizacion de vacaciones
               frm_Ctb_GesPer_01.Show 1
      Case 15: frm_Ctb_InvDpf_04.Show 1      'Dinamica contable DPF
      Case 16: frm_Ctb_InvDpf_01.Show 1      'Inversiones DPF
      Case 18: frm_Ctb_TrnCta_02.Show 1      'Modulo de Transferencias
      Case 19: frm_Ctb_CarArc_01.Show 1      'Modulo de Carga Archivo Recaudo
      Case 21: frm_Ctb_PagCom_05.Show 1      'Modulo de autorizaciones
      Case 22: frm_Ctb_PagCom_01.Show 1      'Modulo de compensaciones
      Case 23: frm_Pruebas_01.Show 1         'Modulo de Pruebas
   End Select
End Sub

Private Sub mnuMnt_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1:  frm_TipCam_01.Show 1          'Tipo de Cambio
      Case 3:  frm_Mnt_Period_01.Show 1      'Periodos
      Case 5:  frm_Mnt_TipCre_01.Show 1      'Clase de Creditos
      Case 6:  frm_Mnt_TipCre_03.Show 1      'Tipo de Creditos
      Case 7:  frm_Mnt_NatCre_01.Show 1      'Naturaleza de Creditos
      Case 8:  frm_Mnt_ClaCre_01.Show 1      'Clasificacion de Creditos
      Case 9:  frm_Mnt_SitCre_01.Show 1      'Situacion de Creditos
      Case 11: frm_Mnt_ClaGar_01.Show 1      'Clase de Garantias
      Case 12: frm_Mnt_TipGar_01.Show 1      'Tipo de Garantias
      Case 13: frm_Mnt_DetGar_01.Show 1      'Garantias
      Case 14: frm_Mnt_SecEco_01.Show 1      'Sector Economico
      Case 15: frm_Mnt_CodCiu_01.Show 1      'Ciiu
      Case 16: frm_Mnt_Provis_01.Show 1      'Provisiones
      Case 17: frm_Mnt_ClaCta_01.Show 1      'Clase de cuentas contables
      Case 18: frm_Mnt_TipMon_01.Show 1      'Tipos de Moneda
      Case 19: frm_Mnt_ParEmp_01.Show 1      'Parametros por empresa
      Case 20: frm_Mnt_PlaCta_01.Show 1      'Plan Contable
      Case 21: frm_Mnt_PorItf_01.Show 1      'ITF
      Case 22: frm_Mnt_LibCon_01.Show 1      'Libros Contables
      Case 23: frm_Mnt_Bancos_01.Show 1      'Bancos
      Case 24: frm_Mnt_EmpSup_01.Show 1      'Entidades Supervisados por SBS
      Case 25: frm_Mnt_ComCie_01.Show 1      'Creditos Comerciales
      Case 26: frm_Mnt_ConLim_01.Show 1      'Capital y Reserva Legal / Patrimonio Efectivo
      Case 27: frm_Mnt_EFGP_01.Show 1        'Mantenimiento Ganancias y Perdidas
      Case 28: frm_Mnt_EFBG_01.Show 1        'Mantenimiento Balance General
      Case 30: frm_RegDes_01.Show 1          'Desembolso al promotor
   End Select
End Sub

Private Sub mnuCon_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1: frm_Con_TipCam_01.Show 1       'Consulta de Tipo de Cambio
      Case 3: frm_Con_OpeFin_01.Show 1       'Consulta de Operaciones Financieras
      Case 5: frm_Con_CreHip_01.Show 1       'Consulta de Créditos Hipotecarios
   End Select
End Sub

Private Sub mnuPro_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1:  frm_Pro_ArcRCC_01.Show 1      'Carga RCC
      Case 2:  frm_Pro_AsiDif_01.Show 1      'Asientos Diferidos
      Case 3:  frm_Pro_CieCre_01.Show 1      'Cierre Credito
      Case 4:  frm_Pro_AsiDev_01.Show 1      'Asientos Devengados
      Case 5:  frm_Pro_AsiAtr_01.Show 1      'Asientos Varios
     'Case 6:  frm_Pro_BalCom_01.Show 1      'Balance General
     'Case 7:  frm_Pro_CieEje_01.Show 1      'Cierre Ejercicio
     'Case 8:  frm_Pro_CuoHip_01.Show 1      'Anexo 7-16-16b
     'Case 10: frm_Pro_LimGlo_01.Show 1      'Limites Globales
      Case 12: frm_Pro_RecPrv_01.Show 1      'Recalculo de Provisiones
      Case 13: frm_Pro_SdoCie_01.Show 1      'Diferencias en saldos cierre
   End Select
End Sub

Private Sub mnuAsi_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1:
         moddat_g_str_CodIte = "1"
         frm_Ctb_AsiCtb_01.Show 1            'Asientos Registro
      Case 2:
         moddat_g_str_CodIte = "2"
         frm_Ctb_AsiCtb_01.Show 1            'Asientos Consulta
   End Select
End Sub

Private Sub mnuMat_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1: frm_Mat_MatCon_01.Show 1       'Matriz Contable
      Case 3: frm_Mat_ConCtb_01.Show 1       'Conceptos Contables
      Case 5: frm_Mat_Produc_01.Show 1       'Cuentas Contables por Producto
      Case 6: frm_Mat_CtaPry_01.Show 1       'Cuentas Contables por Proyecto Hipotecario
      Case 7: frm_Mat_CtaBco_01.Show 1       'Cuentas Contables por Cuentas Bancarias
   End Select
End Sub

Private Sub mnuEdp_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1:  frm_Ctb_AsiCtb_05.Show 1      'Carga de Asientos Masivos
     'Case 2:  frm_Pro_NivDif_01.Show 1      'Especial
      Case 4:  frm_Pro_CtbAho_01.Show 1      'Proceso de Contabilizacion de Plan de Ahorro
      Case 5:  frm_Pro_CtbGas_01.Show 1      'Proceso de Contabilizacion de Gastos de Cierre de Creditos Hipotecarios
      Case 6:  frm_Pro_CtbDes_01.Show 1      'Proceso de Contabilizacion de Desembolsos
      Case 7:  frm_Pro_CtbCuo_01.Show 1      'Proceso de Contabilizacion de Cuotas de Creditos Hipotecarios
      Case 9:  frm_Pro_CtbPpg_01.Show 1      'Proceso de Contabilizacion de Prepagos de Creditos Hipotecarios
      Case 10: frm_Pro_CtbIntPBP_01.Show 1   'Proceso de Contabilizacion de Provision del Interés PBP
      Case 11: frm_Pro_CtbPbp_01.Show 1      'Proceso de Contabilizacion de Asignación PBP
      Case 12: frm_Pro_CtbPrv_01.Show 1      'Proceso de Contabilizacion de Provisiones PBP
      Case 14: frm_Pro_CtbCof_01.Show 1      'Proceso de Contabilizacion de Desembolsos COFIDE
      Case 15: frm_Pro_CtbBbp_01.Show 1      'Proceso de Contabilizacion de Desembolsos BBP
      Case 16: frm_Pro_CtbIntCof_01.Show 1   'Proceso de Contabilizacion de Interes por Pagar COFIDE
   End Select
End Sub

Private Sub mnuRop_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1:  frm_RptCtb_01.Show 1          'Reporte de Creditos Desembolsados
      Case 2:  frm_RptCtb_02.Show 1          'Reporte de Saldos de Creditos Hipotecarios
      Case 4:  frm_RptCtb_08.Show 1          'Balance de Comprobacion SBS
      Case 5:  frm_RptCtb_12.Show 1          'Padron de Deudores
      Case 7:  frm_RptCtb_09.Show 1          'Reporte de Detalle de Cuentas x Pagar
      Case 8:  frm_RptCtb_10.Show 1          'Reporte de Saldos de Cuentas x Pagar
      Case 10: frm_RptCtb_24.Show 1          'Consolidado de Clasificaciones
      Case 11: frm_RptCtb_29.Show 1          'Consolidado de Cartera en Riesgo
      Case 12: frm_RptCtb_13.Show 1          'Provisión de clientes Morosos y Alineados
      Case 13: frm_RptCtb_27.Show 1          'Resumen de Provisiones
      Case 14: frm_RptCtb_19.Show 1          'Estados Ganancias y Perdidas
      Case 15: frm_RptCtb_22.Show 1          'Balance General
      Case 16: frm_RptCtb_30.Show 1          'Variacion de Balance
      Case 18: frm_RptCtb_26.Show 1          'Reporte de Cobranzas a clientes morosos y alienados
      Case 19: frm_RptCtb_28.Show 1          'Reporte de Morosidad (Cartera Atrasada)
      Case 21: frm_RptCtb_32.Show 1          'Reporte de Ratio de Capital Ajustado (CAR)
      Case 22: frm_RptCtb_33.Show 1          'Reporte de Cuentas por cobrar y pagar
      Case 24: frm_RptCtb_34.Show 1          'Carga de archivos (extracto bancario)
      Case 25: frm_RptCtb_35.Show 1          'Reporte de conciliacion banco/operativo
    End Select
End Sub

Private Sub mnuRsu_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1:  frm_RptCtb_04.Show 1          'Generacion de Archivo ITF para SUNAT
      Case 3:  frm_RptSun_04.Show 1          'Libro Caja Bancos
      Case 5:  frm_RptSun_03.Show 1          'Cuentas por Cobrar a accionistas (o socios) y personal
      Case 16: frm_RptSun_02.Show 1          'Reporte de Libro Diario
      Case 17: frm_RptSun_01.Show 1          'Reporte de Libro Mayor
      Case 19: frm_RptSun_06.Show 1          'Registro de Compras
      Case 20: frm_RptSun_05.Show 1          'Registro de Ventas e Ingresos
      Case 22: frm_RptSun_07.Show 1          'Facturacion Electronica
      Case 23: Frm_Ctb_FacEle_01.Show 1      'Facturador manual - Registro
      Case 24: Frm_Ctb_FacEle_02.Show 1      'Facturador manual - Aprobacion
   End Select
End Sub

Private Sub fs_HabSeg()
Dim r_int_Posici     As Integer
Dim r_str_CodMen     As String
Dim r_dbl_TipVta     As Double
Dim r_dbl_TipCom     As Double
   
   'pnl_Seg_NomUsu.Caption = modgen_g_str_CodUsu
   pnl_NumVer.Caption = modgen_g_str_NumRev
   pnl_EntDat.Caption = moddat_g_str_NomEsq & " - " & UCase(moddat_g_str_EntDat)
   r_dbl_TipVta = moddat_gf_ObtieneTipCamDia(1, 2, Format(date, "yyyymmdd"), 1)
   r_dbl_TipCom = moddat_gf_ObtieneTipCamDia(1, 2, Format(date, "yyyymmdd"), 2)
   pnl_TipCam.Caption = "Tipo de Cambio: Compra: S/. " & Format(r_dbl_TipCom, "###0.0000") & " - Venta: S/. " & Format(r_dbl_TipVta, "###0.0000")
   
   'MENU MANTENIMIENTO
   For r_int_Posici = 1 To mnuMnt_Opcion.Count
      If mnuMnt_Opcion(r_int_Posici).Caption <> "-" Then
         mnuMnt_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   'MENU CONSULTAS
   For r_int_Posici = 1 To mnuCon_Opcion.Count
      If mnuCon_Opcion(r_int_Posici).Caption <> "-" Then
         mnuCon_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   'MENU PROCESOS
   For r_int_Posici = 1 To mnuPro_Opcion.Count
      If mnuPro_Opcion(r_int_Posici).Caption <> "-" Then
         mnuPro_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   'MENU ASIENTOS CONTABLES
   For r_int_Posici = 1 To mnuAsi_Opcion.Count
      If mnuAsi_Opcion(r_int_Posici).Caption <> "-" Then
         mnuAsi_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   'MENU MATRICES CONTABLES
   For r_int_Posici = 1 To mnuMat_Opcion.Count
      If mnuMat_Opcion(r_int_Posici).Caption <> "-" Then
         mnuMat_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   'MENU ASIENTOS EDPYMEBANK
   For r_int_Posici = 1 To mnuEdp_Opcion.Count
      If mnuEdp_Opcion(r_int_Posici).Caption <> "-" Then
         mnuEdp_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   'MENU REPORTES OPERATIVOS
   For r_int_Posici = 1 To mnuRop_Opcion.Count
      If mnuRop_Opcion(r_int_Posici).Caption <> "-" Then
         mnuRop_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   'MENU REPORTES SUNAT
   For r_int_Posici = 1 To mnuRsu_Opcion.Count
      If mnuRsu_Opcion(r_int_Posici).Caption <> "-" Then
         mnuRsu_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   'MENU CUENTAS POR PAGAR
   For r_int_Posici = 1 To mnuCxP_Opcion.Count
      If mnuCxP_Opcion(r_int_Posici).Caption <> "-" Then
         mnuCxP_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   'Verificando si todas las Opciones están habilitadas
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM SEG_PLTOPC "
   g_str_Parame = g_str_Parame & " WHERE PLTOPC_CODPLT = '" & UCase(App.EXEName) & "' "
   g_str_Parame = g_str_Parame & "   AND PLTOPC_FLGMEN = 2 "
   g_str_Parame = g_str_Parame & " ORDER BY PLTOPC_CODMEN ASC, PLTOPC_CODSUB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodMen = Trim(g_rst_Princi!PLTOPC_CODMEN)
         Do While Not g_rst_Princi.EOF And r_str_CodMen = Trim(g_rst_Princi!PLTOPC_CODMEN)
            Select Case r_str_CodMen
               Case "MNUMNT": mnuMnt_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUCON": mnuCon_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUPRO": mnuPro_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUASI": mnuAsi_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUMAT": mnuMat_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUEDP": mnuEdp_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUROP": mnuRop_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNURSU": mnuRsu_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUCXP": mnuCxP_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
            End Select
            
            g_rst_Princi.MoveNext
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Verificando por Plantilla de Acceso
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM SEG_PLTPLA "
   g_str_Parame = g_str_Parame & " WHERE PLTPLA_CODPLT = '" & UCase(App.EXEName) & "' "
   g_str_Parame = g_str_Parame & "   AND PLTPLA_TIPUSU = '" & CStr(modgen_g_int_TipUsu) & "' "
   g_str_Parame = g_str_Parame & " ORDER BY PLTPLA_CODMEN ASC, PLTPLA_CODSUB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodMen = Trim(g_rst_Princi!PLTPLA_CODMEN)
         Do While Not g_rst_Princi.EOF And r_str_CodMen = Trim(g_rst_Princi!PLTPLA_CODMEN)
            Select Case r_str_CodMen
               Case "MNUMNT": mnuMnt_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUCON": mnuCon_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUPRO": mnuPro_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUASI": mnuAsi_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUMAT": mnuMat_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUEDP": mnuEdp_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUROP": mnuRop_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNURSU": mnuRsu_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUCXP": mnuCxP_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
            End Select
            
            g_rst_Princi.MoveNext
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Verificando por Personalización de Opciones
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM SEG_PLTUSU "
   g_str_Parame = g_str_Parame & " WHERE PLTUSU_CODUSU = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "   AND PLTUSU_CODPLT = '" & UCase(App.EXEName) & "' "
   g_str_Parame = g_str_Parame & " ORDER BY PLTUSU_CODMEN ASC, PLTUSU_CODSUB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodMen = Trim(g_rst_Princi!PLTUSU_CODMEN)
         Do While Not g_rst_Princi.EOF And r_str_CodMen = Trim(g_rst_Princi!PLTUSU_CODMEN)
            Select Case r_str_CodMen
               Case "MNUMNT": mnuMnt_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUCON": mnuCon_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUPRO": mnuPro_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUASI": mnuAsi_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUMAT": mnuMat_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUEDP": mnuEdp_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUROP": mnuRop_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNURSU": mnuRsu_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUCXP": mnuCxP_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
            End Select
            
            g_rst_Princi.MoveNext
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

