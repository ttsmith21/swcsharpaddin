Attribute VB_Name = "modConfig"
Option Explicit

' ----- Logging Configuration -----
Public Const LOG_ENABLED As Boolean = True  ' Enable or disable logging
Public Const ERROR_LOG_PATH As String = "C:\SolidWorksMacroLogs\ErrorLog.txt" ' Log file location
Public Const SHOW_WARNINGS As Boolean = False ' Set to True to show pop-ups for warnings

' ----- Excel File Paths -----
Public Const MATERIAL_FILE_PATH As String = "O:\Engineering Department\Solidworks\Macros\(Semi)Autopilot\Material-2022v4.xlsx"
Public Const LASER_DATA_FILE_PATH As String = "O:\Engineering Department\Solidworks\Macros\(Semi)Autopilot\NewLaser.xls"

' ----- Application Settings -----
Public Const MAX_RETRIES As Integer = 3  ' Number of retries for Excel file reads
Public Const ENABLE_DEBUG_MODE As Boolean = True ' Set to True to enable verbose debugging

' ----- User Preferences -----
Public Const DEFAULT_SHEET_NAME As String = "Sheet1"
Public Const AUTO_CLOSE_EXCEL As Boolean = True ' Set to False if Excel should remain open


