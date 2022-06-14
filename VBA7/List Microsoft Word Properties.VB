On Error Resume Next

Set objWord = CreateObject("Word.Application")
Wscript.Echo "Active Printer:", objWord.ActivePrinter

For Each objAddIn in objWord.AddIns
    Wscript.Echo "AddIn: ", objAddIn
Next

Wscript.Echo "Application:", objWord.Application
Wscript.Echo "Assistant:", objWord.Assistant

For Each objCaption in objWord.AutoCaptions
    Wscript.Echo "AutoCaptions:", objCaption
Next
Wscript.Echo "Automation Security:", objWord.AutomationSecurity
Wscript.Echo "Background Printing Status:", objWord.BackgroundPrintingStatus
Wscript.Echo "Background Saving Status:", objWord.BackgroundSavingStatus
Wscript.Echo "Browse Extra File Type:", objWord.BrowseExtraFileTypes
Wscript.Echo "Build:", objWord.Build
Wscript.Echo "Caps Lock:", objWord.CapsLock
Wscript.Echo "Caption:", objWord.Caption

For Each objLabel in objWord.CaptionLabels
    Wscript.Echo "Caption Label:", objLabel
Next

Wscript.Echo "Check Language:", objWord.CheckLanguage

For Each objAddIn in objWord.COMAddIns
    Wscript.Echo "COM AddIn:", objAddIn
Next

Wscript.Echo "Creator:", objWord.Creator

For Each objDictionary in objWord.CustomDictionaries
    Wscript.Echo "Custom Dictionary:", objDictionary
Next

Wscript.Echo "Customization Context:", objWord.CustomizationContext
Wscript.Echo "Default Legal Blackline:", objWord.DefaultLegalBlackline
Wscript.Echo "Default Save Format:", objWord.DefaultSaveFormat
Wscript.Echo "Default Table Separator:", objWord.DefaultTableSeparator

For Each objDialog in objWord.Dialogs
    Wscript.Echo "Dialog:", objDialog
Next

Wscript.Echo "Display Alerts:", objWord.DisplayAlerts
Wscript.Echo "Display Recent Files:", objWord.DisplayRecentFiles
Wscript.Echo "Display Screen Tips:", objWord.DisplayScreenTips
Wscript.Echo "Display Scroll Bars:", objWord.DisplayScrollBars

For Each objDocument in objWord.Documents
    Wscript.Echo "Document:", objDocument
Next

Wscript.Echo "Email Template:", objWord.EmailTemplate
Wscript.Echo "Enable Cancel Key:", objWord.EnableCancelKey
Wscript.Echo "Feature Install:", objWord.FeatureInstall

For Each objConverter in objWord.FileConverters
    Wscript.Echo "File Converter:", objConverter
Next

Wscript.Echo "Focus In MailHeader:", objWord.FocusInMailHeader

For Each objFont in objWord.FontNames
    Wscript.Echo "Font Name:", objFont
Next

Wscript.Echo "Height", objWord.Height

For Each objBinding in objWord.KeyBindings
    Wscript.Echo "Key Binding:", objBinding
Next

For Each objFont in objWord.LandscapeFontNames
    Wscript.Echo "Landscape Font Name:", objFont
Next

Wscript.Echo "Language", objWord.Language

For Each objLanguage in objWord.Languages
    Wscript.Echo "Language:", objLanguage
Next

Wscript.Echo "Left", objWord.Left
Wscript.Echo "Mail System:", objWord.MailSystem
Wscript.Echo "MAPI Available:", objWord.MAPIAvailable
Wscript.Echo "Math Coprocessor Available:", objWord.MathCoprocessorAvailable
Wscript.Echo "Mouse Available:", objWord.MouseAvailable
Wscript.Echo "Name:", objWord.Name
Wscript.Echo "Normal Template:", objWord.NormalTemplate
Wscript.Echo "Num Lock:", objWord.NumLock
Wscript.Echo "Parent:", objWord.Parent
Wscript.Echo "Path:", objWord.Path
Wscript.Echo "Path Separator:", objWord.PathSeparator
Wscript.Echo "Print Preview:", objWord.PrintPreview

For Each objFile in objWord.RecentFiles
    Wscript.Echo "Recent File:", objFile
Next

Wscript.Echo "Screen Updating:", objWord.ScreenUpdating
Wscript.Echo "Show Visual Basic Editor:", objWord.ShowVisualBasicEditor
Wscript.Echo "Special Mode:", objWord.SpecialMode
Wscript.Echo "Startup Path:", objWord.StartupPath

For Each objTask in objWord.Tasks
    Wscript.Echo "Task:", objTask
Next

For Each objTemplate in objWord.Templates
    Wscript.Echo "Template:", objTemplate
Next

Wscript.Echo "Top:", objWord.Top
Wscript.Echo "Usable Height:", objWord.UsableHeight
Wscript.Echo "Usable Width:", objWord.UsableWidth
Wscript.Echo "User Address:", objWord.UserAddress
Wscript.Echo "User Control:", objWord.UserControl
Wscript.Echo "User Initials:", objWord.UserInitials
Wscript.Echo "User Name:", objWord.UserName
Wscript.Echo "Version:", objWord.Version
Wscript.Echo "Visible:", objWord.Visible
Wscript.Echo "Width:", objWord.Width

For Each objWindow in objWord.Windows
    Wscript.Echo "Window:", objWindow
Next

Wscript.Echo "Window State:", objWord.WindowState
objWord.Quit