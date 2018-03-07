Attribute VB_Name = "modLang"
'*********************************************
'modLang
'Copyright VisualBasicZone.com 2004 - 2005
'*********************************************
Option Explicit
'Language list
Global gDefaultLanguage As String
Global gLanguageList() As String

Private Type LangType
    Title As String
    strTreePEHEADER As String
    strTreeVBStrucutres As String
    strTreeVBHEADER As String
    strTreeVBProjectInformation As String
    strTreeVBComRegistrationData As String
    strTreeVBObjectTable As String
    strTreeVBObjects As String
    strTreeForms As String
    strTreeModules As String
    strTreeClasses As String
    strTreeUserControls As String
    strTreePropertyPages As String
    strTreeProceduresCode As String
    strTreeImages As String
    strTreeFileVersionInformation As String
    strTreeImportInformation As String
End Type

Global Lang As LangType

