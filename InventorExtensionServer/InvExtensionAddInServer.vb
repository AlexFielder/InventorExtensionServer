Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Linq
Imports Microsoft.Win32

Namespace InventorExtensionServer
    <ProgIdAttribute("InventorExtensionServer.StandardAddInServer"), _
    GuidAttribute("3bb5811b-0cf1-4253-8971-08e392acc02d")> _
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        Private WithEvents m_uiEvents As UserInterfaceEvents
        Private WithEvents m_UserInputEvents As UserInputEvents
        Private WithEvents m_sampleButton As ButtonDefinition

#Region "ApplicationAddInServer Members"

        ' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  
        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            g_inventorApplication = addInSiteObject.Application

            ' Connect to the user-interface events to handle a ribbon reset.
            m_uiEvents = g_inventorApplication.UserInterfaceManager.UserInterfaceEvents
            m_UserInputEvents = g_inventorApplication.CommandManager.UserInputEvents

            AddHandler m_UserInputEvents.OnLinearMarkingMenu, AddressOf OnLinearMarkingMenu

            ' TODO: Add button definitions.

            ' Sample to illustrate creating a button definition.
            Dim largeIconImg As Drawing.Icon = New Drawing.Icon(My.Resources.IconLarge, 48, 48)
            Dim largeIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(largeIconImg)
            Dim smallIconImg As System.Drawing.Icon = New Drawing.Icon(My.Resources.IconSmall, 16, 16)
            Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(smallIconImg)
            Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions
            m_sampleButton = controlDefs.AddButtonDefinition("Command Name",
                                                             My.Settings.ButtonInternalNameTrollfaceProblem,
                                                             CommandTypesEnum.kShapeEditCmdType,
                                                             AddInClientID,
                                                             "This is a basic Hello World command!",
                                                             "This is a basic Hello World command!",
                                                             smallIcon,
                                                             largeIcon, ButtonDisplayEnum.kAlwaysDisplayText)

            m_sampleButton.ProgressiveToolTip.Description = "This is a basic Hello World command!"
            m_sampleButton.ProgressiveToolTip.ExpandedDescription = "This is the progressive tooltip for our new command. Change this description as necessary!"
            m_sampleButton.ProgressiveToolTip.IsProgressive = True
            m_sampleButton.ProgressiveToolTip.Image = PictureDispConverter.ToIPictureDisp(My.Resources.RibbonProgressiveToolTipImage)
            m_sampleButton.ProgressiveToolTip.Title = "This is fine."


            ' Add to the user interface, if it's the first time.
            If firstTime Then
                AddToUserInterface()
            End If
        End Sub

        ''' <summary>
        ''' This method adds the button to the relevant menu in Part, Assembly, Drawing And Presentation environments. For reasons that remain unclear however,
        ''' right-clicking the top level icon in a drawing or Presentation fires a different non-UserInputEvents-capturable menu. The command in this case gets added to the 
        ''' right-click menu of the first node underneath the top one.
        ''' </summary>
        ''' <param name="SelectedEntities"></param>
        ''' <param name="SelectionDevice"></param>
        ''' <param name="LinearMenu"></param>
        ''' <param name="AdditionalInfo"></param>
        Private Sub OnLinearMarkingMenu(SelectedEntities As ObjectsEnumerator, SelectionDevice As SelectionDeviceEnum, LinearMenu As CommandControls, AdditionalInfo As NameValueMap)
            'Iterate each controls in linear menu
            Dim buttonadded As Boolean = False
            Dim ctrlToDelete As CommandControl = (From ctrl As CommandControl In LinearMenu
                                                  Where ctrl.InternalName = My.Settings.ButtonInternalNameTrollfaceProblem
                                                  Select ctrl).FirstOrDefault()
            If Not ctrlToDelete Is Nothing Then
                ctrlToDelete.Delete()
            End If

            For Each ctrl As CommandControl In LinearMenu
                If (ctrl.InternalName = "AppDeleteCmd") Then
                    If SketchBlockOrSketchBlocksSelected(SelectedEntities) Then
                        If Not buttonadded Then
                            LinearMenu.AddButton(m_sampleButton,
                                             False,
                                             True,
                                             ctrl.InternalName,
                                             True)
                        End If
                        LinearMenu.AddSeparator(ctrl.InternalName, True)
                        buttonadded = True
                    End If

                End If

                'Inserts new button and separator in the linear menu
                ' going to add our control after the "MOVE EOP COMMAND"
                If (ctrl.InternalName = "PartMoveEOPMarkerCmd") Then
                    If Not buttonadded Then
                        LinearMenu.AddButton(m_sampleButton,
                                             False,
                                             True,
                                             ctrl.InternalName,
                                             True)

                        LinearMenu.AddSeparator(ctrl.InternalName, True)
                        buttonadded = True
                    End If
                End If
                If (ctrl.InternalName = "AppHowToCmd") Then
                    If Not buttonadded Then
                        LinearMenu.AddButton(m_sampleButton,
                                                False,
                                                True,
                                                ctrl.InternalName,
                                                True)
                        buttonadded = True
                    End If
                End If

                If ctrl.InternalName = "UCxCreateDrawingViewCmd" Then
                    If Not buttonadded Then
                        LinearMenu.AddSeparator(ctrl.InternalName, False)
                        LinearMenu.AddButton(m_sampleButton,
                                             False,
                                             True,
                                             ctrl.InternalName,
                                             False)
                        buttonadded = True
                    End If
                End If
            Next

        End Sub

        ''' <summary>
        ''' Determines on the fly whether we have a BrowserFolder selected.
        ''' </summary>
        ''' <param name="selectedEntities"></param>
        ''' <returns></returns>
        Private Function IsPartBrowserFolder(selectedEntities As ObjectsEnumerator) As Boolean
            For Each entity As Object In selectedEntities
                'Dim folder As BrowserFolder = Nothing
                Dim folder As BrowserFolder = TryCast(entity, BrowserFolder)
                'Dim folder As BrowserFolder = entity
                If Not folder Is Nothing Then
                    'If Not TypeOf entity Is BrowserFolder Then
                    Return True
                    Exit Function
                Else
                    Return False
                    Exit Function
                End If
            Next
        End Function

        ''' <summary>
        ''' Determines on the fly whether we have a sketch block or sketch blocks collection selected in the browser.
        ''' </summary>
        ''' <param name="selectedEntities"></param>
        ''' <returns></returns>
        Private Function SketchBlockOrSketchBlocksSelected(selectedEntities As ObjectsEnumerator) As Boolean
            For Each entity As Object In selectedEntities
                Dim sketchblockentity As SketchBlockDefinition = TryCast(entity, SketchBlockDefinition)
                'Dim folder As BrowserFolder = entity
                If Not sketchblockentity Is Nothing Then
                    'If Not TypeOf entity Is BrowserFolder Then
                    Return True
                    Exit Function
                Else
                    Return False
                    Exit Function
                End If
            Next
        End Function

        ' This method is called by Inventor when the AddIn is unloaded. The AddIn will be
        ' unloaded either manually by the user or when the Inventor session is terminated.
        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            ' TODO:  Add ApplicationAddInServer.Deactivate implementation

            ' Release objects.
            m_uiEvents = Nothing
            g_inventorApplication = Nothing

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
        End Sub

        ' This property is provided to allow the AddIn to expose an API of its own to other 
        ' programs. Typically, this  would be done by implementing the AddIn's API
        ' interface in a class and returning that class object through this property.
        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        ' Note:this method is now obsolete, you should use the 
        ' ControlDefinition functionality for implementing commands.
        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
        End Sub

#End Region

#Region "User interface definition"
        ' Sub where the user-interface creation is done.  This is called when
        ' the add-in loaded and also if the user interface is reset.
        Private Sub AddToUserInterface()
            ' This is where you'll add code to add buttons to the ribbon.

            '** Sample to illustrate creating a button on a new panel of the Tools tab of the Part ribbon.

            '' Get the part ribbon.
            Dim partRibbon As Ribbon = g_inventorApplication.UserInterfaceManager.Ribbons.Item("Part") 'possible options here are: ZeroDoc, Part, Assembly, Drawing, Presentation

            '' Get the "Tools" tab.
            Dim toolsTab As RibbonTab = partRibbon.RibbonTabs.Item("id_TabTools") ' we know this is the ribbontab we're after. To generate a list of all 'internal' names refer to this: https://github.com/AlexFielder/iLogic/blob/master/PrintRibbonNames.bas (Import into the default.ivb inside of Inventor vba and run PrintRibbon() Command)

            '' Create a new panel.
            Dim customPanel As RibbonPanel = toolsTab.RibbonPanels.Add("Sample", "MysSample", AddInClientID)

            '' Add a button.
            customPanel.CommandControls.AddButton(m_sampleButton, True, True)
        End Sub

        Private Sub m_uiEvents_OnResetRibbonInterface(Context As NameValueMap) Handles m_uiEvents.OnResetRibbonInterface
            ' The ribbon was reset, so add back the add-ins user-interface.
            AddToUserInterface()
        End Sub

        ' Sample handler for the button.
        Private Sub m_sampleButton_OnExecute(Context As NameValueMap) Handles m_sampleButton.OnExecute
            MsgBox("Hello World!")
        End Sub
#End Region

    End Class
End Namespace


Public Module Globals
    ' Inventor application object.
    Public g_inventorApplication As Inventor.Application

#Region "Function to get the add-in client ID."
    ' This function uses reflection to get the GuidAttribute associated with the add-in.
    Public Function AddInClientID() As String
        Dim guid As String = ""
        Try
            Dim t As Type = GetType(InventorExtensionServer.StandardAddInServer)
            Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
            Dim guidAttribute As GuidAttribute = CType(customAttributes(0), GuidAttribute)
            guid = "{" + guidAttribute.Value.ToString() + "}"
        Catch
        End Try

        Return guid
    End Function
#End Region

#Region "hWnd Wrapper Class"
    ' This class is used to wrap a Win32 hWnd as a .Net IWind32Window class.
    ' This is primarily used for parenting a dialog to the Inventor window.
    '
    ' For example:
    ' myForm.Show(New WindowWrapper(g_inventorApplication.MainFrameHWND))
    '
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As IntPtr _
          Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

        Private _hwnd As IntPtr
    End Class
#End Region

#Region "Image Converter"
    ' Class used to convert bitmaps and icons from their .Net native types into
    ' an IPictureDisp object which is what the Inventor API requires. A typical
    ' usage is shown below where MyIcon is a bitmap or icon that's available
    ' as a resource of the project.
    '
    ' Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.MyIcon)

    Public NotInheritable Class PictureDispConverter
        <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)> _
        Private Shared Function OleCreatePictureIndirect( _
            <MarshalAs(UnmanagedType.AsAny)> ByVal picdesc As Object, _
            ByRef iid As Guid, _
            <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean) As stdole.IPictureDisp
        End Function

        Shared iPictureDispGuid As Guid = GetType(stdole.IPictureDisp).GUID

        Private NotInheritable Class PICTDESC
            Private Sub New()
            End Sub

            'Picture Types
            Public Const PICTYPE_BITMAP As Short = 1
            Public Const PICTYPE_ICON As Short = 3

            <StructLayout(LayoutKind.Sequential)> _
            Public Class Icon
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Icon))
                Friend picType As Integer = PICTDESC.PICTYPE_ICON
                Friend hicon As IntPtr = IntPtr.Zero
                Friend unused1 As Integer
                Friend unused2 As Integer

                Friend Sub New(ByVal icon As System.Drawing.Icon)
                    Me.hicon = icon.ToBitmap().GetHicon()
                End Sub
            End Class

            <StructLayout(LayoutKind.Sequential)> _
            Public Class Bitmap
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Bitmap))
                Friend picType As Integer = PICTDESC.PICTYPE_BITMAP
                Friend hbitmap As IntPtr = IntPtr.Zero
                Friend hpal As IntPtr = IntPtr.Zero
                Friend unused As Integer

                Friend Sub New(ByVal bitmap As System.Drawing.Bitmap)
                    Me.hbitmap = bitmap.GetHbitmap()
                End Sub
            End Class
        End Class

        Public Shared Function ToIPictureDisp(ByVal icon As System.Drawing.Icon) As stdole.IPictureDisp
            Dim pictIcon As New PICTDESC.Icon(icon)
            Return OleCreatePictureIndirect(pictIcon, iPictureDispGuid, True)
        End Function

        Public Shared Function ToIPictureDisp(ByVal bmp As System.Drawing.Bitmap) As stdole.IPictureDisp
            Dim pictBmp As New PICTDESC.Bitmap(bmp)
            Return OleCreatePictureIndirect(pictBmp, iPictureDispGuid, True)
        End Function
    End Class
#End Region

End Module
