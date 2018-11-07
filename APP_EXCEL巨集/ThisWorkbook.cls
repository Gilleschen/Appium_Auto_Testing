VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo skipp
    Application.CommandBars("功能說明").delete
    
    
skipp:
    ActiveWorkbook.Save
End Sub

Private Sub Workbook_Open()
    Dim myNewBar As CommandBar           '宣告工具列物件
    Dim ScriptCreator, ScriptEditor, CheckAPP, RunScript, CheckAPPandDevice, CheckCommand, CheckValue, CheckExpectResult As CommandBarButton '宣告工具列按鈕物件

    Set myNewBar = Application.CommandBars.add      '新增一個工具列
    On Error GoTo skipp

    myNewBar.Name = "功能說明" '工具列命名
    myNewBar.Position = msoBarTop
   

    With myNewBar
        Set RunScript = .Controls.add(msoControlButton)
        Set CheckAPPandDevice = .Controls.add(msoControlButton)
        Set CheckCommand = .Controls.add(msoControlButton)
        Set CheckValue = .Controls.add(msoControlButton)
        Set CheckExpectResult = .Controls.add(msoControlButton)
        Set ScriptCreator = .Controls.add(msoControlButton)
        Set ScriptEditor = .Controls.add(msoControlButton)
        Set CheckAPP = .Controls.add(msoControlButton)
        
        With ScriptCreator
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = "腳本產生器"
            .TooltipText = "建立新腳本"
            .FaceId = 592
            .Tag = "Jhih Chen"
            .OnAction = "LaunchScriptCreator"
        End With
        
        With ScriptEditor
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = "腳本編輯器"
            .TooltipText = "編輯現有腳本"
            .FaceId = 593
            .Tag = "Jhih Chen"
            .OnAction = "LaunchScriptEditor"
        End With
        
        With CheckAPP
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = "指令檢查"
            .TooltipText = "檢查特定腳本指令"
            .FaceId = 202
            .Tag = "Jhih Chen"
            .OnAction = "LaunchCheckAPP"
        End With
        
        
        With RunScript

            '.Style = msoButtonCaption   '只顯示文字 底下這3種型式選一種
            '.Style = msoButtonIcon       '只顯示小圖示
            '同時顯示文字和小圖示
            .Style = msoButtonIconAndCaption 'msoButtonIconAndCaption 'msoButtonIconAndWrapCaption 'msoButtonIconAndWrapCaptionBelow 'msoButtonIconAndCaptionBelow 'msoButtonIconAndCaption ' msoCombolabel
            .BeginGroup = True
            .Caption = "執行腳本"                  '顯示在工具列上的按鈕文字
            .TooltipText = "執行腳本" ' "TUTK QA團隊製作"  '滑鼠移過去時，所顯示的提示文字
            .FaceId = 156 '1763 '1778                       '小圖示
            .Tag = "Jhih Chen"
            .OnAction = "RunScript"     '設定按下此鍵時所要執行的巨集
        End With
        
        With CheckAPPandDevice
            .Style = msoButtonIconAndCaption  ' msoCombolabel
            .BeginGroup = True
            .Caption = "檢查資訊"                 '
            .TooltipText = "檢查APP&Device工作表資訊"
            .FaceId = 720
            .Tag = "Jhih Chen"
            .OnAction = "CheckAPPandDevice"
        
        End With
        
        With CheckCommand
            .Style = msoButtonIconAndCaption  ' msoCombolabel
            .BeginGroup = True
            .Caption = "檢查案例語法"                 '
            .TooltipText = "檢查每Case結束後是否執行QiutAPP"
            .FaceId = 720
            .Tag = "Jhih Chen"
            .OnAction = "CheckCommand"
        End With
            
        With CheckValue
            .Style = msoButtonIconAndCaption  ' msoCombolabel
            .BeginGroup = True
            .Caption = "檢查案例輸入值"                 '
            .TooltipText = "檢查所有命令及參數是否正確"
            .FaceId = 720
            .Tag = "Jhih Chen"
            .OnAction = "CheckValue"
        End With
        
        With CheckExpectResult
            .Style = msoButtonIconAndCaption  ' msoCombolabel
            .BeginGroup = True
            .Caption = "檢查期望結果"                 '
            .TooltipText = "確認Case之期望結果均列於ExpectResult工作表(非所有Case都要填寫ExpectResult)"
            .FaceId = 720 '25
            .Tag = "Jhih Chen"
            .OnAction = "CheckExpectResult_Ver2"
        
        End With
        
        
        
        'FaceId :http://juve.tistory.com/archive/20131205

        .Position = msoBarTop             '工具列擺放在上層
        .Visible = True
    End With
    
    Call loadingDeviceInfo
    
    
skipp:

End Sub


