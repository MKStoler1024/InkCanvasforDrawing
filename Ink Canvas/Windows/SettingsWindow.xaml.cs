using Ink_Canvas.Helpers;
using iNKORE.UI.WPF.Modern;
using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Ink_Canvas
{
    public partial class SettingsWindow : Window
    {
        private MainWindow mainWindow;
        private bool isLoaded = false;

        public SettingsWindow(MainWindow owner)
        {
            InitializeComponent();
            mainWindow = owner;
            Owner = owner;
            
            ApplyTheme();
            LoadSettings();
            isLoaded = true;
        }

        private void ApplyTheme()
        {
            if (mainWindow != null)
            {
                if (mainWindow.GetMainWindowTheme() == "Light")
                {
                    ThemeManager.SetRequestedTheme(this, ElementTheme.Light);
                }
                else
                {
                    ThemeManager.SetRequestedTheme(this, ElementTheme.Dark);
                }
            }
        }

        private void LoadSettings()
        {
            // 启动选项
            ToggleSwitchIsAutoUpdate.IsOn = MainWindow.Settings.Startup.IsAutoUpdate;
            ToggleSwitchIsAutoUpdateWithSilence.IsOn = MainWindow.Settings.Startup.IsAutoUpdateWithSilence;
            ToggleSwitchFoldAtStartup.IsOn = MainWindow.Settings.Startup.IsFoldAtStartup;

            // 初始化自动更新时间段
            InitializeTimeComboBoxes();
            if (MainWindow.Settings.Startup.AutoUpdateWithSilenceStartTime != null)
                AutoUpdateWithSilenceStartTimeComboBox.SelectedItem = MainWindow.Settings.Startup.AutoUpdateWithSilenceStartTime;
            if (MainWindow.Settings.Startup.AutoUpdateWithSilenceEndTime != null)
                AutoUpdateWithSilenceEndTimeComboBox.SelectedItem = MainWindow.Settings.Startup.AutoUpdateWithSilenceEndTime;

            // 画板选项
            ToggleSwitchCompressPicturesUploaded.IsOn = MainWindow.Settings.Canvas.IsCompressPicturesUploaded;
            ToggleSwitchShowCursor.IsOn = MainWindow.Settings.Canvas.IsShowCursor;
            ComboBoxEraserSize.SelectedIndex = MainWindow.Settings.Canvas.EraserSize;
            ToggleSwitchHideStrokeWhenSelecting.IsOn = MainWindow.Settings.Canvas.HideStrokeWhenSelecting;
            ComboBoxHyperbolaAsymptoteOption.SelectedIndex = (int)MainWindow.Settings.Canvas.HyperbolaAsymptoteOption;

            // 手势选项
            ComboBoxMatrixTransformCenterPoint.SelectedIndex = (int)MainWindow.Settings.Gesture.MatrixTransformCenterPoint;
            ToggleSwitchAutoSwitchTwoFingerGesture.IsOn = MainWindow.Settings.Gesture.AutoSwitchTwoFingerGesture;
            ToggleSwitchEnableTwoFingerRotationOnSelection.IsOn = MainWindow.Settings.Gesture.IsEnableTwoFingerRotationOnSelection;

            // 墨迹识别
            ToggleSwitchEnableInkToShape.IsOn = MainWindow.Settings.InkToShape.IsInkToShapeEnabled;

            // 外观选项
            ComboBoxTheme.SelectedIndex = MainWindow.Settings.Appearance.Theme;
            ToggleSwitchEnableDisPlayFloatBarText.IsOn = MainWindow.Settings.Appearance.IsEnableDisPlayFloatBarText;
            ToggleSwitchEnableDisPlayNibModeToggle.IsOn = MainWindow.Settings.Appearance.IsEnableDisPlayNibModeToggler;
            ToggleSwitchColorfulViewboxFloatingBar.IsOn = MainWindow.Settings.Appearance.IsColorfulViewboxFloatingBar;
            SliderFloatingBarBottomMargin.Value = MainWindow.Settings.Appearance.FloatingBarBottomMargin;
            SliderFloatingBarScale.Value = MainWindow.Settings.Appearance.FloatingBarScale;
            SliderBlackboardScale.Value = MainWindow.Settings.Appearance.BlackboardScale;

            // PPT 选项
            ToggleSwitchShowPPTNavigationPanelBottom.IsOn = MainWindow.Settings.PowerPointSettings.IsShowBottomPPTNavigationPanel;
            ToggleSwitchShowButtonPPTNavigationBottom.IsOn = MainWindow.Settings.PowerPointSettings.IsShowPPTNavigationBottom;
            ToggleSwitchShowPPTNavigationPanelSide.IsOn = MainWindow.Settings.PowerPointSettings.IsShowSidePPTNavigationPanel;
            ToggleSwitchShowButtonPPTNavigationSides.IsOn = MainWindow.Settings.PowerPointSettings.IsShowPPTNavigationSides;
            ToggleSwitchSupportPowerPoint.IsOn = MainWindow.Settings.PowerPointSettings.PowerPointSupport;
            ToggleSwitchSupportWPS.IsOn = MainWindow.Settings.PowerPointSettings.IsSupportWPS;
            ToggleSwitchShowCanvasAtNewSlideShow.IsOn = MainWindow.Settings.PowerPointSettings.IsShowCanvasAtNewSlideShow;
            ToggleSwitchEnableTwoFingerGestureInPresentationMode.IsOn = MainWindow.Settings.PowerPointSettings.IsEnableTwoFingerGestureInPresentationMode;
            ToggleSwitchEnableFingerGestureSlideShowControl.IsOn = MainWindow.Settings.PowerPointSettings.IsEnableFingerGestureSlideShowControl;
            ToggleSwitchAutoSaveScreenShotInPowerPoint.IsOn = MainWindow.Settings.PowerPointSettings.IsAutoSaveScreenShotInPowerPoint;
            ToggleSwitchAutoSaveStrokesInPowerPoint.IsOn = MainWindow.Settings.PowerPointSettings.IsAutoSaveStrokesInPowerPoint;
            ToggleSwitchNotifyPreviousPage.IsOn = MainWindow.Settings.PowerPointSettings.IsNotifyPreviousPage;
            ToggleSwitchNotifyHiddenPage.IsOn = MainWindow.Settings.PowerPointSettings.IsNotifyHiddenPage;
            ToggleSwitchNotifyAutoPlayPresentation.IsOn = MainWindow.Settings.PowerPointSettings.IsNotifyAutoPlayPresentation;

            // 高级选项
            ToggleSwitchIsSpecialScreen.IsOn = MainWindow.Settings.Advanced.IsSpecialScreen;
            ToggleSwitchIsLogEnabled.IsOn = MainWindow.Settings.Advanced.IsLogEnabled;
            ToggleSwitchIsSecondConfimeWhenShutdownApp.IsOn = MainWindow.Settings.Advanced.IsSecondConfimeWhenShutdownApp;
            ToggleSwitchIsQuadIR.IsOn = MainWindow.Settings.Advanced.IsQuadIR;
            TouchMultiplierSlider.Value = MainWindow.Settings.Advanced.TouchMultiplier;
            NibModeBoundsWidthSlider.Value = MainWindow.Settings.Advanced.NibModeBoundsWidth;
            FingerModeBoundsWidthSlider.Value = MainWindow.Settings.Advanced.FingerModeBoundsWidth;
            NibModeBoundsWidthThresholdValueSlider.Value = MainWindow.Settings.Advanced.NibModeBoundsWidthThresholdValue;
            FingerModeBoundsWidthThresholdValueSlider.Value = MainWindow.Settings.Advanced.FingerModeBoundsWidthThresholdValue;
            NibModeBoundsWidthEraserSizeSlider.Value = MainWindow.Settings.Advanced.NibModeBoundsWidthEraserSize;
            FingerModeBoundsWidthEraserSizeSlider.Value = MainWindow.Settings.Advanced.FingerModeBoundsWidthEraserSize;

            // 自动选项
            ToggleSwitchAutoFoldInEasiNote.IsOn = MainWindow.Settings.Automation.IsAutoFoldInEasiNote;
            ToggleSwitchAutoFoldInEasiNoteIgnoreDesktopAnno.IsOn = MainWindow.Settings.Automation.IsAutoFoldInEasiNoteIgnoreDesktopAnno;
            ToggleSwitchAutoFoldInEasiCamera.IsOn = MainWindow.Settings.Automation.IsAutoFoldInEasiCamera;
            ToggleSwitchAutoFoldInEasiNote3C.IsOn = MainWindow.Settings.Automation.IsAutoFoldInEasiNote3C;
            ToggleSwitchAutoFoldInSeewoPincoTeacher.IsOn = MainWindow.Settings.Automation.IsAutoFoldInSeewoPincoTeacher;
            ToggleSwitchAutoFoldInHiteTouchPro.IsOn = MainWindow.Settings.Automation.IsAutoFoldInHiteTouchPro;
            ToggleSwitchAutoFoldInHiteCamera.IsOn = MainWindow.Settings.Automation.IsAutoFoldInHiteCamera;
            ToggleSwitchAutoFoldInWxBoardMain.IsOn = MainWindow.Settings.Automation.IsAutoFoldInWxBoardMain;
            ToggleSwitchAutoFoldInOldZyBoard.IsOn = MainWindow.Settings.Automation.IsAutoFoldInOldZyBoard;
            ToggleSwitchAutoFoldInMSWhiteboard.IsOn = MainWindow.Settings.Automation.IsAutoFoldInMSWhiteboard;
            ToggleSwitchAutoFoldInPPTSlideShow.IsOn = MainWindow.Settings.Automation.IsAutoFoldInPPTSlideShow;
            ToggleSwitchAutoKillPptService.IsOn = MainWindow.Settings.Automation.IsAutoKillPptService;
            ToggleSwitchAutoKillEasiNote.IsOn = MainWindow.Settings.Automation.IsAutoKillEasiNote;
            ToggleSwitchAutoSaveStrokesAtClear.IsOn = MainWindow.Settings.Automation.IsAutoSaveStrokesAtClear;
            ToggleSwitchSaveScreenshotsInDateFolders.IsOn = MainWindow.Settings.Automation.IsSaveScreenshotsInDateFolders;
            ToggleSwitchAutoSaveStrokesAtScreenshot.IsOn = MainWindow.Settings.Automation.IsAutoSaveStrokesAtScreenshot;
            SideControlMinimumAutomationSlider.Value = MainWindow.Settings.Automation.MinimumAutomationStrokeNumber;
            AutoSavedStrokesLocation.Text = MainWindow.Settings.Automation.AutoSavedStrokesLocation;
            ToggleSwitchAutoDelSavedFiles.IsOn = MainWindow.Settings.Automation.AutoDelSavedFiles;
            ComboBoxAutoDelSavedFilesDaysThreshold.SelectedIndex = GetComboBoxIndexByDaysThreshold(MainWindow.Settings.Automation.AutoDelSavedFilesDaysThreshold);
        }

        private void InitializeTimeComboBoxes()
        {
            for (int i = 0; i < 24; i++)
            {
                AutoUpdateWithSilenceStartTimeComboBox.Items.Add(string.Format("{0:D2}:00", i));
                AutoUpdateWithSilenceEndTimeComboBox.Items.Add(string.Format("{0:D2}:00", i));
            }
        }

        private int GetComboBoxIndexByDaysThreshold(int days)
        {
            if (days == 1) return 0;
            if (days == 3) return 1;
            if (days == 5) return 2;
            if (days == 7) return 3;
            if (days == 15) return 4;
            if (days == 30) return 5;
            if (days == 60) return 6;
            if (days == 100) return 7;
            if (days == 365) return 8;
            return 4;
        }

        #region 所有事件处理

        // Startup Events
        private void ToggleSwitchIsAutoUpdate_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Startup.IsAutoUpdate = ToggleSwitchIsAutoUpdate.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchIsAutoUpdateWithSilence_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Startup.IsAutoUpdateWithSilence = ToggleSwitchIsAutoUpdateWithSilence.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void AutoUpdateWithSilenceStartTimeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            if (AutoUpdateWithSilenceStartTimeComboBox.SelectedItem != null)
                MainWindow.Settings.Startup.AutoUpdateWithSilenceStartTime = (string)AutoUpdateWithSilenceStartTimeComboBox.SelectedItem;
            MainWindow.SaveSettingsToFile();
        }

        private void AutoUpdateWithSilenceEndTimeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            if (AutoUpdateWithSilenceEndTimeComboBox.SelectedItem != null)
                MainWindow.Settings.Startup.AutoUpdateWithSilenceEndTime = (string)AutoUpdateWithSilenceEndTimeComboBox.SelectedItem;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchRunAtStartup_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            if (ToggleSwitchRunAtStartup.IsOn)
            {
                MainWindow.StartAutomaticallyDel("InkCanvas");
                MainWindow.StartAutomaticallyDel("Ink Canvas Annotation");
                MainWindow.StartAutomaticallyCreate("InkCanvasforDrawing");
            }
            else
            {
                MainWindow.StartAutomaticallyDel("InkCanvas");
                MainWindow.StartAutomaticallyDel("Ink Canvas Annotation");
                MainWindow.StartAutomaticallyDel("InkCanvasforDrawing");
            }
        }

        private void ToggleSwitchFoldAtStartup_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Startup.IsFoldAtStartup = ToggleSwitchFoldAtStartup.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        // Canvas Events
        private void ComboBoxEraserSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Canvas.EraserSize = ComboBoxEraserSize.SelectedIndex;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchCompressPicturesUploaded_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Canvas.IsCompressPicturesUploaded = ToggleSwitchCompressPicturesUploaded.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchShowCursor_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Canvas.IsShowCursor = ToggleSwitchShowCursor.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchHideStrokeWhenSelecting_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Canvas.HideStrokeWhenSelecting = ToggleSwitchHideStrokeWhenSelecting.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ComboBoxHyperbolaAsymptoteOption_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Canvas.HyperbolaAsymptoteOption = (OptionalOperation)ComboBoxHyperbolaAsymptoteOption.SelectedIndex;
            MainWindow.SaveSettingsToFile();
        }

        // Gesture Events
        private void ComboBoxMatrixTransformCenterPoint_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Gesture.MatrixTransformCenterPoint = (MatrixTransformCenterPointOptions)ComboBoxMatrixTransformCenterPoint.SelectedIndex;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoSwitchTwoFingerGesture_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Gesture.AutoSwitchTwoFingerGesture = ToggleSwitchAutoSwitchTwoFingerGesture.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchEnableTwoFingerRotation_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Gesture.IsEnableTwoFingerRotationOnSelection = ToggleSwitchEnableTwoFingerRotationOnSelection.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        // InkToShape Events
        private void ToggleSwitchEnableInkToShape_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.InkToShape.IsInkToShapeEnabled = ToggleSwitchEnableInkToShape.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        // Appearance Events
        private void ComboBoxTheme_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.Theme = ComboBoxTheme.SelectedIndex;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchEnableDisPlayFloatBarText_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.IsEnableDisPlayFloatBarText = ToggleSwitchEnableDisPlayFloatBarText.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchEnableDisPlayNibModeToggle_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.IsEnableDisPlayNibModeToggler = ToggleSwitchEnableDisPlayNibModeToggle.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchIsColorfulViewboxFloatingBar_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.IsColorfulViewboxFloatingBar = ToggleSwitchColorfulViewboxFloatingBar.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void SliderFloatingBarBottomMargin_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.FloatingBarBottomMargin = e.NewValue;
            MainWindow.SaveSettingsToFile();
        }

        private void SliderFloatingBarScale_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.FloatingBarScale = e.NewValue;
            MainWindow.SaveSettingsToFile();
        }

        private void SliderBlackboardScale_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.BlackboardScale = e.NewValue;
            MainWindow.SaveSettingsToFile();
        }

        private void BtnSetFloatingBarMargin_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag != null && double.TryParse(btn.Tag.ToString(), out double margin))
            {
                SliderFloatingBarBottomMargin.Value = margin;
            }
        }

        private void BtnSetFloatingBarScale_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag != null && double.TryParse(btn.Tag.ToString(), out double scalePercent))
            {
                SliderFloatingBarScale.Value = scalePercent;
            }
        }

        private void BtnSetBlackboardScale_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag != null && double.TryParse(btn.Tag.ToString(), out double scalePercent))
            {
                SliderBlackboardScale.Value = scalePercent;
            }
        }

        // PPT Events
        private void ToggleSwitchShowPPTNavigationPanelBottom_OnToggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsShowBottomPPTNavigationPanel = ToggleSwitchShowPPTNavigationPanelBottom.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchShowButtonPPTNavigationBottom_OnToggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsShowPPTNavigationBottom = ToggleSwitchShowButtonPPTNavigationBottom.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchShowPPTNavigationPanelSide_OnToggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsShowSidePPTNavigationPanel = ToggleSwitchShowPPTNavigationPanelSide.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchShowButtonPPTNavigationSides_OnToggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsShowPPTNavigationSides = ToggleSwitchShowButtonPPTNavigationSides.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchSupportPowerPoint_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.PowerPointSupport = ToggleSwitchSupportPowerPoint.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchSupportWPS_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsSupportWPS = ToggleSwitchSupportWPS.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchShowCanvasAtNewSlideShow_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsShowCanvasAtNewSlideShow = ToggleSwitchShowCanvasAtNewSlideShow.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchEnableTwoFingerGestureInPresentationMode_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsEnableTwoFingerGestureInPresentationMode = ToggleSwitchEnableTwoFingerGestureInPresentationMode.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchEnableFingerGestureSlideShowControl_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsEnableFingerGestureSlideShowControl = ToggleSwitchEnableFingerGestureSlideShowControl.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoSaveScreenShotInPowerPoint_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsAutoSaveScreenShotInPowerPoint = ToggleSwitchAutoSaveScreenShotInPowerPoint.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoSaveStrokesInPowerPoint_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsAutoSaveStrokesInPowerPoint = ToggleSwitchAutoSaveStrokesInPowerPoint.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchNotifyPreviousPage_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsNotifyPreviousPage = ToggleSwitchNotifyPreviousPage.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchNotifyHiddenPage_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsNotifyHiddenPage = ToggleSwitchNotifyHiddenPage.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchNotifyAutoPlayPresentation_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsNotifyAutoPlayPresentation = ToggleSwitchNotifyAutoPlayPresentation.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        // Advanced Events
        private void ToggleSwitchIsSpecialScreen_OnToggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.IsSpecialScreen = ToggleSwitchIsSpecialScreen.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void TouchMultiplierSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.TouchMultiplier = e.NewValue;
            MainWindow.SaveSettingsToFile();
        }

        private void NibModeBoundsWidthSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.NibModeBoundsWidth = (int)e.NewValue;
            MainWindow.SaveSettingsToFile();
        }

        private void FingerModeBoundsWidthSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.FingerModeBoundsWidth = (int)e.NewValue;
            MainWindow.SaveSettingsToFile();
        }

        private void NibModeBoundsWidthThresholdValueSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.NibModeBoundsWidthThresholdValue = e.NewValue;
            MainWindow.SaveSettingsToFile();
        }

        private void FingerModeBoundsWidthThresholdValueSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.FingerModeBoundsWidthThresholdValue = e.NewValue;
            MainWindow.SaveSettingsToFile();
        }

        private void NibModeBoundsWidthEraserSizeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.NibModeBoundsWidthEraserSize = e.NewValue;
            MainWindow.SaveSettingsToFile();
        }

        private void FingerModeBoundsWidthEraserSizeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.FingerModeBoundsWidthEraserSize = e.NewValue;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchIsQuadIR_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.IsQuadIR = ToggleSwitchIsQuadIR.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchIsLogEnabled_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.IsLogEnabled = ToggleSwitchIsLogEnabled.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchIsSecondConfimeWhenShutdownApp_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.IsSecondConfimeWhenShutdownApp = ToggleSwitchIsSecondConfimeWhenShutdownApp.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        #region 高级选项 - 触摸倍数计算

        private void BorderCalculateMultiplier_TouchDown(object sender, TouchEventArgs e)
        {
            var args = e.GetTouchPoint(null).Bounds;
            double value;
            if (!MainWindow.Settings.Advanced.IsQuadIR) 
                value = args.Width;
            else 
                value = Math.Sqrt(args.Width * args.Height); // 四边红外
            TextBlockShowCalculatedMultiplier.Text = (5 / (value * 1.1)).ToString();
        }

        private void ToggleSwitchIsEnableEdgeGestureUtil_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.IsEnableEdgeGestureUtil = ToggleSwitchIsEnableEdgeGestureUtil.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        #endregion

        #region 自动选项 - 保存位置

        private void AutoSavedStrokesLocationButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog folderBrowser = new System.Windows.Forms.FolderBrowserDialog();
            if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                AutoSavedStrokesLocation.Text = folderBrowser.SelectedPath;
            }
        }

        private void SetAutoSavedStrokesLocationToDiskDButton_Click(object sender, RoutedEventArgs e)
        {
            AutoSavedStrokesLocation.Text = @"D:\Ink Canvas";
        }

        private void SetAutoSavedStrokesLocationToDocumentFolderButton_Click(object sender, RoutedEventArgs e)
        {
            AutoSavedStrokesLocation.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Ink Canvas";
        }

        #endregion

        #region 链接和窗口控制

        private void HyperlinkSourceToPresentRepository_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("https://github.com/InkCanvas/Ink-Canvas-Artistry");
        }

        private void HyperlinkSourceToOringinalRepository_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("https://github.com/WXRIW/Ink-Canvas");
        }

        #endregion

        private void SCManipulationBoundaryFeedback(object sender, ManipulationBoundaryFeedbackEventArgs e)
        {
            e.Handled = true;
        }
    }
}
