using Ink_Canvas.Helpers;
using iNKORE.UI.WPF.Modern;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;

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
                    ResourceDictionary rd = new ResourceDictionary() { Source = new Uri("Resources/Styles/Light-PopupWindow.xaml", UriKind.Relative) };
                    Application.Current.Resources.MergedDictionaries.Add(rd);
                }
                else
                {
                    ThemeManager.SetRequestedTheme(this, ElementTheme.Dark);
                    ResourceDictionary rd = new ResourceDictionary() { Source = new Uri("Resources/Styles/Dark-PopupWindow.xaml", UriKind.Relative) };
                    Application.Current.Resources.MergedDictionaries.Add(rd);
                }
            }
        }

        private void LoadSettings()
        {
            // 启动选项
            ToggleSwitchIsAutoUpdate.IsOn = MainWindow.Settings.Startup.IsAutoUpdate;
            ToggleSwitchIsAutoUpdateWithSilence.IsOn = MainWindow.Settings.Startup.IsAutoUpdateWithSilence;
            ToggleSwitchRunAtStartup.IsOn = MainWindow.Settings.Startup.IsFoldAtStartup;
            ToggleSwitchFoldAtStartup.IsOn = MainWindow.Settings.Startup.IsFoldAtStartup;

            // 初始化自动更新时间段ComboBox
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
            ComboBoxMatrixTransformCenterPoint.SelectedIndex = MainWindow.Settings.Gesture.MatrixTransformCenterPointIndex;
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
            ToggleSwitchEnableDisPlayNibModeToggle.IsOn = MainWindow.Settings.Advanced.IsEnableEdgeGestureUtil;
            ToggleSwitchIsLogEnabled.IsOn = MainWindow.Settings.Advanced.IsLogEnabled;
            ToggleSwitchIsSecondConfimeWhenShutdownApp.IsOn = MainWindow.Settings.Advanced.IsSecondConfimeWhenShutdownApp;
            ToggleSwitchIsQuadIR.IsOn = MainWindow.Settings.Advanced.IsSpecialScreen;
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
            ToggleSwitchAutoDelSavedFiles.IsOn = MainWindow.Settings.Automation.IsAutoDelSavedFiles;
            SideControlMinimumAutomationSlider.Value = MainWindow.Settings.Automation.MinimumAutomationStrokeNumber;
            AutoSavedStrokesLocation.Text = MainWindow.Settings.Automation.AutoSavedStrokesLocation;
            ComboBoxAutoDelSavedFilesDaysThreshold.SelectedIndex = GetComboBoxIndexByDaysThreshold(MainWindow.Settings.Automation.AutoDelSavedFilesDaysThreshold);

            // 关于
            AppVersionTextBlock.Text = mainWindow.AppVersion;
        }

        private void InitializeTimeComboBoxes()
        {
            for (int i = 0; i < 24; i++)
            {
                AutoUpdateWithSilenceStartTimeComboBox.Items.Add($"{i:D2}:00");
                AutoUpdateWithSilenceEndTimeComboBox.Items.Add($"{i:D2}:00");
            }
        }

        private int GetComboBoxIndexByDaysThreshold(int days)
        {
            return days switch
            {
                1 => 0,
                3 => 1,
                5 => 2,
                7 => 3,
                15 => 4,
                30 => 5,
                60 => 6,
                100 => 7,
                365 => 8,
                _ => 4
            };
        }

        // ...existing event handlers from MainWindow...
        #region 启动选项事件

        private void ToggleSwitchIsAutoUpdate_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Startup.IsAutoUpdate = ToggleSwitchIsAutoUpdate.IsOn;
            IsAutoUpdateWithSilenceBlock.Visibility = ToggleSwitchIsAutoUpdate.IsOn ? Visibility.Visible : Visibility.Collapsed;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchIsAutoUpdateWithSilence_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Startup.IsAutoUpdateWithSilence = ToggleSwitchIsAutoUpdateWithSilence.IsOn;
            AutoUpdateTimePeriodBlock.Visibility = MainWindow.Settings.Startup.IsAutoUpdateWithSilence ? Visibility.Visible : Visibility.Collapsed;
            mainWindow.SaveSettingsToFile();
        }

        private void AutoUpdateWithSilenceStartTimeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Startup.AutoUpdateWithSilenceStartTime = (string)AutoUpdateWithSilenceStartTimeComboBox.SelectedItem;
            mainWindow.SaveSettingsToFile();
        }

        private void AutoUpdateWithSilenceEndTimeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Startup.AutoUpdateWithSilenceEndTime = (string)AutoUpdateWithSilenceEndTimeComboBox.SelectedItem;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchRunAtStartup_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            if (ToggleSwitchRunAtStartup.IsOn)
            {
                mainWindow.StartAutomaticallyDel("InkCanvas");
                mainWindow.StartAutomaticallyDel("Ink Canvas Annotation");
                mainWindow.StartAutomaticallyCreate("InkCanvasforDrawing");
            }
            else
            {
                mainWindow.StartAutomaticallyDel("InkCanvas");
                mainWindow.StartAutomaticallyDel("Ink Canvas Annotation");
                mainWindow.StartAutomaticallyDel("InkCanvasforDrawing");
            }
        }

        private void ToggleSwitchFoldAtStartup_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Startup.IsFoldAtStartup = ToggleSwitchFoldAtStartup.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        #endregion

        #region 画板选项事件

        private void ToggleSwitchCompressPicturesUploaded_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Canvas.IsCompressPicturesUploaded = ToggleSwitchCompressPicturesUploaded.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchShowCursor_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Canvas.IsShowCursor = ToggleSwitchShowCursor.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ComboBoxEraserSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Canvas.EraserSize = ComboBoxEraserSize.SelectedIndex;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchHideStrokeWhenSelecting_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Canvas.HideStrokeWhenSelecting = ToggleSwitchHideStrokeWhenSelecting.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ComboBoxHyperbolaAsymptoteOption_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Canvas.HyperbolaAsymptoteOption = (OptionalOperation)ComboBoxHyperbolaAsymptoteOption.SelectedIndex;
            mainWindow.SaveSettingsToFile();
        }

        #endregion

        #region 手势选项事件

        private void ComboBoxMatrixTransformCenterPoint_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Gesture.MatrixTransformCenterPointIndex = ComboBoxMatrixTransformCenterPoint.SelectedIndex;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoSwitchTwoFingerGesture_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Gesture.AutoSwitchTwoFingerGesture = ToggleSwitchAutoSwitchTwoFingerGesture.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchEnableTwoFingerRotation_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Gesture.IsEnableTwoFingerRotationOnSelection = ToggleSwitchEnableTwoFingerRotationOnSelection.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        #endregion

        #region 墨迹识别事件

        private void ToggleSwitchEnableInkToShape_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.InkToShape.IsInkToShapeEnabled = ToggleSwitchEnableInkToShape.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        #endregion

        #region 外观选项事件

        private void ComboBoxTheme_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.Theme = ComboBoxTheme.SelectedIndex;
            mainWindow.SystemEvents_UserPreferenceChanged(null, null);
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchEnableDisPlayFloatBarText_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.IsEnableDisPlayFloatBarText = ToggleSwitchEnableDisPlayFloatBarText.IsOn;
            mainWindow.SaveSettingsToFile();
            mainWindow.LoadSettings();
        }

        private void ToggleSwitchEnableDisPlayNibModeToggle_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.IsEnableDisPlayNibModeToggler = ToggleSwitchEnableDisPlayNibModeToggle.IsOn;
            mainWindow.SaveSettingsToFile();
            mainWindow.LoadSettings();
        }

        private void ToggleSwitchIsColorfulViewboxFloatingBar_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.IsColorfulViewboxFloatingBar = ToggleSwitchColorfulViewboxFloatingBar.IsOn;
            mainWindow.SaveSettingsToFile();
            mainWindow.LoadSettings();
        }

        private void SliderFloatingBarBottomMargin_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.FloatingBarBottomMargin = e.NewValue;
            mainWindow.ViewboxFloatingBarMarginAnimation();
            mainWindow.SaveSettingsToFile();
        }

        private void SliderFloatingBarScale_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.FloatingBarScale = e.NewValue;
            mainWindow.ApplyScaling();
            mainWindow.SaveSettingsToFile();
        }

        private void SliderBlackboardScale_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.BlackboardScale = e.NewValue;
            mainWindow.ApplyScaling();
            mainWindow.SaveSettingsToFile();
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

        #endregion

        #region PPT 选项事件

        private void ToggleSwitchShowPPTNavigationPanelBottom_OnToggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsShowBottomPPTNavigationPanel = ToggleSwitchShowPPTNavigationPanelBottom.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchShowButtonPPTNavigationBottom_OnToggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsShowPPTNavigationBottom = ToggleSwitchShowButtonPPTNavigationBottom.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchShowPPTNavigationPanelSide_OnToggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsShowSidePPTNavigationPanel = ToggleSwitchShowPPTNavigationPanelSide.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchShowButtonPPTNavigationSides_OnToggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsShowPPTNavigationSides = ToggleSwitchShowButtonPPTNavigationSides.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchSupportPowerPoint_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.PowerPointSupport = ToggleSwitchSupportPowerPoint.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchSupportWPS_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsSupportWPS = ToggleSwitchSupportWPS.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchShowCanvasAtNewSlideShow_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsShowCanvasAtNewSlideShow = ToggleSwitchShowCanvasAtNewSlideShow.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchEnableTwoFingerGestureInPresentationMode_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsEnableTwoFingerGestureInPresentationMode = ToggleSwitchEnableTwoFingerGestureInPresentationMode.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchEnableFingerGestureSlideShowControl_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsEnableFingerGestureSlideShowControl = ToggleSwitchEnableFingerGestureSlideShowControl.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoSaveScreenShotInPowerPoint_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsAutoSaveScreenShotInPowerPoint = ToggleSwitchAutoSaveScreenShotInPowerPoint.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoSaveStrokesInPowerPoint_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsAutoSaveStrokesInPowerPoint = ToggleSwitchAutoSaveStrokesInPowerPoint.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchNotifyPreviousPage_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsNotifyPreviousPage = ToggleSwitchNotifyPreviousPage.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchNotifyHiddenPage_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsNotifyHiddenPage = ToggleSwitchNotifyHiddenPage.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchNotifyAutoPlayPresentation_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.PowerPointSettings.IsNotifyAutoPlayPresentation = ToggleSwitchNotifyAutoPlayPresentation.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        #endregion

        #region 高级选项事件

        private void ToggleSwitchIsSpecialScreen_OnToggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.IsSpecialScreen = ToggleSwitchIsSpecialScreen.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void TouchMultiplierSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.TouchMultiplier = e.NewValue;
            mainWindow.SaveSettingsToFile();
        }

        private void NibModeBoundsWidthSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.NibModeBoundsWidth = e.NewValue;
            mainWindow.SaveSettingsToFile();
        }

        private void FingerModeBoundsWidthSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.FingerModeBoundsWidth = e.NewValue;
            mainWindow.SaveSettingsToFile();
        }

        private void NibModeBoundsWidthThresholdValueSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.NibModeBoundsWidthThresholdValue = e.NewValue;
            mainWindow.SaveSettingsToFile();
        }

        private void FingerModeBoundsWidthThresholdValueSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.FingerModeBoundsWidthThresholdValue = e.NewValue;
            mainWindow.SaveSettingsToFile();
        }

        private void NibModeBoundsWidthEraserSizeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.NibModeBoundsWidthEraserSize = e.NewValue;
            mainWindow.SaveSettingsToFile();
        }

        private void FingerModeBoundsWidthEraserSizeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.FingerModeBoundsWidthEraserSize = e.NewValue;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchIsLogEnabled_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.IsLogEnabled = ToggleSwitchIsLogEnabled.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchIsSecondConfimeWhenShutdownApp_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.IsSecondConfimeWhenShutdownApp = ToggleSwitchIsSecondConfimeWhenShutdownApp.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchIsQuadIR_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.IsSpecialScreen = ToggleSwitchIsQuadIR.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchIsEnableEdgeGestureUtil_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.IsEnableEdgeGestureUtil = ToggleSwitchIsEnableEdgeGestureUtil.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void BorderCalculateMultiplier_TouchDown(object sender, TouchEventArgs e)
        {
            // 计算触摸倍数逻辑
            mainWindow.BorderCalculateMultiplier_TouchDown(sender, e);
        }

        #endregion

        #region 自动选项事件

        private void ToggleSwitchAutoFoldInEasiNote_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInEasiNote = ToggleSwitchAutoFoldInEasiNote.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInEasiNoteIgnoreDesktopAnno_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInEasiNoteIgnoreDesktopAnno = ToggleSwitchAutoFoldInEasiNoteIgnoreDesktopAnno.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInEasiCamera_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInEasiCamera = ToggleSwitchAutoFoldInEasiCamera.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInEasiNote3C_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInEasiNote3C = ToggleSwitchAutoFoldInEasiNote3C.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInSeewoPincoTeacher_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInSeewoPincoTeacher = ToggleSwitchAutoFoldInSeewoPincoTeacher.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInHiteTouchPro_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInHiteTouchPro = ToggleSwitchAutoFoldInHiteTouchPro.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInHiteCamera_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInHiteCamera = ToggleSwitchAutoFoldInHiteCamera.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInWxBoardMain_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInWxBoardMain = ToggleSwitchAutoFoldInWxBoardMain.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInOldZyBoard_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInOldZyBoard = ToggleSwitchAutoFoldInOldZyBoard.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInMSWhiteboard_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInMSWhiteboard = ToggleSwitchAutoFoldInMSWhiteboard.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInPPTSlideShow_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInPPTSlideShow = ToggleSwitchAutoFoldInPPTSlideShow.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoKillPptService_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoKillPptService = ToggleSwitchAutoKillPptService.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoKillEasiNote_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoKillEasiNote = ToggleSwitchAutoKillEasiNote.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoSaveStrokesAtClear_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoSaveStrokesAtClear = ToggleSwitchAutoSaveStrokesAtClear.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchSaveScreenshotsInDateFolders_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsSaveScreenshotsInDateFolders = ToggleSwitchSaveScreenshotsInDateFolders.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoSaveStrokesAtScreenshot_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoSaveStrokesAtScreenshot = ToggleSwitchAutoSaveStrokesAtScreenshot.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoDelSavedFiles_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoDelSavedFiles = ToggleSwitchAutoDelSavedFiles.IsOn;
            mainWindow.SaveSettingsToFile();
        }

        private void SideControlMinimumAutomationSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.MinimumAutomationStrokeNumber = (int)e.NewValue;
            mainWindow.SaveSettingsToFile();
        }

        private void AutoSavedStrokesLocationTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.AutoSavedStrokesLocation = AutoSavedStrokesLocation.Text;
            mainWindow.SaveSettingsToFile();
        }

        private void AutoSavedStrokesLocationButton_Click(object sender, RoutedEventArgs e)
        {
            mainWindow.AutoSavedStrokesLocationButton_Click(sender, e);
        }

        private void SetAutoSavedStrokesLocationToDiskDButton_Click(object sender, RoutedEventArgs e)
        {
            mainWindow.SetAutoSavedStrokesLocationToDiskDButton_Click(sender, e);
            AutoSavedStrokesLocation.Text = MainWindow.Settings.Automation.AutoSavedStrokesLocation;
        }

        private void SetAutoSavedStrokesLocationToDocumentFolderButton_Click(object sender, RoutedEventArgs e)
        {
            mainWindow.SetAutoSavedStrokesLocationToDocumentFolderButton_Click(sender, e);
            AutoSavedStrokesLocation.Text = MainWindow.Settings.Automation.AutoSavedStrokesLocation;
        }

        private void ComboBoxAutoDelSavedFilesDaysThreshold_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            int days = ComboBoxAutoDelSavedFilesDaysThreshold.SelectedIndex switch
            {
                0 => 1,
                1 => 3,
                2 => 5,
                3 => 7,
                4 => 15,
                5 => 30,
                6 => 60,
                7 => 100,
                8 => 365,
                _ => 15
            };
            MainWindow.Settings.Automation.AutoDelSavedFilesDaysThreshold = days;
            mainWindow.SaveSettingsToFile();
        }

        #endregion

        #region 关于及其他事件

        private void HyperlinkSourceToPresentRepository_Click(object sender, RoutedEventArgs e)
        {
            mainWindow.HyperlinkSourceToPresentRepository_Click(sender, e);
        }

        private void HyperlinkSourceToOringinalRepository_Click(object sender, RoutedEventArgs e)
        {
            mainWindow.HyperlinkSourceToOringinalRepository_Click(sender, e);
        }

        private void BtnRestart_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(System.Windows.Forms.Application.ExecutablePath, "-m");
            mainWindow.CloseIsFromButton = true;
            Application.Current.Shutdown();
        }

        private void BtnResetToSuggestion_Click(object sender, RoutedEventArgs e)
        {
            mainWindow.BtnResetToSuggestion_Click(null, null);
            LoadSettings();
        }

        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            mainWindow.CloseIsFromButton = true;
            Application.Current.Shutdown();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void SCManipulationBoundaryFeedback(object sender, ManipulationBoundaryFeedbackEventArgs e)
        {
            e.Handled = true;
        }

        #endregion
    }
}
