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
            ToggleSwitchRunAtStartup.IsOn = MainWindow.Settings.Startup.IsEnableNibMode;
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
            ComboBoxMatrixTransformCenterPoint.SelectedIndex = 0; // 默认值
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
            SideControlMinimumAutomationSlider.Value = MainWindow.Settings.Automation.MinimumAutomationStrokeNumber;
            AutoSavedStrokesLocation.Text = MainWindow.Settings.Automation.AutoSavedStrokesLocation;
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

        #region 启动选项事件

        private void ToggleSwitchIsAutoUpdate_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Startup.IsAutoUpdate = ToggleSwitchIsAutoUpdate.IsOn;
            IsAutoUpdateWithSilenceBlock.Visibility = ToggleSwitchIsAutoUpdate.IsOn ? Visibility.Visible : Visibility.Collapsed;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchIsAutoUpdateWithSilence_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Startup.IsAutoUpdateWithSilence = ToggleSwitchIsAutoUpdateWithSilence.IsOn;
            AutoUpdateTimePeriodBlock.Visibility = MainWindow.Settings.Startup.IsAutoUpdateWithSilence ? Visibility.Visible : Visibility.Collapsed;
            MainWindow.SaveSettingsToFile();
        }

        private void AutoUpdateWithSilenceStartTimeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Startup.AutoUpdateWithSilenceStartTime = (string)AutoUpdateWithSilenceStartTimeComboBox.SelectedItem;
            MainWindow.SaveSettingsToFile();
        }

        private void AutoUpdateWithSilenceEndTimeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
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

        #endregion

        #region 画板选项事件

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

        private void ComboBoxEraserSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Canvas.EraserSize = ComboBoxEraserSize.SelectedIndex;
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

        #endregion

        #region 手势选项事件

        private void ComboBoxMatrixTransformCenterPoint_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Gesture.MatrixTransformCenterPointIndex = ComboBoxMatrixTransformCenterPoint.SelectedIndex;
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

        #endregion

        #region 墨迹识别事件

        private void ToggleSwitchEnableInkToShape_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.InkToShape.IsInkToShapeEnabled = ToggleSwitchEnableInkToShape.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        #endregion

        #region 外观选项事件

        private void ComboBoxTheme_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.Theme = ComboBoxTheme.SelectedIndex;
            mainWindow.SystemEvents_UserPreferenceChanged(null, null);
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchEnableDisPlayFloatBarText_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.IsEnableDisPlayFloatBarText = ToggleSwitchEnableDisPlayFloatBarText.IsOn;
            MainWindow.SaveSettingsToFile();
            mainWindow.LoadSettings();
        }

        private void ToggleSwitchEnableDisPlayNibModeToggle_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.IsEnableDisPlayNibModeToggler = ToggleSwitchEnableDisPlayNibModeToggle.IsOn;
            MainWindow.SaveSettingsToFile();
            mainWindow.LoadSettings();
        }

        private void ToggleSwitchIsColorfulViewboxFloatingBar_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.IsColorfulViewboxFloatingBar = ToggleSwitchColorfulViewboxFloatingBar.IsOn;
            MainWindow.SaveSettingsToFile();
            mainWindow.LoadSettings();
        }

        private void SliderFloatingBarBottomMargin_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.FloatingBarBottomMargin = e.NewValue;
            mainWindow.ViewboxFloatingBarMarginAnimation();
            MainWindow.SaveSettingsToFile();
        }

        private void SliderFloatingBarScale_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.FloatingBarScale = e.NewValue;
            mainWindow.ApplyScaling();
            MainWindow.SaveSettingsToFile();
        }

        private void SliderBlackboardScale_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Appearance.BlackboardScale = e.NewValue;
            mainWindow.ApplyScaling();
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

        #endregion

        #region PPT 选项事件

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

        #endregion

        #region 高级选项事件

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
            MainWindow.Settings.Advanced.NibModeBoundsWidth = e.NewValue;
            MainWindow.SaveSettingsToFile();
        }

        private void FingerModeBoundsWidthSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.FingerModeBoundsWidth = e.NewValue;
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

        private void ToggleSwitchIsQuadIR_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.IsSpecialScreen = ToggleSwitchIsQuadIR.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchIsEnableEdgeGestureUtil_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Advanced.IsEnableEdgeGestureUtil = ToggleSwitchIsEnableEdgeGestureUtil.IsOn;
            MainWindow.SaveSettingsToFile();
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
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInEasiNoteIgnoreDesktopAnno_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInEasiNoteIgnoreDesktopAnno = ToggleSwitchAutoFoldInEasiNoteIgnoreDesktopAnno.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInEasiCamera_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInEasiCamera = ToggleSwitchAutoFoldInEasiCamera.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInEasiNote3C_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInEasiNote3C = ToggleSwitchAutoFoldInEasiNote3C.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInSeewoPincoTeacher_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInSeewoPincoTeacher = ToggleSwitchAutoFoldInSeewoPincoTeacher.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInHiteTouchPro_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInHiteTouchPro = ToggleSwitchAutoFoldInHiteTouchPro.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInHiteCamera_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInHiteCamera = ToggleSwitchAutoFoldInHiteCamera.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInWxBoardMain_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInWxBoardMain = ToggleSwitchAutoFoldInWxBoardMain.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInOldZyBoard_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInOldZyBoard = ToggleSwitchAutoFoldInOldZyBoard.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInMSWhiteboard_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInMSWhiteboard = ToggleSwitchAutoFoldInMSWhiteboard.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoFoldInPPTSlideShow_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoFoldInPPTSlideShow = ToggleSwitchAutoFoldInPPTSlideShow.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoKillPptService_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoKillPptService = ToggleSwitchAutoKillPptService.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoKillEasiNote_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoKillEasiNote = ToggleSwitchAutoKillEasiNote.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoSaveStrokesAtClear_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoSaveStrokesAtClear = ToggleSwitchAutoSaveStrokesAtClear.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchSaveScreenshotsInDateFolders_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsSaveScreenshotsInDateFolders = ToggleSwitchSaveScreenshotsInDateFolders.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoSaveStrokesAtScreenshot_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoSaveStrokesAtScreenshot = ToggleSwitchAutoSaveStrokesAtScreenshot.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void ToggleSwitchAutoDelSavedFiles_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.IsAutoDelSavedFiles = ToggleSwitchAutoDelSavedFiles.IsOn;
            MainWindow.SaveSettingsToFile();
        }

        private void SideControlMinimumAutomationSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.MinimumAutomationStrokeNumber = (int)e.NewValue;
            MainWindow.SaveSettingsToFile();
        }

        private void AutoSavedStrokesLocationTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!isLoaded) return;
            MainWindow.Settings.Automation.AutoSavedStrokesLocation = AutoSavedStrokesLocation.Text;
            MainWindow.SaveSettingsToFile();
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
            MainWindow.SaveSettingsToFile();
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
            MainWindow.CloseIsFromButton = true;
            Application.Current.Shutdown();
        }

        private void BtnResetToSuggestion_Click(object sender, RoutedEventArgs e)
        {
            mainWindow.BtnResetToSuggestion_Click(null, null);
            LoadSettings();
        }

        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.CloseIsFromButton = true;
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
