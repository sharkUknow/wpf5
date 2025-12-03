using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Collections.Generic; // 需要這個來使用 List<T>

namespace _2025_WpfApp4
{
    /// <summary>
    /// MyDocumentViewer.xaml 的互動邏輯
    /// </summary>
    public partial class MyDocumentViewer : Window
    {
        // 初始顏色設定
        Color fontColor = Colors.Black;
        Color fontBackgorund = Colors.Transparent;

        public MyDocumentViewer()
        {
            InitializeComponent();

            // 1. 初始化顏色選擇器
            FontColorPicker.SelectedColor = fontColor;
            backgroundColorPicker.SelectedColor = fontBackgorund;

            // 2. 載入系統字型
            foreach (FontFamily font in Fonts.SystemFontFamilies)
            {
                FontFamilyComboBox.Items.Add(font);
            }
            // 假設您的 ComboBox 名稱為 FontFamilyComboBox
            if (FontFamilyComboBox.Items.Count > 1)
            {
                FontFamilyComboBox.SelectedIndex = 1;
            }
            else if (FontFamilyComboBox.Items.Count > 0)
            {
                FontFamilyComboBox.SelectedIndex = 0;
            }


            // 3. 載入字體大小
            FontSizeComboBox.ItemsSource = new List<double>()
            {
                8,9,10,11,12,14,16,18,20,22,24,26,28,36,48,72
            };
            // 假設您的 ComboBox 名稱為 FontSizeComboBox
            if (FontSizeComboBox.Items.Count > 4)
            {
                FontSizeComboBox.SelectedIndex = 4;
            }

        }

        // --- 顏色、字體、大小變更事件 ---

        private void FontColorPicker_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color?> e)
        {
            if (e.NewValue.HasValue)
            {
                fontColor = e.NewValue.Value;
                SolidColorBrush fontBrush = new SolidColorBrush(fontColor);
                MainRichTextBox.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, fontBrush);
            }
        }

        private void FontFamilyComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (FontFamilyComboBox.SelectedItem != null)
            {
                MainRichTextBox.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, FontFamilyComboBox.SelectedItem);
            }
        }

        private void FontSizeComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (FontSizeComboBox.SelectedItem != null)
            {
                MainRichTextBox.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, FontSizeComboBox.SelectedItem);
            }
        }

        private void DisplayStatus()
        {
            // 1. 取得選取範圍
            TextSelection textSelection = MainRichTextBox.Selection;
            string statusMessage = "Format: ";

            // 2. 檢查文字樣式 (粗體、斜體、底線)
            bool isBold = (textSelection.GetPropertyValue(TextElement.FontWeightProperty) is FontWeight selectedFontWeight) && (selectedFontWeight == FontWeights.Bold);
            bool isItalic = (textSelection.GetPropertyValue(TextElement.FontStyleProperty) is FontStyle selectedFontStyle) && (selectedFontStyle == FontStyles.Italic);
            // TextDecorationsProperty 的值可能為 null (表示未設定) 或 TextDecorationCollection
            bool isUnderline = (textSelection.GetPropertyValue(Inline.TextDecorationsProperty) is TextDecorationCollection textDecorationCollection) && (textDecorationCollection == TextDecorations.Underline);

            // 3. 建構樣式狀態字串
            if (isBold) statusMessage += "Bold | ";
            if (isItalic) statusMessage += "Italic | ";
            if (isUnderline) statusMessage += "Underline | ";

            if (!isBold && !isItalic && !isUnderline) statusMessage += "Normal | ";


            // 4. 檢查字型系列 (Font Family)
            var fontFamilyPropertyValue = textSelection.GetPropertyValue(TextElement.FontFamilyProperty);
            if (fontFamilyPropertyValue != DependencyProperty.UnsetValue)
            {
                FontFamily selectedFontFamily = (FontFamily)fontFamilyPropertyValue;
                statusMessage += $"Font: {selectedFontFamily.Source} | ";
            }

            // 5. 檢查字體大小 (Font Size)
            var fontSizePropertyValue = textSelection.GetPropertyValue(TextElement.FontSizeProperty);
            if (fontSizePropertyValue != DependencyProperty.UnsetValue)
            {
                double selectedFontSize = (double)fontSizePropertyValue;
                statusMessage += $"Size: {selectedFontSize:0.##} | ";
            }

            // 6. 檢查前景顏色 (Foreground Color)
            var foregroundPropertyValue = textSelection.GetPropertyValue(TextElement.ForegroundProperty);
            if (foregroundPropertyValue != DependencyProperty.UnsetValue && foregroundPropertyValue is SolidColorBrush solidColorBrush)
            {
                Color foregroundColor = solidColorBrush.Color;
                // 格式化為 RGB 十六進制字串
                statusMessage += $"Color: #{foregroundColor.R:X2}{foregroundColor.G:X2}{foregroundColor.B:X2} | ";
            }

            // 7. 檢查背景顏色 (Background Color)
            var backgroundPropertyValue = textSelection.GetPropertyValue(TextElement.BackgroundProperty);
            if (backgroundPropertyValue != DependencyProperty.UnsetValue && backgroundPropertyValue is SolidColorBrush backgroundSolidColorBrush)
            {
                Color selectedBackgroundColor = backgroundSolidColorBrush.Color;
                // 格式化為 RGB 十六進制字串
                statusMessage += $"Background: #{selectedBackgroundColor.R:X2}{selectedBackgroundColor.G:X2}{selectedBackgroundColor.B:X2}";
            }

            // 8. 更新狀態列
            StatusTextBlock.Text = statusMessage;
        }

        // **修正過的** 背景顏色變更事件
        private void backgroundColorPicker_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color?> e)
        {
            if (e.NewValue.HasValue)
            {
                fontBackgorund = e.NewValue.Value; // 修正：更新 fontBackgorund 欄位
                SolidColorBrush backgroundBrush = new SolidColorBrush(fontBackgorund);
                MainRichTextBox.Selection.ApplyPropertyValue(TextElement.BackgroundProperty, backgroundBrush);
            }
        }

        // --- 檔案命令事件 ---

        private void NewCommand_Executed(object sender, System.Windows.Input.ExecutedRoutedEventArgs e)
        {
            MyDocumentViewer newDocumentViewer = new MyDocumentViewer();
            newDocumentViewer.Show();
        }

        private void OpenCommand_Executed(object sender, System.Windows.Input.ExecutedRoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                // 注意：您原始的 Filter 中有 "hrml" 可能是 "html" 的筆誤，我保留了您的寫法。
                Filter = "hrml(*.html)|*.html|Rich Text Format (*.rtf)|*.rtf|All files (*.*)|*.*",
                DefaultExt = ".rtf",
                Multiselect = false
            };

            if (openFileDialog.ShowDialog() == true)
            {
                TextRange range = new TextRange(MainRichTextBox.Document.ContentStart, MainRichTextBox.Document.ContentEnd);

                try
                {
                    using (FileStream fileStream = new FileStream(openFileDialog.FileName, FileMode.Open))
                    {
                        // 載入 RTF 格式
                        range.Load(fileStream, DataFormats.Rtf);
                    }
                }
                catch (IOException ex)
                {
                    MessageBox.Show($"無法開啟檔案: {ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void SaveCommand_Executed(object sender, System.Windows.Input.ExecutedRoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog()
            {
                // 注意：您原始的 Filter 中有 "hrml" 可能是 "html" 的筆誤，我保留了您的寫法。
                Filter = "hrml(*.html) | *.html | Rich Text Format(*.rtf) | *.rtf | All files(*.*) | *.* ",
                DefaultExt = ".rtf",
                AddExtension = true
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                TextRange range = new TextRange(MainRichTextBox.Document.ContentStart, MainRichTextBox.Document.ContentEnd);

                try
                {
                    using (FileStream fileStream = new FileStream(saveFileDialog.FileName, FileMode.Create))
                    {
                        // 儲存 RTF 格式
                        range.Save(fileStream, DataFormats.Rtf);
                    }
                }
                catch (IOException ex)
                {
                    MessageBox.Show($"無法儲存檔案: {ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // --- 清空按鈕事件 ---

        private void TrashButton_Click(object sender, RoutedEventArgs e)
        {
            MainRichTextBox.Document.Blocks.Clear();
        }

        // --- RichTextBox 選取範圍變更事件 (已修正 NullReferenceException) ---

        private void MainRichTextBox_SelectionChanged(object sender, RoutedEventArgs e)
        {
            // 粗體檢查
            var property_bold = MainRichTextBox.Selection.GetPropertyValue(TextElement.FontWeightProperty);
            BoldButton.IsChecked = (property_bold != DependencyProperty.UnsetValue) && (property_bold.Equals(FontWeights.Bold));

            // 斜體檢查
            var property_italic = MainRichTextBox.Selection.GetPropertyValue(TextElement.FontStyleProperty);
            ItalicButton.IsChecked = (property_italic != DependencyProperty.UnsetValue) && (property_italic.Equals(FontStyles.Italic));

            // 底線檢查
            var property_underline = MainRichTextBox.Selection.GetPropertyValue(Inline.TextDecorationsProperty);
            UnderlineButton.IsChecked = (property_underline != DependencyProperty.UnsetValue) && (property_underline.Equals(TextDecorations.Underline));

            // 字體檢查
            var property_fontfamily = MainRichTextBox.Selection.GetPropertyValue(TextElement.FontFamilyProperty);
            if (property_fontfamily != DependencyProperty.UnsetValue)
            {
                FontFamilyComboBox.SelectedItem = property_fontfamily;
            }

            // 字體大小檢查
            var property_fontsize = MainRichTextBox.Selection.GetPropertyValue(TextElement.FontSizeProperty);
            if (property_fontsize != DependencyProperty.UnsetValue)
            {
                FontSizeComboBox.SelectedItem = property_fontsize;
            }

            // 🌟 修正 NullReferenceException 發生點：字體顏色檢查
            var property_fontcolor = MainRichTextBox.Selection.GetPropertyValue(TextElement.ForegroundProperty);
            // 檢查是否為 UnsetValue 且是否能安全地轉換為 SolidColorBrush
            if (property_fontcolor != DependencyProperty.UnsetValue && property_fontcolor is SolidColorBrush colorBrush)
            {
                FontColorPicker.SelectedColor = colorBrush.Color;
            }

            // 🌟 修正 NullReferenceException 發生點：背景顏色檢查
            var property_fontbackground = MainRichTextBox.Selection.GetPropertyValue(TextElement.BackgroundProperty);
            // 檢查是否為 UnsetValue 且是否能安全地轉換為 SolidColorBrush
            if (property_fontbackground != DependencyProperty.UnsetValue && property_fontbackground is SolidColorBrush backgroundBrush)
            {
                backgroundColorPicker.SelectedColor = backgroundBrush.Color;
            }
            DisplayStatus();
        }
    }
}