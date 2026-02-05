using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Microsoft.UI.Xaml.Documents;
using Windows.Foundation;
using Windows.Foundation.Collections;
using GroupDocs.Viewer;
using GroupDocs.Viewer.Options;


// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace DocPic
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            // Set basic style
            // Replace the legacy title bar
            ExtendsContentIntoTitleBar = true;
            SetTitleBar(AppTitleBar);
            
            // Apply theme-aware WebView2 background
            ApplyWebViewTheme();
            
            // Listen for theme changes
            if (Content is FrameworkElement rootElement)
            {
                rootElement.ActualThemeChanged += RootElement_ActualThemeChanged;
            }
            
            // Handle WebView2 navigation to apply theme
            DocumentViewer.NavigationCompleted += DocumentViewer_NavigationCompleted;
        }

        private async void DocumentViewer_NavigationCompleted(WebView2 sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs args)
        {
            if (args.IsSuccess)
            {
                await InjectThemeStyles();
            }
        }

        private async System.Threading.Tasks.Task InjectThemeStyles()
        {
            // Get the current theme
            var theme = (Content as FrameworkElement)?.ActualTheme ?? ElementTheme.Default;
            
            if (theme == ElementTheme.Dark)
            {
                // Inject CSS to make text white in dark mode
                var css = @"
                    body, p, div, span, h1, h2, h3, h4, h5, h6, td, th, li, a {
                        color: white !important;
                    }
                    a:visited {
                        color: #bb86fc !important;
                    }
                ";
                
                var script = $@"
                    (function() {{
                        var style = document.createElement('style');
                        style.textContent = `{css}`;
                        document.head.appendChild(style);
                    }})();
                ";
                
                await DocumentViewer.ExecuteScriptAsync(script);
            }
        }

        private void RootElement_ActualThemeChanged(FrameworkElement sender, object args)
        {
            ApplyWebViewTheme();
            // Reload the document to apply new theme
            if (DocumentViewer.Source != null)
            {
                var currentSource = DocumentViewer.Source;
                DocumentViewer.Source = null;
                DocumentViewer.Source = currentSource;
            }
        }

        private void ApplyWebViewTheme()
        {
            // Get the current theme
            var theme = (Content as FrameworkElement)?.ActualTheme ?? ElementTheme.Default;
            var themeKey = theme == ElementTheme.Dark ? "Dark" : "Light";
            
            // Get the current theme color from resources
            if (Application.Current.Resources.ThemeDictionaries.TryGetValue(themeKey, out var themeDict))
            {
                if (themeDict is ResourceDictionary dict && dict.TryGetValue("WebViewBackgroundColor", out var colorObj))
                {
                    if (colorObj is Windows.UI.Color color)
                    {
                        DocumentViewer.DefaultBackgroundColor = color;
                    }
                }
            }
        }

        private async void OpenFile(object sender, RoutedEventArgs e)
        {
            var picker = new Windows.Storage.Pickers.FileOpenPicker();
            
            // You need to initialize the picker with a window handle
            var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);
            
            // Set file type filter for Word documents
            picker.FileTypeFilter.Add(".docx");
            
            var file = await picker.PickSingleFileAsync();
            if (file != null)
            {
                try
                {
                    // Load and render the document
                    await RenderDocxAsync(file.Path);
                }
                catch (Exception ex)
                {
                    // Show error dialog
                    ContentDialog errorDialog = new ContentDialog
                    {
                        Title = "错误",
                        Content = $"无法打开文档: {ex.Message}",
                        CloseButtonText = "确定",
                        XamlRoot = this.Content.XamlRoot
                    };
                    await errorDialog.ShowAsync();
                }
            }
        }

        private async System.Threading.Tasks.Task RenderDocxAsync(string filePath)
        {
            await System.Threading.Tasks.Task.Run(() =>
            {
                // Create output directory for HTML pages
                var outputDirectory = Path.Combine(Path.GetTempPath(), "DocPic_Output", Path.GetFileNameWithoutExtension(filePath));
                if (Directory.Exists(outputDirectory))
                {
                    Directory.Delete(outputDirectory, true);
                }
                Directory.CreateDirectory(outputDirectory);

                // Use GroupDocs.Viewer to render document to HTML
                using (var viewer = new Viewer(filePath))
                {
                    var viewOptions = HtmlViewOptions.ForEmbeddedResources(
                        Path.Combine(outputDirectory, "page_{0}.html")
                    );

                    viewer.View(viewOptions);
                }

                // Read the first page HTML and display it
                var firstPagePath = Path.Combine(outputDirectory, "page_1.html");
                if (File.Exists(firstPagePath))
                {
                    // Update UI on the UI thread
                    this.DispatcherQueue.TryEnqueue(async () =>
                    {
                        // Ensure WebView2 is initialized
                        await DocumentViewer.EnsureCoreWebView2Async();
                        
                        // Apply theme before loading
                        ApplyWebViewTheme();
                        
                        // Navigate to the HTML file
                        DocumentViewer.Source = new Uri(firstPagePath);
                    });
                }
            });
        }
    }
}
