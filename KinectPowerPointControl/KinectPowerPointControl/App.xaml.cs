// Standard .NET Framework namespaces
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
// WPF Application namespace
using System.Windows;

namespace KinectPowerPointControl
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// 
    /// The App class represents the WPF application itself.
    /// This is the entry point for the application lifecycle.
    /// 
    /// In this implementation, the App class is minimal - it simply inherits from Application
    /// and relies on the XAML (App.xaml) to specify the startup window (MainWindow).
    /// 
    /// Application-level events (like Startup, Exit) can be handled here if needed.
    /// Currently, exit cleanup is handled in MainWindow's Current_Exit event handler.
    /// </summary>
    public partial class App : Application
    {
        // This class is intentionally minimal - most application logic is in MainWindow
        // You could add application-level initialization here if needed, such as:
        // - Global exception handling
        // - Application-wide settings
        // - Resource loading
    }
}
