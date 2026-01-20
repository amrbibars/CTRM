using System.Text;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CTRM;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        // Primary screen size (WPF gives you DPI-aware values)
        double sw = SystemParameters.PrimaryScreenWidth;
        double sh = SystemParameters.PrimaryScreenHeight;

        // Screen aspect ratio
        double screenAspect = sw / sh;

        // How much of the screen you want to use
        const double scale = 0.6; // 60%

        // Start with width based on scale, compute matching height
        double w = sw * scale;
        double h = w / screenAspect;

        // If height exceeds our scaled limit, clamp and recompute width
        double maxH = sh * scale;
        if (h > maxH)
        {
            h = maxH;
            w = h * screenAspect;
        }

        this.Width = w;
        this.Height = h;

        this.MinWidth = 360;
        this.MinHeight = 470;
    }
    // Event handler for the button click to open StaffingAndScheduling window
    private void OpenWindow_Click(object sender, RoutedEventArgs e)
    {
        var fe = (FrameworkElement)sender;

        // Tag must contain a Type (e.g., typeof(StaffingAndScheduling))
        if (fe.Tag is not Type winType || !typeof(Window).IsAssignableFrom(winType))
        {
            MessageBox.Show("Button Tag must be set to a window type, e.g. {x:Type local:StaffingAndScheduling}.");
            return;
        }

        // Reuse existing instance if already opened (per-owner singleton)
        var existing = this.OwnedWindows.Cast<Window>().FirstOrDefault(w => w.GetType() == winType);
        if (existing is null)
        {
            var w = (Window)Activator.CreateInstance(winType)!;
            w.Owner = this;
            w.ShowInTaskbar = true;
            // 30% offset from MainWindow (you asked for ~15% earlier; adjust as you like)
            ShowChildWithOwnerOffset(w, this, 0.30);
        }
        else
        {
            if (existing.WindowState == WindowState.Minimized) existing.WindowState = WindowState.Normal;
            existing.Activate();
            existing.Focus();
        }
    }
    // Make the sub windows appear offset from the main window
    private static void ShowChildWithOwnerOffset(Window child, Window owner, double fraction = 0.15)
    {
        // Make it owned so it appears in front of the main window and minimizes with it
        child.Owner = owner;
        child.WindowStartupLocation = WindowStartupLocation.Manual;

        // Target position = owner's top-left + 15% of owner's width/height
        double targetLeft = owner.Left + (owner.Width * fraction);
        double targetTop = owner.Top + (owner.Height * fraction);

        // Determine child size (fallbacks if not explicitly set)
        double childWidth = double.IsNaN(child.Width) ? (child.ActualWidth > 0 ? child.ActualWidth : 800) : child.Width;
        double childHeight = double.IsNaN(child.Height) ? (child.ActualHeight > 0 ? child.ActualHeight : 600) : child.Height;

        // Clamp to the virtual desktop so it doesn't go off-screen (multi-monitor safe)
        double vsLeft = SystemParameters.VirtualScreenLeft;
        double vsTop = SystemParameters.VirtualScreenTop;
        double vsRight = vsLeft + SystemParameters.VirtualScreenWidth;
        double vsBottom = vsTop + SystemParameters.VirtualScreenHeight;

        if (targetLeft + childWidth > vsRight) targetLeft = vsRight - childWidth;
        if (targetTop + childHeight > vsBottom) targetTop = vsBottom - childHeight;
        if (targetLeft < vsLeft) targetLeft = vsLeft;
        if (targetTop < vsTop) targetTop = vsTop;

        child.Left = targetLeft;
        child.Top = targetTop;

        child.Show();         // or child.ShowDialog(); if you want it modal
        child.Activate();     // ensure it gets focus
    }

}