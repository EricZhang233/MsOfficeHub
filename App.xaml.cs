using Microsoft.UI.Xaml;

namespace MsOfficeHub
{
    public partial class App : Application
    {
        private Window? _window;

        public App()
        {
            InitializeComponent();
        }

        protected override void OnLaunched(LaunchActivatedEventArgs args)
        {
            ConfigService.Initialize();
            _window = new MainWindow();
            _window.Activate();
        }
    }
}
