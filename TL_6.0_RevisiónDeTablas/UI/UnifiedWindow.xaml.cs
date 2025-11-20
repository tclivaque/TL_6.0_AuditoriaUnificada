using System.Windows;
using System.Windows.Controls;

namespace TL60_AuditoriaUnificada.UI
{
    public partial class UnifiedWindow : Window
    {
        public UnifiedWindow(
            UserControl metradosPlugin,
            UserControl cobiePlugin,
            UserControl uniclassPlugin,
            UserControl tablasPlugin,
            UserControl projectBrowserPlugin)
        {
            InitializeComponent();

            // Cargar los plugins en sus respectivos contenedores
            if (metradosPlugin != null)
            {
                MetradosPluginContainer.Content = metradosPlugin;
            }

            if (cobiePlugin != null)
            {
                CobiePluginContainer.Content = cobiePlugin;
            }

            if (uniclassPlugin != null)
            {
                UniclassPluginContainer.Content = uniclassPlugin;
            }

            if (tablasPlugin != null)
            {
                TablasPluginContainer.Content = tablasPlugin;
            }

            if (projectBrowserPlugin != null)
            {
                ProjectBrowserPluginContainer.Content = projectBrowserPlugin;
            }

            // Seleccionar por defecto la primera pestaña
            PluginsTabControl.SelectedIndex = 0;
        }
    }
}
