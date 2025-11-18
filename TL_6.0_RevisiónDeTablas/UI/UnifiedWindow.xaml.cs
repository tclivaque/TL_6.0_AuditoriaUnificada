using System.Windows;
using System.Windows.Controls;

namespace TL60_RevisionDeTablas.UI
{
    public partial class UnifiedWindow : Window
    {
        public UnifiedWindow(UserControl cobiePlugin, UserControl tablasPlugin)
        {
            InitializeComponent();

            // Cargar los plugins en sus respectivos contenedores
            if (cobiePlugin != null)
            {
                CobiePluginContainer.Content = cobiePlugin;
            }

            if (tablasPlugin != null)
            {
                TablasPluginContainer.Content = tablasPlugin;
            }

            // Seleccionar por defecto la pestaña de COBie (primera pestaña)
            PluginsTabControl.SelectedIndex = 0;
        }
    }
}
