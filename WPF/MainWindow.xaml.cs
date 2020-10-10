using Microsoft.Win32;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;


namespace WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        FileModel file = new FileModel();

        public MainWindow()
        {
            InitializeComponent();
            this.Title = string.Empty;
            fileText.DataContext = file;
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Find and select the MS Access database file.",
                Filter = "Access Database (*.mdb) | *.mdb|Access Database (*.accdb) | *.accdb",
                Multiselect = false
            };
            if (openFileDialog.ShowDialog() == true)
            {
                file.FileName = openFileDialog.FileName;
            }
        }

        public class FileModel : INotifyPropertyChanged
        {
            private string _fileName;

            public string FileName
            {
                get { return _fileName; }
                set
                {
                    _fileName = value;
                    OnPropertyChanged();
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;

            protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
