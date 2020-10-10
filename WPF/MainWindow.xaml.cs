using Microsoft.Office.Interop.Access.Dao;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;

namespace WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private FileInfoModel file = new FileInfoModel();

        public MainWindow()
        {
            InitializeComponent();
            this.Title = string.Empty;
            fileText.DataContext = file;
            propsText.DataContext = file;
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
                file.Properties = ChangeAllowBypassKey(file.FileName);
            }
        }

        public class FileInfoModel : INotifyPropertyChanged
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

            private string _properties;

            public string Properties
            {
                get { return _properties; }
                set
                {
                    _properties = value;
                    OnPropertyChanged();
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;

            protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        private string ChangeAllowBypassKey(string dbPath)
        {
            var strBuilder = new StringBuilder();
            try
            {
                var dbe = new DBEngine();
                var db = dbe.OpenDatabase(dbPath);

                Property prop = db.Properties["AllowBypassKey"];

                switch (MessageBox.Show("Allow bypass key?", "Allow bypass key?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No))
                {
                    case MessageBoxResult.Yes:
                        prop.Value = true;
                        strBuilder.AppendLine("Property 'AllowBypassKey' is set to 'True'.");
                        strBuilder.AppendLine("You can access the design (developer) mode by keep pressing SHIFT key while opening the file.");
                        break;

                    case MessageBoxResult.No:
                        prop.Value = false;
                        strBuilder.AppendLine("Property 'AllowBypassKey' is set to 'False'.");
                        strBuilder.AppendLine("You can no longer use the SHIFT key to enter the design mode.");
                        break;
                }

                return ListProperties(db, strBuilder);
            }
            catch (Exception)
            {
                throw;
            }
        }

        private string ListProperties(Database db, StringBuilder strBuilder)
        {
            strBuilder.AppendLine();
            strBuilder.AppendLine("Database properties:");
            strBuilder.AppendLine(string.Concat(Enumerable.Repeat("_", 50)));
            foreach (var p in DumpProperties(db.Properties))
            {
                strBuilder.AppendLine(p);
            }

            strBuilder.AppendLine();
            strBuilder.AppendLine("Containers:");
            strBuilder.AppendLine(string.Concat(Enumerable.Repeat("_", 50)));
            foreach (Microsoft.Office.Interop.Access.Dao.Container c in db.Containers)
            {
                Console.WriteLine("{0}:{1}", c.Name, c.Owner);
                foreach (var p in DumpProperties(c.Properties))
                {
                    strBuilder.AppendLine(p);
                }
            }

            strBuilder.AppendLine();
            strBuilder.AppendLine("Documents and properties for a each container:");
            strBuilder.AppendLine(string.Concat(Enumerable.Repeat("_", 50)));
            foreach (Document d in db.Containers["Databases"].Documents)
            {
                Console.WriteLine($"{d.Name}");
                foreach (var p in DumpProperties(d.Properties))
                {
                    strBuilder.AppendLine(p);
                }
            }

            return strBuilder.ToString();
        }

        private static IEnumerable<string> DumpProperties(Properties props)
        {
            foreach (Property p in props)
            {
                object val;
                try
                {
                    val = (object)p.Value;
                }
                catch (Exception e)
                {
                    val = e.Message;
                }

                yield return $"{p.Name} ({val}) = {(DataTypeEnum)p.Type}";
            }
        }
    }
}