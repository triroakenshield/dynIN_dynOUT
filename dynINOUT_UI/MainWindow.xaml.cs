using System.Collections.ObjectModel;
using System.Windows;

namespace dynINOUT_UI
{
    /// <summary>Логика взаимодействия для MainWindow.xaml</summary>
    public partial class MainWindow
    {
        public ObservableCollection<Row> BindingList { get; private set; }

        public MainWindow()
        {
            InitializeComponent();
           //Loaded += MainWindow_Loaded;
        }

        public void AddBlockNameList(ObservableCollection<string> blockNameList)
        {
            BindingList = new ObservableCollection<Row>();

            foreach (var i in blockNameList)
                BindingList.Add(new Row { Key = i, Value=true});
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}