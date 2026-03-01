using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using LabWork18.ModelClasses;
using LabWork18.HelperClasses;
using Microsoft.Win32;

namespace LabWork18
{
    public partial class MainWindow : Window
    {
        private ModelEF model;
        private List<Users> users;
        private List<Auto> autos;

        public MainWindow()
        {
            InitializeComponent();

            model = new ModelEF();
            users = new List<Users>();
            autos = new List<Auto>();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            users = model.Users.ToList();
            autos = model.Auto.ToList();

            comboBoxUsers.ItemsSource = users.Select(u => u.FullName);
        }

        private void comboBoxUsers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comboBoxUsers.SelectedIndex != -1)
            {
                string selectedUser = comboBoxUsers.SelectedItem.ToString();
                int selectedUserId = users.FirstOrDefault(u => u.FullName == selectedUser).ID;

                var userAutos = autos.Where(a => a.UserID == selectedUserId).ToList();

                comboBoxAutos.ItemsSource = userAutos.Select(a => a.Model);
            }
        }

        private void SaveDocument_Click(object sender, RoutedEventArgs e)
        {
            if (comboBoxUsers.SelectedIndex == -1 || comboBoxAutos.SelectedIndex == -1)
            {
                MessageBox.Show("Пожалуйста, выберите пользователя и автомобиль.");
                return;
            }

            string selectedUserFullName = comboBoxUsers.SelectedItem.ToString();
            Users selectedUser = users.FirstOrDefault(u => u.FullName == selectedUserFullName);

            string selectedAutoModel = comboBoxAutos.SelectedItem.ToString();
            Auto selectedAuto = autos.FirstOrDefault(a => a.Model == selectedAutoModel && a.UserID == selectedUser.ID);

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Word Document (*.docx)|*.docx",
                DefaultExt = ".docx",
                FileName = "Новый документ.docx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                string filename = "Шаблон.docx";
                WordHelper wordHelper = new WordHelper(filename);

                Dictionary<string, string> items = new Dictionary<string, string>
                {
                    { "{FullName}", selectedUser.FullName },
                    { "{Adress}", selectedUser.Adress },
                    { "{PSeria}", selectedUser.PSeria.ToString() },
                    { "{PNumber}", selectedUser.PNumber.ToString() },
                    { "{PVidan}", selectedUser.PVidan },
                    { "{VIN}", selectedAuto.VIN },
                    { "{SeriaPasport}", selectedAuto.SeriaPasport },
                    { "{NumbePasport}", selectedAuto.NumbePasport },
                    { "{VidanPasport}", selectedAuto.VidanPasport },
                    { "{Model}", selectedAuto.Model },
                    { "{TypeV}", selectedAuto.TypeV },
                    { "{Category}", selectedAuto.Category },
                    { "{RegistrationMark}", selectedAuto.RegistrationMark },
                    { "{YearOfRelease}", selectedAuto.YearOfRelease.HasValue ? selectedAuto.YearOfRelease.Value.ToShortDateString() : string.Empty },
                    { "{EngineNumber}", selectedAuto.EngineNumber },
                    { "{Chassis}", selectedAuto.Chassis },
                    { "{Bodywork}", selectedAuto.Bodywork },
                    { "{Color}", selectedAuto.Color }
                };

                wordHelper.Process(items, saveFileDialog.FileName);

                MessageBox.Show("Документ успешно сохранён.");
            }
        }
    }
}