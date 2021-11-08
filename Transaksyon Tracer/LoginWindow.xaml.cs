using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;
using Path = System.IO.Path;

namespace Transaksyon_Tracer
{
    /// <summary>
    /// Interaction logic for login_form.xaml
    /// </summary>
    public partial class LoginWindow : Window
    {
        public static string getFirstname = "";
        readonly string sqliteConnectionString = @"Data Source=//DESKTOP-7DENF7N\Transaksyon Tracer\Database\TransaksyonTracerDatabase.sqlite;Version=3;";

        public LoginWindow()
        {
            InitializeComponent();
        }
        /////////////////////////////////////////////////////////////////////////////LOAD/CLOSE FORM
        
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            userName.Focus();
            CreateFolder();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            //Application.Current.Shutdown();
        }

        /////////////////////////////////////////////////////////////////////////////OTHER METHOD
        
        private void CreateFolder()
        {
            string root = @"C:\Transaksyon Tracer";
            //string rootD = @"D:\Transaksyon Tracer";
            //if (Directory.Exists(root) || Directory.Exists(rootD))
            if (Directory.Exists(root))
            {
                return;
            }
            else
            {
                Directory.CreateDirectory(root);
                Directory.CreateDirectory(Path.Combine(root, @"Images"));
                Directory.CreateDirectory(Path.Combine(root, @"Documents"));
                Directory.CreateDirectory(Path.Combine(root, @"Database"));
                Directory.CreateDirectory(Path.Combine(root, @"Backup\Database"));

                //D
                //Directory.CreateDirectory(rootD);
                //Directory.CreateDirectory(Path.Combine(rootD, @"Backup\Database"));
            }
        }

        private void Clear(DependencyObject obj)
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {

                if (obj is TextBox box)
                    box.Text = string.Empty;
                if (obj is CheckBox checkbox)
                    checkbox.IsChecked = false;
                if (obj is ComboBox combobox)
                    combobox.Text = string.Empty;
                if (obj is RadioButton radioButton)
                    radioButton.IsChecked = false;
                if (obj is PasswordBox passwordbox)
                    passwordbox.Password = string.Empty;

                Clear(VisualTreeHelper.GetChild(obj, i));
            }
        }

        private void WaitCursor()
        {
            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait; 
        }

        private void NormalCursor()
        {
            Mouse.OverrideCursor = null;
        }

        /////////////////////////////////////////////////////////////////////////////Button

        int attempts = 5;

        private void Login_Button_Click(object sender, RoutedEventArgs e)
        {
            ////Regex regex = new Regex("[^a-zA-Z0-9_.]+");
            Regex rx = new Regex(@"^[^a-zA-ZñÑ0-9_.]+");

            if (userName.Text == string.Empty || userPassword.Password == string.Empty)
            {
                MessageBox.Show("Please enter valid username or password!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
                Clear(this);
                return;
            }

            else if (rx.IsMatch(userName.Text) || rx.IsMatch(userPassword.Password))
            {
                MessageBox.Show("Not accepting special character's!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
                Clear(this);
                return;
            }
            else if (login_checkboxAdmin.IsChecked == false && login_checkboxStandard.IsChecked == false)
            {
                MessageBox.Show("Please select atleast 1 user type!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }
            else if (login_checkboxAdmin.IsChecked == true)
            {
                try
                {
                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        using (SQLiteCommand cmd_username = con.CreateCommand())
                        {
                            con.Open();
                            cmd_username.CommandType = CommandType.Text;
                            cmd_username.CommandText = "Select * from user_account where username = '" + userName.Text.Trim() + "'and password = '" + userPassword.Password.Trim() + "'";
                            SQLiteDataReader sdr_admin;
                            sdr_admin = cmd_username.ExecuteReader();
                            int count_admin = 0;
                            string userRole_admin = string.Empty;

                            while (sdr_admin.Read())
                            {
                                count_admin++;
                                userRole_admin = sdr_admin["userType"].ToString();
                            }
                            if (attempts > 1 && userRole_admin != "Administrator")
                            {
                                attempts--;
                                MessageBox.Show("Invalid Username or Password! " + attempts + " attempts left!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
                                Clear(this);
                                //return;
                            }
                            else if (userRole_admin == "Administrator")
                            {
                                getFirstname = userName.Text;
                                WaitCursor();
                                AdminWindow win = new AdminWindow();
                                win.Show();
                                this.Close();
                                NormalCursor();
                            }
                            else if (attempts > 0)
                            {
                                MessageBox.Show("Access denied, the application will close!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
                                Application.Current.Shutdown();
                            }

                        }

                    }
                }
                catch (Exception)
                {
                    throw;
                }
            }
            else if (login_checkboxStandard.IsChecked == true)
            {
                try
                {

                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        using (SQLiteCommand cmd_staff = con.CreateCommand())
                        {
                            con.Open();
                            cmd_staff.CommandType = CommandType.Text;
                            cmd_staff.CommandText = "Select * from user_account where username = '" + userName.Text.Trim() + "' and password = '" + userPassword.Password.Trim() + "'";

                            SQLiteDataReader sdr_standard;
                            sdr_standard = cmd_staff.ExecuteReader();
                            int count_standard = 0;
                            string userRole_standard = string.Empty;
                            while (sdr_standard.Read())
                            {
                                count_standard++;
                                userRole_standard = sdr_standard["userType"].ToString();
                            }
                            if (attempts > 1 && userRole_standard != "Standard" && userRole_standard != "Administrator")
                            {
                                attempts--;
                                MessageBox.Show("Invalid Username or Password! " + attempts + " attempts left!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
                                Clear(this);
                                //return;
                            }

                            else if (userRole_standard == "Standard" || userRole_standard == "Administrator")
                            {
                                getFirstname = userName.Text;
                                WaitCursor();
                                MainWindow win = new MainWindow();
                                win.Show();
                                this.Close();
                                NormalCursor();
                            }

                            else if (attempts > 0)
                            {
                                MessageBox.Show("Access denied, the application will close!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
                                Application.Current.Shutdown();
                            }

                        }

                    }
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }
        private void ButtonBack_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Are you sure, you want to Exit?", "Transaksyon Tracer", MessageBoxButton.YesNo, MessageBoxImage.Information, MessageBoxResult.No);
            if (result == MessageBoxResult.Yes)
            {
               Application.Current.Shutdown();
            }
            else
            {
                return;
            }
        }

        /////////////////////////////////////////////////////////////////////////////TEXTCHANGED

        private void UserName_TextChanged(object sender, TextChangedEventArgs e)
        {
            userName.Focus();
        }

        private void ButtonMinimize_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }
    }
}