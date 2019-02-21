using System;
using System.Windows;
using System.Windows.Controls;
using System.Data.SqlClient;

namespace GLApp
{
    /// <summary>
    /// Логика взаимодействия для RegWindow.xaml
    /// </summary>
    public partial class RegWindow : Window
    {
        public RegWindow()
        {
            InitializeComponent();
        }
        
        private void RegisterBtn_Click(object sender, RoutedEventArgs e)
        {
            bool RegSuccess = false;
            try
            {
                UserData.Connection = new SqlConnection( UserData.ConnectionString + "User ID=Register;Password=RegisterPassword;" );
                UserData.OpenConnection();
                if ( string.IsNullOrEmpty( Name.Text ) )
                {
                    Name.Text = "UnAssignedName";
                }
                if ( string.IsNullOrEmpty( Surname.Text ) )
                {

                    Surname.Text = "UnassignedSurname";
                }
                SqlCommand command = new SqlCommand("AddUser", UserData.Connection)
                {
                    CommandType = System.Data.CommandType.StoredProcedure
                };
                command.Parameters.AddWithValue("login", login.Text);
                command.Parameters.AddWithValue("pwd", passwd.Password);
                command.Parameters.AddWithValue("uname", Name.Text);
                command.Parameters.AddWithValue("surname", Surname.Text);
                command.Parameters.AddWithValue("utype", ((ComboBoxItem)UserType.SelectedItem).Tag);
                command.Parameters.AddWithValue("phone_numb", phoneNumb.Text );
                
                if (command.ExecuteNonQuery() != 2 ) 
                {
                    throw new System.Exception("Произошла ошибка, пользователь не был добавлен");
                }
                else
                {
                    RegSuccess = true;
                }
            }
            catch (Exception E)
            {
                MessageBox.Show("Ошибка регистрации\nНеправильно заданы параметры");
                Console.WriteLine(E.Message);
            }
            finally
            {
                UserData.CloseConnection();
            }
            if (RegSuccess)
            {
                UserData.pwd = passwd.Password;
                UserData.uname = login.Text;
                this.Close();
            }
        }
    }
}
