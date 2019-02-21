using System.Windows;
using System.Data.SqlClient;
using System;
using System.Windows.Controls;

namespace GLApp
{
    /// <summary>
    /// Логика взаимодействия для logWindow.xaml
    /// </summary>
    public partial class logWindow : Window
    {
        public logWindow()
        {
            InitializeComponent();
        }

        private void RegBtn_Click(object sender, RoutedEventArgs e)
        {
            RegWindow rw = new RegWindow();
            this.Hide();
            rw.ShowDialog();
            if ( ! string.IsNullOrEmpty( UserData.uname ) )
            {
                login.Text = UserData.uname;
                passwd.Password = UserData.pwd;
            }
            this.ShowDialog();
        }

        private void LoginBtn_Click(object sender, RoutedEventArgs e)
        {
            UserData.SetUD( login.Text, passwd.Password );
            UserData.Connection = new SqlConnection( UserData.GetCS() );
            try
            {
                UserData.OpenConnection();
                SqlCommand command = new SqlCommand( "GetType", UserData.Connection )
                {
                    CommandType = System.Data.CommandType.StoredProcedure
                };
                command.Parameters.AddWithValue( "login", login.Text );
                SqlDataReader reader = command.ExecuteReader();
                reader.Read();
                UserData.UserType = ( Int16 )reader.GetInt32( 0 );
                if ( UserData.UserType == 0 )
                {
                    throw new System.Exception( "Пользователь не найден среди участников" );
                }
                else
                {
                    command = new SqlCommand( "FindData", UserData.Connection )
                    {
                        CommandType = System.Data.CommandType.StoredProcedure
                    };
                    command.Parameters.AddWithValue( "login", login.Text );
                    try
                    {
                        reader.Close();
                        reader = command.ExecuteReader();
                        reader.Read();
                        if ( reader.HasRows )
                        {
                            Console.WriteLine( reader[0] );
                            UserData.ID = (Int32) reader[0] ;
                            if ( UserData.UserType == 1 || UserData.UserType == 2 )
                            {
                                CurrentUserData.Name = reader["Name"].ToString();
                                CurrentUserData.SName = reader["Surname"].ToString();
                                CurrentUserData.Phone = reader["phone_number"].ToString();
                                CurrentUserData.RegDate = Convert.ToDateTime( reader["Register_Date"] );
                            }
                        }
                    }
                    catch ( Exception E )
                    {
                        MessageBox.Show( E.Message );
                    }
                }
                UserData.SuccLogin = true;
                this.Close();
            }
            catch ( Exception E )
            {
                Console.WriteLine( E.Message );
                MessageBox.Show( "Ошибка входа\nНе найден пользователь с такими данными" );
            }
            finally
            {
                UserData.CloseConnection();
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if ( !UserData.SuccLogin )
            System.Environment.Exit(0);
        }

        private void textbox_KeyUp( object sender, System.Windows.Input.KeyEventArgs e )
        {
            if ( e.Key == System.Windows.Input.Key.Enter )
            {
                LoginBtn_Click( null, null );
                loginBtn.Focus();
            }
        }
    }
}
