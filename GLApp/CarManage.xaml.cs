using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Data.SqlClient;

namespace GLApp
{
    /// <summary>
    /// Interaction logic for CarManage.xaml
    /// </summary>
    public partial class CarManage : Window
    {
        public CarManage()
        {
            InitializeComponent();
            RefreshDataBtn_Click( null, null );
        }

        private class Car
        {
            public string CarModel;
            public string License;
            public int ID;
            public DateTime RegDate;
        }

        List<Car> Cars = new List<Car>();

        private void RefreshDataBtn_Click( object sender, RoutedEventArgs e )
        {
            try
            {
                UserData.OpenConnection();
                SqlCommand command = new SqlCommand( "GetCarsList", UserData.Connection );
                command.CommandType = System.Data.CommandType.StoredProcedure;
                command.Parameters.AddWithValue( "ID_driver", UserData.ID );
                SqlDataReader reader = command.ExecuteReader();
                CarsList.Items.Clear();
                if ( reader.HasRows )
                {
                    while ( reader.Read() )
                    {
                        Car CU = new Car();
                        CU.ID = Convert.ToInt32( reader["ID_car"] );
                        CU.CarModel = Convert.ToString( reader["Car_Model"] );
                        CU.License = Convert.ToString( reader["License"] );
                        CU.RegDate = Convert.ToDateTime( reader["Register_date"] );

                        Cars.Add( CU );

                        CarsList.Items.Add( new TextBlock()
                            {
                                TextWrapping = TextWrapping.Wrap,
                                Text = "#" + CU.ID + " - Модель: " + CU.CarModel + "\r\nДата регистрации авто: " + CU.RegDate.ToLocalTime() + " - Номер: " + CU.License,
                                Tag = CU.ID
                            } );
                    }
                }
            }
            catch ( Exception E )
            {
                System.Windows.Forms.MessageBox.Show( E.Message );
            }
            finally
            {
                UserData.CloseConnection();
            }
        }

        private void DeleteCarBtn_Click( object sender, RoutedEventArgs e )
        {
            if ( CarsList.SelectedItem != null )
            {
                try
                {
                    UserData.OpenConnection();
                    SqlCommand command = new SqlCommand( "DeleteCar", UserData.Connection );
                    command.CommandType = System.Data.CommandType.StoredProcedure;
                    command.Parameters.AddWithValue( "CarID", ( Convert.ToInt32( ( ( TextBlock )CarsList.SelectedItem ).Tag ) ) );

                    if ( command.ExecuteNonQuery() != 2 )
                    {
                        System.Windows.Forms.MessageBox.Show( "Ошибка удаления автомобиля" );
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show( "Автомобиль успешно удален" );
                        RefreshDataBtn_Click( null, null );
                    }
                }
                catch ( Exception E )
                {
                    System.Windows.Forms.MessageBox.Show( E.Message );
                }
                finally
                {
                    UserData.CloseConnection();
                }
            }
        }

        private void AddCarBtn_Click( object sender, RoutedEventArgs e )
        {
            if ( !string.IsNullOrEmpty( licenseTB.Text ) && !string.IsNullOrEmpty( modelTB.Text ) ) 
            try
            {
                UserData.OpenConnection();
                SqlCommand command = new SqlCommand( "AddCar", UserData.Connection );
                command.CommandType = System.Data.CommandType.StoredProcedure;
                command.Parameters.AddWithValue( "CarModel", modelTB.Text );
                command.Parameters.AddWithValue( "License ", licenseTB.Text );
                command.Parameters.AddWithValue( "ID_driver", UserData.ID );
                
                if ( command.ExecuteNonQuery() != 2 )
                {
                    System.Windows.Forms.MessageBox.Show( "Ошибка добавления автомобиля" );
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show( "Автомобиль успешно добавлен" );
                    RefreshDataBtn_Click( null, null );
                }
            }
            catch ( Exception E )
            {
                System.Windows.Forms.MessageBox.Show( E.Message );
            }
            finally
            {
                UserData.CloseConnection();
            }
        }

    }
}
