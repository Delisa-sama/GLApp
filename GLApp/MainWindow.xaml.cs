using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Xml;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLApp
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        class OrderItem
        {
            public enum State
            {
                NoInfo = -2,
                Rejected = -1,
                Success = 0,
                Waiting = 1,
                InProgress = 2
            }

            public Int32 ID;
            public Int32 ClientID;
            public double Price;
            public string FromAddress;
            public string ToAddress;
            public DateTime PickTime;
            public bool[] Additions; //Детское сиденье, Кондиционер, Пьяный пассажир, Перевозка животного
            public string ClData;
            public string DrData;
            public State Status;

            public OrdResItem ConvertToOrdRes()
            {
                return new OrdResItem()
                {
                    ID = this.ID,
                    ClientID = this.ClientID,
                    Price = this.Price,
                    FromAddress = this.FromAddress,
                    ToAddress = this.ToAddress,
                    PickTime = this.PickTime,
                    Additions = this.Additions,
                    ClData = this.ClData,
                    DrData = this.DrData,
                    Status = this.Status
                };
            }
        }

        class OrdResItem : OrderItem
        {
            public DateTime BeginDate;
            public DateTime EndDate;
            public int DriverID;
            public int OpID;
        }

        class UserEntity
        {
            public string login;
            public string name;
            public string surname;
            public string phone_number;
            public int ID;
        }

        class LOG
        {
            public string Data;
            public int ID;
            public DateTime AppendDate;
            public string Target;
            public string Operation;
        }

        string NewLine = Environment.NewLine;
        
        System.Windows.Forms.DateTimePicker DTP = null;

        private OrderItem COI = null; //Current Order Item, текущий заказ

        private QuestionItem QItem = null;

        public MainWindow()
        {
            logWindow lw = new logWindow();
            lw.ShowDialog();
            InitializeComponent();
            ClientGB.Visibility = Visibility.Hidden;
            CurrentOrderGB.Visibility = Visibility.Hidden;
            DriverGB.Visibility = Visibility.Hidden;
            ManagerGB.Visibility = Visibility.Hidden;
            AdminGB.Visibility = Visibility.Hidden;
            TechSupportBtn.Visibility = Visibility.Visible;
            switch ( UserData.UserType )
            {
                case 1:
                case 2:
                    if ( !UserHaveOrder() )
                    {
                        if ( UserData.UserType == 1 )
                        {
                            ClientGB.Visibility = Visibility.Visible;
                            ClientGB.Header = CurrentUserData.Name + " " + CurrentUserData.SName;
                            DTP = new System.Windows.Forms.DateTimePicker()
                            {
                                MinDate = DateTime.Now.AddMinutes( 15 ),
                                MaxDate = DateTime.Now.AddDays( 14 ),
                                ShowUpDown = true,
                                CustomFormat = "dd.MM - HH:mm",
                                Format = System.Windows.Forms.DateTimePickerFormat.Custom
                            };
                            WFHost.Child = DTP;
                        }
                        else
                        {
                            DriverGB.Visibility = Visibility.Visible;
                            DriverGB.Header = CurrentUserData.Name + " " + CurrentUserData.SName;
                            RefreshList_Click( null, null );
                        }
                    }
                    break;
                case 3:
                    ManagerGB.Visibility = Visibility.Visible;
                    ManagerGB.Header = "Менеджер #" + UserData.ID;
                    TechSupportBtn.Visibility = Visibility.Hidden;
                    QuestionsRefreshBtn_Click( null, null );
                    break;
                case 4:
                    AdminGB.Visibility = Visibility.Visible;
                    TechSupportBtn.Visibility = Visibility.Hidden;
                    DispatcherTimer DTimer = new DispatcherTimer();
                    DTimer.Interval = IntervalDT;
                    DTimer.IsEnabled = true;
                    DTimer.Tick += new EventHandler( DTimer_Tick );
                    DTimer_Tick( null, null);
                    ClearPreviousLogs.Content = new TextBlock()
                    {
                        TextWrapping = TextWrapping.Wrap,
                        Text = "Удалить предыдущие сжатые логи"
                    };
                    break;
            }
        }

        #region Общее

        private void SendDataOnLabels()
        {
            COGBFromLabel.Content = new TextBlock()
            {
                Text = "От: " + COI.FromAddress,
                TextWrapping = TextWrapping.Wrap
            };
            COGBToLabel.Content = new TextBlock()
            {
                Text = "До: " + COI.ToAddress,
                TextWrapping = TextWrapping.Wrap
            };
            COGBPhoneNumbersLabel.Content = new TextBlock()
            {
                Text = "Водитель: " + COI.DrData + "\nКлиент: " + COI.ClData,
                TextWrapping = TextWrapping.Wrap
            };
            COGBPickTime.Content = new TextBlock()
            {
                Text = "Время подачи: " + COI.PickTime.ToShortDateString() + " " + COI.PickTime.ToShortTimeString(),
                TextWrapping = TextWrapping.Wrap
            };
            if ( COI.Additions[0] )
            {
                COGBAdditionsList.Items.Add( new Label() { Content = "Детское сиденье", Padding = new Thickness( -1 ) } );
            }
            if ( COI.Additions[1] )
            {
                COGBAdditionsList.Items.Add( new Label() { Content = "Кондиционер", Padding = new Thickness( -1 ) } );
            }
            if ( COI.Additions[2] )
            {
                COGBAdditionsList.Items.Add( new Label() { Content = "Пьяный пассажир", Padding = new Thickness( -1 ) } );
            }
            if ( COI.Additions[3] )
            {
                COGBAdditionsList.Items.Add( new Label() { Content = "Перевозка животного", Padding = new Thickness( -1 ) } );
            }
            COGBPrice.Content = COI.Price.ToString( "N2" ) + " руб.";
            string Smsg;
            switch ( COI.Status )
            {
                case OrderItem.State.InProgress:
                    Smsg = "В пути";
                    break;
                case OrderItem.State.Waiting:
                    Smsg = "Ожидание водителя";
                    break;
                default:
                    Smsg = "Статус не определен";
                    break;
            }
            COGBStatus.Content = Smsg;
        }

        private OrderItem GetOrderItem( SqlDataReader DataReader )
        {
            OrderItem.State Stat = OrderItem.State.NoInfo;
            try
            {
                switch ( ( Int16 )DataReader["Status"] )
                {
                    case 1: Stat = OrderItem.State.Waiting;
                        break;
                    case 2: Stat = OrderItem.State.InProgress;
                        break;
                    case 0: Stat = OrderItem.State.Success;
                        break;
                    case -1: Stat = OrderItem.State.Rejected;
                        break;
                }
            }
            catch { }
            OrderItem NewItem = new OrderItem();
            try
            {
                NewItem.ID = ( Int32 )DataReader["OrderID"];
            }
            catch { }
            try
            {
                NewItem.Price = Convert.ToDouble( DataReader["Price"] );
            }
            catch { }
            try
            {
                NewItem.FromAddress = ( string )DataReader["FromAddress"];
            }
            catch { }
            try
            {
                NewItem.ToAddress = ( string )DataReader["ToAddress"];
            }
            catch { }
            try
            {
                NewItem.PickTime = ( DateTime )DataReader["PickUpTime"];
            }
            catch { }
            try
            {
                NewItem.Additions = new bool[4] 
            {
                Convert.ToBoolean(DataReader["ChildSeat"]), 
                Convert.ToBoolean(DataReader["Conditioner"]), 
                Convert.ToBoolean(DataReader["DrunkPassenger"]),
                Convert.ToBoolean(DataReader["PetTransportation"]),
            };
            }
            catch { }
            try
            {
                NewItem.ClData = ( string )DataReader["CData"]; //Эксепшн ИндексАутОфРендж, в сообщении лежит имя столбца. В текущем заказе проблем не возникает
            }
            catch { }
            try
            {
                NewItem.DrData = DBNull.Value == DataReader["DData"] ? "Водитель не найден" : ( string )DataReader["DData"];
            }
            catch { }
            NewItem.Status = Stat;
            return NewItem;
        }

        private bool UserHaveOrder()
        {
            try
            {
                UserData.OpenConnection();
                SqlCommand Command = new SqlCommand( "GetCurrentOrder", UserData.Connection )
                {
                    CommandType = System.Data.CommandType.StoredProcedure
                };
                Command.Parameters.AddWithValue( "UserID", UserData.ID );
                Command.Parameters.AddWithValue( "UserType", UserData.UserType );
                SqlDataReader Read = Command.ExecuteReader();
                if ( Read.HasRows )
                {
                    CurrentOrderGB.Visibility = Visibility.Visible;
                    Read.Read();
                    COI = GetOrderItem( Read );
                    Read.Close();
                    SendDataOnLabels();
                    if ( UserData.UserType == 2 )
                    {
                        SetOrderCompleted.Visibility = System.Windows.Visibility.Hidden;
                        RejectOrder.Margin = new Thickness( 85, 335, 85, 10 );
                    }
                    else
                    {
                        SetOrderCompleted.Visibility = System.Windows.Visibility.Visible;
                        RejectOrder.Margin = new Thickness( 0, 345, 168, 0 );
                        if ( COI.Status == OrderItem.State.Waiting )
                        {
                            SetOrderCompleted.IsEnabled = false;
                        }
                        else
                        {
                            SetOrderCompleted.IsEnabled = true;
                        }
                    }
                    return true;
                }
                return false;
            }
            catch ( Exception E )
            {
                MessageBox.Show( E.Message );
                return false;
            }
            finally
            {
                UserData.CloseConnection();
            }
        }

        private void RejectOrder_Click( object sender, RoutedEventArgs e )
        {
            if ( MessageBox.Show( "Вы уверены, что хотите отменить заказ?", "Отмена заказа", MessageBoxButton.YesNo, MessageBoxImage.Question ) == MessageBoxResult.Yes )
            {
                try
                {
                    UserData.OpenConnection();
                    SqlCommand Command = new SqlCommand( "RejectOrder", UserData.Connection )
                    {
                        CommandType = System.Data.CommandType.StoredProcedure
                    };
                    Command.Parameters.AddWithValue( "OrderID", COI.ID );
                    Command.Parameters.AddWithValue( "UserType", UserData.UserType );
                    if ( Command.ExecuteNonQuery() != 2 )
                    {
                        MessageBox.Show( "Не удалось отменить заказ" );
                    }
                    else
                    {
                        CurrentOrderGB.Visibility = System.Windows.Visibility.Hidden;
                        if ( UserData.UserType == 1 )
                        {
                            ClientGB.Visibility = Visibility.Visible;
                            ClientGB.Header = CurrentUserData.Name + " " + CurrentUserData.SName;
                            DTP = new System.Windows.Forms.DateTimePicker()
                            {
                                MinDate = DateTime.Now.AddMinutes( 15 ),
                                MaxDate = DateTime.Now.AddDays( 14 ),
                                ShowUpDown = true,
                                CustomFormat = "dd.MM - HH:mm",
                                Format = System.Windows.Forms.DateTimePickerFormat.Custom
                            };
                            WFHost.Child = DTP;
                        }
                        else
                        {
                            DriverGB.Visibility = Visibility.Visible;
                            DriverGB.Header = CurrentUserData.Name + " " + CurrentUserData.SName;
                            RefreshList_Click( null, null );
                        }
                    }
                }
                catch ( Exception E )
                {
                    MessageBox.Show( E.Message );
                }
                finally
                {
                    UserData.CloseConnection();
                }
            }
        }

        private void TechSupport_Click( object sender, RoutedEventArgs e )
        {
            TechSupport TSWindow = new TechSupport();
            TSWindow.ShowDialog();
        }

        private void Window_Closed( object sender, EventArgs e )
        {
            Environment.Exit( 0 );
        }

        #endregion

        #region Работа с клиентом

        private void SwapAdresses_Click( object sender, RoutedEventArgs e )
        {
            string Temp = FromAd.Text;
            FromAd.Text = ToAd.Text;
            ToAd.Text = Temp;
        }

        double price = -100;

        private void SendOrder_Click( object sender, RoutedEventArgs e )
        {
            SendOrder.IsEnabled = false;
            if ( price < 0 )
            {
                MessageBox.Show( "Опишите заказ подробнее" );
                SendOrder.IsEnabled = true;
            }
            else
            {
                try
                {
                    UserData.OpenConnection();
                    SqlCommand command = new SqlCommand( "PlaceOrder", UserData.Connection )
                    {
                        CommandType = System.Data.CommandType.StoredProcedure
                    };
                    command.Parameters.AddWithValue( "fromAd", FromAd.Text );
                    command.Parameters.AddWithValue( "toAd", ToAd.Text );
                    command.Parameters.AddWithValue( "PickUpT", DTP.Value );
                    command.Parameters.AddWithValue( "price", price );
                    command.Parameters.AddWithValue( "Childseat", CSCB.IsChecked );
                    command.Parameters.AddWithValue( "Drunk", DrCB.IsChecked );
                    command.Parameters.AddWithValue( "Pet", AnCB.IsChecked );
                    command.Parameters.AddWithValue( "Cond", CoCB.IsChecked );
                    command.Parameters.AddWithValue( "ClientID", UserData.ID );
                    int res = command.ExecuteNonQuery();
                    if ( res == 0 )
                    {
                        throw new Exception( "Невозможно разместить заказ" );
                    }
                    else
                    {
                        MessageBox.Show( "Заказ успешно размещен" );
                        ClientGB.Visibility = System.Windows.Visibility.Hidden;
                        SendOrder.IsEnabled = true;
                        CurrentOrderGB.Visibility = System.Windows.Visibility.Visible;
                        UserHaveOrder();
                    }
                }
                catch ( Exception E )
                {
                    MessageBox.Show( E.Message );
                    SendOrder.IsEnabled = true;
                }
            }
        }

        private void CheckBox_Click( object sender, RoutedEventArgs e )
        {
            UpdatePricing();
        }

        private void CarClass_SelectionChanged( object sender, SelectionChangedEventArgs e )
        {
            UpdatePricing();
        }

        private void SetOrderCompleted_Click( object sender, RoutedEventArgs e )
        {
            UserData.OpenConnection();
            try
            {
                SqlCommand command = new SqlCommand( "SetOrderCompleted", UserData.Connection )
                    {
                        CommandType = System.Data.CommandType.StoredProcedure
                    };
                command.Parameters.AddWithValue( "OrderID", COI.ID );
                if ( command.ExecuteNonQuery() != 2 )
                {
                    MessageBox.Show( "Ошибка выполнения заказа" );
                }
                else
                {
                    MessageBox.Show( "Подтверждение выполнения получено" );
                    CurrentOrderGB.Visibility = System.Windows.Visibility.Hidden;
                    ClientGB.Visibility = Visibility.Visible;
                    ClientGB.Header = CurrentUserData.Name + " " + CurrentUserData.SName;
                    DTP = new System.Windows.Forms.DateTimePicker()
                    {
                        MinDate = DateTime.Now.AddMinutes( 15 ),
                        MaxDate = DateTime.Now.AddDays( 14 ),
                        ShowUpDown = true,
                        CustomFormat = "dd.MM - HH:mm",
                        Format = System.Windows.Forms.DateTimePickerFormat.Custom
                    };
                    WFHost.Child = DTP;
                }
            }
            catch ( Exception E )
            {
                MessageBox.Show( E.Message );
            }
            finally
            {
                UserData.CloseConnection();
            }
        }

        private void UpdatePricing()
        {
            double Price = 100;
            double Overprice = 0;
            if ( DrCB.IsChecked == true )
            {
                Overprice += 0.5;
            }
            if ( CoCB.IsChecked == true )
            {
                Overprice += 0.1;
            }
            if ( AnCB.IsChecked == true )
            {
                Overprice += 0.2;
            }
            try
            {
                int Multiplier = Convert.ToInt32( ( ( ComboBoxItem )carClass.SelectedItem ).Tag );
                Price = Multiplier * 100;
                Price += Price * Overprice;
                EstPrice.Text = Price.ToString( "N2" ) + " руб.";
            }
            catch
            { }
            if ( !string.IsNullOrEmpty( ToAd.Text ) && !string.IsNullOrEmpty( FromAd.Text ) )
            {
                price = Price;
            }
        }

        #endregion

        #region Работа с водителем

        private void RefreshList_Click( object sender, RoutedEventArgs e )
        {
            try
            {
                UserData.OpenConnection();
                SqlCommand command = new SqlCommand( "GetOrders", UserData.Connection )
                {
                    CommandType = System.Data.CommandType.StoredProcedure
                };
                OrdersList.Items.Clear();
                try
                {
                    using ( SqlDataReader Data = command.ExecuteReader() )
                    {
                        if ( Data.HasRows )
                        {
                            while ( Data.Read() )
                            {
                                COI = GetOrderItem( Data );
                                RadioButton RB = new RadioButton()
                                {
                                    Tag = COI,
                                    Content = new TextBlock()
                                    {
                                        Text = "Откуда: " + Data["FromAddress"] + "\nКуда: " + Data["ToAddress"],
                                        TextWrapping = TextWrapping.Wrap
                                    }
                                };
                                RB.Checked += new RoutedEventHandler( RB_Checked );
                                OrdersList.Items.Add( RB );
                            }
                            Data.Close();
                        }
                        else
                        {
                            OrdersList.Items.Add( new Label() { Content = "Подходящих заказов нет" } );
                        }
                    }
                }
                catch ( Exception E )
                {
                    MessageBox.Show( E.Message );
                }
            }
            catch ( Exception E )
            {
                MessageBox.Show( E.Message );
            }
            finally
            {
                UserData.CloseConnection();
            }
        }

        void RB_Checked( object sender, RoutedEventArgs e )
        {
            COI = ( OrderItem )( ( RadioButton )sender ).Tag;
            DriverPickTime.Content = new TextBlock()
            {
                Text = COI.PickTime.ToShortDateString() + " " + COI.PickTime.ToShortTimeString(),
                TextWrapping = TextWrapping.Wrap
            };
            DriverTripPrice.Content = new TextBlock()
            {
                Text = COI.Price.ToString( "N2" ),
                TextWrapping = TextWrapping.Wrap
            };
            StringBuilder Additions = new StringBuilder();
            if ( COI.Additions[0] )
            {
                Additions.AppendLine( "Детское кресло" );
            }
            if ( COI.Additions[1] )
            {
                Additions.AppendLine( "Кондиционер" );
            }
            if ( COI.Additions[2] )
            {
                Additions.AppendLine( "Пьяный пассажир" );
            }
            if ( COI.Additions[3] )
            {
                Additions.AppendLine( "Перевозка животных" );
            }
            DriverAdditional.Content = new TextBlock()
            {
                Text = Additions.ToString(),
                TextWrapping = TextWrapping.Wrap
            };
        }

        private void AcceptOrderBtn_Click( object sender, RoutedEventArgs e )
        {
            if ( COI != null )
            {
                UserData.OpenConnection();
                try
                {
                    SqlCommand command = new SqlCommand( "AcceptOrder", UserData.Connection )
                    {
                        CommandType = System.Data.CommandType.StoredProcedure
                    };
                    command.Parameters.AddWithValue( "OrderID", COI.ID );
                    command.Parameters.AddWithValue( "DriverID", UserData.ID );
                    if (command.ExecuteNonQuery() != 2)
                    {
                        MessageBox.Show( "Ошибка принятия заказа" );
                    }
                    else
                    {
                        MessageBox.Show( "Заказ принят" );
                        DriverGB.Visibility = System.Windows.Visibility.Hidden;
                        UserHaveOrder();
                    }
                }
                finally
                {
                    UserData.CloseConnection();
                }
            }
        }

        private void CarManaging_Click( object sender, RoutedEventArgs e )
        {
            CarManage CM = new CarManage();
            CM.ShowDialog();
        }

        #endregion

        #region Работа с менеджером

        private void QuestionsListBox_SelectionChanged( object sender, SelectionChangedEventArgs e )
        {
            if ( QuestionsListBox.SelectedItem != null && QuestionsListBox.SelectedItem is TextBlock )
            {
                QItem = QList.ElementAt( Convert.ToInt32( ( ( TextBlock )QuestionsListBox.SelectedItem ).Tag ) );
                QuestionText.Text = QItem.Text;
            }
        }

        private List<QuestionItem> QList = new List<QuestionItem>();

        private void QuestionsRefreshBtn_Click( object sender, RoutedEventArgs e )
        {
            UserData.OpenConnection();
            try
            {
                SqlCommand command = new SqlCommand( "GetQuestions", UserData.Connection );
                command.CommandType = System.Data.CommandType.StoredProcedure;
                SqlDataReader Reader = command.ExecuteReader();
                QuestionsListBox.Items.Clear();
                QList.Clear();
                if ( Reader.HasRows )
                {
                    while ( Reader.Read() )
                    {
                        QuestionItem QItem = new QuestionItem();
                        QItem.ID = Reader.GetInt32( 0 );
                        QItem.Header = Reader.GetString( 1 );
                        QItem.Text = Reader.GetString( 2 );
                        QItem.Answer = DBNull.Value == Reader.GetValue( 3 ) ? "Нет ответа" : Reader.GetString( 3 );
                        QItem.UserLogin = Reader["login"].ToString();
                        QList.Add( QItem );
                        QuestionsListBox.Items.Add( new TextBlock()
                        {
                            TextWrapping = TextWrapping.Wrap,
                            Tag = QList.IndexOf( QItem ),
                            Text = "Вопрос #" + QItem.ID.ToString() + ". " + QItem.Header
                        } );
                    }
                }
                else
                {
                    QuestionsListBox.Items.Add( new Label() { Content = "Не найдено ни одного вопроса" } );
                }
            }
            catch ( Exception E )
            {
                MessageBox.Show( E.Message );
            }
            finally
            {
                UserData.CloseConnection();
            }
        }

        private void SendAnswerBtn_Click( object sender, RoutedEventArgs e )
        {
            if ( QItem != null && !string.IsNullOrEmpty( ManagerAnswerTextBox.Text ) )
            {
                try
                {
                    UserData.OpenConnection();
                    SqlCommand command = new SqlCommand( "AnswerQuestion", UserData.Connection );
                    command.Parameters.AddWithValue( "answer", ManagerAnswerTextBox.Text );
                    command.Parameters.AddWithValue( "managerID", UserData.ID );
                    command.Parameters.AddWithValue( "QID", QItem.ID );//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!Параметр не передается серверу, но находится в команде????
                    if ( command.ExecuteNonQuery() != 2 )
                    {
                        MessageBox.Show( "Ошибка добавления ответа" );
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show( "Ответ размещен" );
                    }

                }
                catch ( Exception E )
                {
                    MessageBox.Show( E.Message );
                }
                finally
                {
                    UserData.CloseConnection();
                }
            }
        }

        #endregion

        #region Работа с администратором

        internal int UserTypeSelected = 1;

        private List<OrdResItem> ORList = new List<OrdResItem>();

        internal DispatcherTimer DTimer = new DispatcherTimer();

        private TimeSpan IntervalDT = new TimeSpan( 0, 0, 15 );

        void DTimer_Tick( object sender, EventArgs e )
        {
            try
            {
                UserData.OpenConnection();
                SqlCommand command = new SqlCommand( "GetAllOrders", UserData.Connection );
                SqlDataReader reader = command.ExecuteReader();
                if ( reader.HasRows )
                {
                    ORList.Clear();
                    AdminOrdersList.Items.Clear();
                    while ( reader.Read() )
                    {
                        OrdResItem ORI = GetOrderItem( reader ).ConvertToOrdRes();
                        ORI.BeginDate = DBNull.Value == reader["Datetime_beg"] ? new DateTime( 0 ) : reader.GetDateTime( 2 );
                        ORI.EndDate = DBNull.Value == reader["Datetime_end"] ? new DateTime( 0 ) : reader.GetDateTime( 3 );
                        ORI.DriverID = DBNull.Value == reader["ID_driver"] ? -1 : ( Int32 )reader["ID_driver"];
                        ORI.OpID = ( Int32 )reader["ID_operation"];
                        ORList.Add( ORI );

                        StringBuilder OrdResData = new StringBuilder();
                        OrdResData.AppendLine( "Откуда: " + ORI.FromAddress );
                        OrdResData.AppendLine( "Куда: " + ORI.ToAddress );
                        OrdResData.AppendLine( "Стоимость: " + ORI.Price + " руб.\tID заказа: " + ORI.ID );
                        if ( ORI.ClientID != 0 )
                        {
                            OrdResData.AppendLine( "ID клиента: " + ORI.ClientID );
                        }
                        if ( ORI.DriverID != 0 )
                        {
                            OrdResData.AppendLine( "ID водителя: " + ORI.ClientID );
                            OrdResData.AppendLine( "Заказ принят: " + ORI.BeginDate.ToLocalTime() );
                            if ( ORI.Status == OrderItem.State.Success || ORI.Status == OrderItem.State.Rejected )
                            {
                                OrdResData.AppendLine( "Заказ завершен: " + ORI.EndDate.ToLocalTime() );
                            }
                        }
                        OrdResData.AppendLine( "Подать машину к: " + ORI.PickTime + "\tСтатус заказа: " + ORI.Status );
                        if ( ORI.Additions[0] )
                        {
                            OrdResData.AppendLine( "Требуется детское кресло" );
                        }
                        if ( ORI.Additions[1] )
                        {
                            OrdResData.AppendLine( "Требуется кондиционер" );
                        }
                        if ( ORI.Additions[2] )
                        {
                            OrdResData.AppendLine( "Перевозка пассажира в состоянии алкогольного опьянения" );
                        }
                        if ( ORI.Additions[3] )
                        {
                            OrdResData.AppendLine( "Перевозка животных" );
                        }
                        OrdResData.AppendLine( "Код операции: " + ORI.OpID );

                        AdminOrdersList.Items.Add( new TextBlock()
                            {
                                Text = OrdResData.ToString(),
                                TextWrapping = TextWrapping.Wrap,
                                Tag = ORList.IndexOf( ORI )
                            } );
                    }
                }
                else
                {
                    AdminOrdersList.Items.Clear();
                    AdminOrdersList.Items.Add( new TextBlock()
                    {
                        TextWrapping = TextWrapping.Wrap,
                        Text = "Нет активных заказов"
                    } );
                }
            }
            catch ( Exception E )
            {
                MessageBox.Show( E.Message );
            }
            finally
            {
                UserData.CloseConnection();
            }

        }

        private void RefreshOrdResList_Click( object sender, RoutedEventArgs e )
        {
            DTimer_Tick( null, null );
            DTimer.Stop();
            DTimer.Start();
        }
        
        //Не работает ни изменение интервала, ни остановка таймера
        private void UpdateTimerInterval_Click( object sender, RoutedEventArgs e )
        {
            if ( sender != null )
            {
                DTimer.Stop(); 
                int Seconds = Convert.ToInt32(( ( MenuItem )sender ).Tag);
                IntervalDT = new TimeSpan( 0, 0, Seconds );
                DTimer.Interval = IntervalDT;
                DTimer.Start();
            }
        }

        private void DeleteSelectedOrder_Click( object sender, RoutedEventArgs e )
        {
            if ( AdminOrdersList.SelectedItem != null )
            {
                if ( ConfirmDeleteCB.IsChecked == false || MessageBox.Show( "Вы уверены, что хотите удалить запись?", "", MessageBoxButton.YesNo, MessageBoxImage.Question ) == MessageBoxResult.Yes )
                {
                    try
                    {
                        UserData.OpenConnection();
                        SqlCommand command = new SqlCommand( "RemoveOrder", UserData.Connection );
                        command.CommandType = System.Data.CommandType.StoredProcedure;
                        OrdResItem ORI = ORList.ElementAt( Convert.ToInt32( ( ( TextBlock )AdminOrdersList.SelectedItem ).Tag ) );
                        command.Parameters.AddWithValue( "OrderID", ORI.ID );
                        command.Parameters.AddWithValue( "OperationID", ORI.OpID );
                        if ( command.ExecuteNonQuery() != 2 )
                        {
                            MessageBox.Show( "Ошибка удаления записи" );
                        }
                        else
                        {
                            DTimer_Tick( null, null );
                        }
                    }
                    catch ( Exception E )
                    {
                        MessageBox.Show( E.Message );
                    }
                    finally
                    {
                        UserData.CloseConnection();
                    }
                }
            }
        }

        private void AdminOrdersList_SelectionChanged( object sender, SelectionChangedEventArgs e )
        {
            if ( AdminOrdersList.SelectedItem != null )
            {
                DelOrdResBtn.IsEnabled = true;
                EditUser.IsEnabled = true;
            }
            else
            {
                DelOrdResBtn.IsEnabled = false;
                EditUser.IsEnabled = false;
            }
        }

        private void AdminTabControl_SelectionChanged( object sender, SelectionChangedEventArgs e )
        {
            if ( UserData.UserType == 4 && AdminTabControl.SelectedItem != null )
                if ( Convert.ToInt32( ( ( TabItem )AdminTabControl.SelectedItem ).Tag ) != 2 )
                {
                    DTimer.IsEnabled = false;
                }
                else
                {
                    if ( AutoUpdateCB.IsChecked == true )
                    {
                        DTimer.IsEnabled = true; ;
                    }
                }
        }

        private void ComboBox_SelectionChanged( object sender, SelectionChangedEventArgs e )
        {
            if ( UserTypeSelector.SelectedItem != null )
            {
                UserTypeSelected = Convert.ToInt32( ( ( ComboBoxItem )UserTypeSelector.SelectedItem ).Tag ) == 0 ? 1 : Convert.ToInt32( ( ( ComboBoxItem )UserTypeSelector.SelectedItem ).Tag );
                if ( AdminUsersList != null )
                {
                    RefreshUsersList_Click( null, null );
                }
                if ( AddInfoGB != null )
                switch ( UserTypeSelected )
                {
                    case 1:
                    case 2:
                        AddInfoGB.Visibility = System.Windows.Visibility.Visible;
                        break;
                    default:
                        AddInfoGB.Visibility = System.Windows.Visibility.Hidden;
                        break;
                }
            }
        }

        private void AddUserBtn_Click( object sender, RoutedEventArgs e )
        {
            if ( !string.IsNullOrEmpty( loginTB.Text ) && !string.IsNullOrEmpty( passwd.Password ) )
            {
                try
                {
                    UserData.OpenConnection();
                    SqlCommand command = new SqlCommand( "AddUser", UserData.Connection );
                    command.CommandType = System.Data.CommandType.StoredProcedure;
                    if ( string.IsNullOrEmpty( nameTB.Text ) )
                    {
                        nameTB.Text = "UnAssignedName";
                    }
                    if ( string.IsNullOrEmpty( surnameTB.Text ) )
                    {
                        surnameTB.Text = "UnassignedSurname";
                    }
                    if ( string.IsNullOrEmpty( phoneTB.Text ) )
                    {
                        phoneTB.Text = "88005553535";
                    }
                    command.Parameters.AddWithValue( "login", loginTB.Text );
                    command.Parameters.AddWithValue( "pwd", passwd.Password );
                    command.Parameters.AddWithValue( "uname", nameTB.Text );
                    command.Parameters.AddWithValue( "surname", surnameTB.Text );
                    command.Parameters.AddWithValue( "utype", UserTypeSelected );
                    command.Parameters.AddWithValue( "phone_numb", phoneTB.Text );

                    if ( command.ExecuteNonQuery() != 2 )
                    {
                        throw new System.Exception( "Произошла ошибка, пользователь не был добавлен" );
                    }
                    else
                    {
                        MessageBox.Show( "Пользователь успешно добавлен" );
                    }
                }
                catch ( Exception Exc )
                {
                    MessageBox.Show( Exc.Message );
                }
                finally
                {
                    UserData.CloseConnection();
                }
            }
            else
            {
                MessageBox.Show( "Введите логин и пароль" );
            }
        }

        private void EditUser_Click( object sender, RoutedEventArgs e )
        {
            if ( AdminUsersList.SelectedItem != null && !string.IsNullOrEmpty( loginTB.Text ) )
            {
                UserEntity UE = UsersList.ElementAt( Convert.ToInt32( ( TextBlock )AdminUsersList.Tag ) );
                try
                {
                    UserData.OpenConnection();
                    SqlCommand command = new SqlCommand( "UpdateUser", UserData.Connection );
                    command.CommandType = System.Data.CommandType.StoredProcedure;
                    command.Parameters.AddWithValue( "ID", UE.ID );
                    command.Parameters.AddWithValue( "type", UserTypeSelected );
                    command.Parameters.AddWithValue( "login", loginTB.Text );
                    if ( string.IsNullOrEmpty( passwd.Password ) )
                    {
                        command.Parameters.AddWithValue( "pass", "0" );
                    }
                    else
                    {
                        command.Parameters.AddWithValue( "pass", passwd.Password );
                    }
                    command.Parameters.AddWithValue( "name", nameTB.Text );
                    command.Parameters.AddWithValue( "surname", surnameTB.Text );
                    command.Parameters.AddWithValue( "phone_number", phoneTB.Text );

                    if ( command.ExecuteNonQuery() != 2 )
                    {
                        System.Windows.Forms.MessageBox.Show( "Запись не была изменена" );
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show( "Данные о пользователе успешно обновлены" );
                    }
                }
                catch ( Exception E )
                {
                    Console.WriteLine( E.Message );
                    System.Windows.Forms.MessageBox.Show( "Ошибка изменения пользователя" );
                }
                finally
                {
                    UserData.CloseConnection();
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show( "Определите логин" );
            }
        }

        private void DeleteUser_Click( object sender, RoutedEventArgs e )
        {
            if ( AdminUsersList.SelectedItem != null )
            {
                UserEntity UE = UsersList.ElementAt( Convert.ToInt32((TextBlock)AdminUsersList.Tag) );
                try
                {
                    UserData.OpenConnection();
                    SqlCommand command = new SqlCommand( "DeleteUser", UserData.Connection );
                    command.CommandType = System.Data.CommandType.StoredProcedure;

                    command.Parameters.AddWithValue( "login", UE.login);
                    command.Parameters.AddWithValue( "type", UserTypeSelected );

                    int exq = command.ExecuteNonQuery();

                    if ( exq != 1 )
                    {
                        throw new System.Exception( "Произошла ошибка, пользователь не был удален" );
                    }
                    else
                    {
                        MessageBox.Show( "Пользователь был успешно удален" );
                    }
                }
                catch ( Exception Exc )
                {
                    MessageBox.Show( Exc.Message );
                }
                finally
                {
                    UserData.CloseConnection();
                }
            }
        }
        
        List<UserEntity> UsersList = new List<UserEntity>();

        private void RefreshUsersList_Click( object sender, RoutedEventArgs e )
        {
            try
            {
                UserData.OpenConnection();
                SqlCommand command = new SqlCommand( "GetUsers", UserData.Connection );
                command.CommandType = System.Data.CommandType.StoredProcedure;
                command.Parameters.AddWithValue( "UserType", UserTypeSelected );
                SqlDataReader reader = command.ExecuteReader();
                AdminUsersList.Items.Clear();
                UsersList.Clear();
                if ( reader.HasRows )
                {
                    while ( reader.Read() )
                    {
                        UserEntity UE = new UserEntity()
                        {
                            ID = Convert.ToInt32( reader["ID"] ),
                            login = Convert.ToString( reader["login"] )
                        };

                        if ( UserTypeSelected < 3 )
                        {
                            UE.name = Convert.ToString( reader["name"] );
                            UE.surname = Convert.ToString( reader["surname"] );
                            UE.phone_number = Convert.ToString( reader["phone_number"] );
                        }

                        UsersList.Add( UE );
                        AdminUsersList.Items.Add( new TextBlock()
                        {
                            TextWrapping = TextWrapping.Wrap,
                            Text = "ID:  " + UE.ID + "\t Логин:  " + UE.login,
                            Tag = UsersList.IndexOf( UE )
                        } );
                    }
                }
                else
                {
                    AdminUsersList.Items.Add( new TextBlock()
                        {
                            TextWrapping = TextWrapping.Wrap,
                            Text = "Не найдено ни одного пользователя"
                        } );
                }
            }
            catch ( Exception E )
            {
                MessageBox.Show( E.Message );
            }
            finally
            {
                UserData.CloseConnection();
            }
        }

        private void AdminUsersList_Loaded( object sender, RoutedEventArgs e )
        {
            if ( UserData.UserType == 4 )
            {
                RefreshUsersList_Click( null, null );
            }
        }

        private UserEntity CurrentUE = null;

        internal void AdminUsersList_SelectionChanged( object sender, SelectionChangedEventArgs e )
        {
            if ( AdminUsersList.SelectedItem != null && UsersList.Count > 0 )
            {
                UserEntity UE = UsersList.ElementAt( Convert.ToInt32( ( ( TextBlock )AdminUsersList.SelectedItem ).Tag ) );
                if ( UE != null )
                {
                    loginTB.Text = UE.login;
                    nameTB.Text = UE.name;
                    surnameTB.Text = UE.surname;
                    phoneTB.Text = UE.phone_number;
                    CurrentUE = UE;
                }
            }
        }

        private void AutoUpdateCB_Checked( object sender, RoutedEventArgs e )
        {
            DTimer.Start();
        }

        private void AutoUpdateCB_Unchecked( object sender, RoutedEventArgs e )
        {
            DTimer.Stop();
        }
        
        private List<LOG> LList = new List<LOG>();

        private void LogRefreshBtn_Click( object sender, RoutedEventArgs e )
        {
            try
            {
                UserData.OpenConnection();
                if ( !( sender is int ) || ( sender is int && Convert.ToInt32( sender ) != -1 ) )
                {
                    LogsList.Items.Clear();
                }
                LList.Clear();
                SqlCommand command = new SqlCommand( "GetLogs", UserData.Connection );
                command.CommandType = System.Data.CommandType.StoredProcedure;
                if ( sender is int )
                {
                    command.Parameters.AddWithValue("RecNumber", Convert.ToInt32( sender ));
                }
                else
                {
                    command.Parameters.AddWithValue("RecNumber", 1000 );
                }
                SqlDataReader reader = command.ExecuteReader();
                if ( reader.HasRows )
                {
                    while ( reader.Read() )
                    {
                        LOG LU = new LOG() 
                        {
                            ID = Convert.ToInt32( reader["ID"] ),
                            Operation = Convert.ToString( reader["Operation"] ),
                            Data = Convert.ToString( reader["Data"] ),
                            Target = Convert.ToString( reader["TargetTable"] ),
                            AppendDate = Convert.ToDateTime( reader["AppendDate"] ),
                        };
                        LList.Add( LU );

                        if ( !(sender is int) || (sender is int && Convert.ToInt32( sender ) != -1) )
                        {
                            LogsList.Items.Add( new TextBlock()
                            {
                                Text = "#" + LU.ID + " операция " + LU.Operation + " в " + LU.Target + NewLine + "Описание даных: " + LU.Data + NewLine + "Время операции: " + LU.AppendDate,
                                TextWrapping = TextWrapping.Wrap
                            } );
                        }
                    }
                }
                else
                {
                    LogsList.Items.Add( new TextBlock()
                    {
                        Text = "Нет записей",
                        TextWrapping = TextWrapping.Wrap
                    } );
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
            RecCountLbl.Content = new TextBlock()
            {
                Text = "Найдено записей: " + LList.Count,
                TextWrapping = TextWrapping.Wrap
            };
        }

        private void RecordsNumber_TextChanged( object sender, TextChangedEventArgs e )
        {
            int RecCount;
            if ( int.TryParse( RecordsNumber.Text, out RecCount ) )
            {
                LogRefreshBtn_Click( RecCount, null );
            }
        }

        private void CompressLogsBtn_Click( object sender, RoutedEventArgs e )
        {
            LogRefreshBtn_Click( -1, null );

            int CompressedLogs = 0;

            UserData.OpenConnection();
            if ( ClearPreviousLogs.IsChecked == true )
            {
                SqlCommand command = new SqlCommand( "ClearCompressedLogs", UserData.Connection );
                command.CommandType = System.Data.CommandType.StoredProcedure;

                if ( command.ExecuteNonQuery() != 0 )
                {
                    System.Windows.Forms.MessageBox.Show( "Предыдущие логи удалены" );
                }
            }
            try
            {
                SqlCommand command = new SqlCommand( "CompressLogs", UserData.Connection );
                command.CommandType = System.Data.CommandType.StoredProcedure;

                if ( ( CompressedLogs = command.ExecuteNonQuery() ) == 0 )
                {
                    Console.WriteLine( "Ошибка при добавлении записи" );
                }
            }
            catch ( Exception E )
            {
                Console.WriteLine( E.Message );
            }
            finally
            {
                UserData.CloseConnection();
            }
            System.Windows.Forms.MessageBox.Show( "Обработано записей: " + LList.Count + ". Сжатых записей: " + CompressedLogs, "Результаты сжатия", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information );
        }

        private void WriteXMLBtn_Click( object sender, RoutedEventArgs e )
        {
            
            System.Windows.Forms.FolderBrowserDialog folderDialog = new System.Windows.Forms.FolderBrowserDialog()
            {
                SelectedPath = @"C:\"
            };
            if ( folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK && !string.IsNullOrEmpty( folderDialog.SelectedPath ) )
            {
                XmlDocument xDoc = new XmlDocument();
                XmlElement RootNode = xDoc.CreateElement( "database" );
                
                #region XML Users

                for ( int i = 1; i <= 4; i++ )
                {
                    List<UserEntity> TypedUserList = new List<UserEntity>();

                    string UserType = "";
                    switch ( i )
                    {
                        case 1:
                            UserType = "Clients";
                            break;
                        case 2:
                            UserType = "Drivers";
                            break;
                        case 3:
                            UserType = "Managers";
                            break;
                        case 4:
                            UserType = "Administrators";
                            break;
                    }

                    XmlElement ClassNode = xDoc.CreateElement( UserType, "Users" );

                    try
                    {
                        UserData.OpenConnection();
                        SqlCommand command = new SqlCommand( "GetUsers", UserData.Connection );
                        command.CommandType = System.Data.CommandType.StoredProcedure;
                        command.Parameters.AddWithValue( "UserType", i );
                        SqlDataReader reader = command.ExecuteReader();
                        if ( reader.HasRows )
                        {
                            while ( reader.Read() )
                            {
                                UserEntity UE = new UserEntity()
                                {
                                    ID = Convert.ToInt32( reader["ID"] ),
                                    login = Convert.ToString( reader["login"] )
                                };

                                if ( i < 3 )
                                {
                                    UE.name = Convert.ToString( reader["name"] );
                                    UE.surname = Convert.ToString( reader["surname"] );
                                    UE.phone_number = Convert.ToString( reader["phone_number"] );
                                }

                                TypedUserList.Add( UE );
                            }
                        }
                        reader.Close();
                    }
                    finally
                    {
                        UserData.OpenConnection();
                    }


                    foreach ( UserEntity UE in TypedUserList )
                    {
                        XmlElement UserNode = xDoc.CreateElement( "user" );

                        XmlNode Login = xDoc.CreateElement( "login" );
                        XmlAttribute ID = xDoc.CreateAttribute( "listID" );

                        Login.InnerText = UE.login;
                        ID.InnerText = UE.ID.ToString();

                        UserNode.AppendChild( Login );
                        UserNode.Attributes.Append( ID );

                        if ( i < 3 )
                        {
                            XmlNode Name = xDoc.CreateElement( "name" );
                            XmlNode Surname = xDoc.CreateElement( "surname" );
                            XmlNode Phone = xDoc.CreateElement( "phone_number" );

                            Phone.InnerText = UE.phone_number;
                            Name.InnerText = UE.name;
                            Surname.InnerText = UE.surname;

                            UserNode.AppendChild( Phone );
                            UserNode.AppendChild( Surname );
                            UserNode.AppendChild( Name );
                        }

                        ClassNode.AppendChild( UserNode );
                    }
                    RootNode.AppendChild( ClassNode );
                }

                #endregion

                #region XML Orders & Reservs

                {
                    XmlElement ClassNode = xDoc.CreateElement( "Orders", "Data" );

                    List<OrdResItem> XMLORList = new List<OrdResItem>();

                    try
                    {
                        UserData.OpenConnection();
                        SqlCommand command = new SqlCommand( "GetAllOrders", UserData.Connection );
                        command.CommandType = System.Data.CommandType.StoredProcedure;
                        SqlDataReader reader = command.ExecuteReader();
                        if ( reader.HasRows )
                        {
                            while ( reader.Read() )
                            {
                                OrdResItem ORI = GetOrderItem( reader ).ConvertToOrdRes();
                                ORI.BeginDate = DBNull.Value == reader["Datetime_beg"] ? new DateTime( 0 ) : reader.GetDateTime( 2 );
                                ORI.EndDate = DBNull.Value == reader["Datetime_end"] ? new DateTime( 0 ) : reader.GetDateTime( 3 );
                                ORI.DriverID = DBNull.Value == reader["ID_driver"] ? -1 : ( Int32 )reader["ID_driver"];
                                ORI.OpID = ( Int32 )reader["ID_operation"];
                                XMLORList.Add( ORI );
                            }
                        }
                        reader.Close();
                    }
                    finally
                    {
                        UserData.CloseConnection();
                    }


                    foreach ( OrdResItem ORI in XMLORList )
                    {
                        XmlElement UserNode = xDoc.CreateElement( "order" );

                        XmlNode Status      = xDoc.CreateElement( "status" );
                        XmlNode DT_B        = xDoc.CreateElement( "Begin_Date" );
                        XmlNode DT_E        = xDoc.CreateElement( "End_Date" );
                        XmlNode FromAdd     = xDoc.CreateElement( "fromAddress" );
                        XmlNode ToAdd       = xDoc.CreateElement( "toAddress" );
                        XmlNode ID_cl       = xDoc.CreateElement( "ClientID" );
                        XmlNode ID_dr       = xDoc.CreateElement( "DriverID" );
                        XmlNode ChildSeat   = xDoc.CreateElement( "Childseat" );
                        XmlNode Conditioner = xDoc.CreateElement( "Conditioner" );
                        XmlNode DrunkPass   = xDoc.CreateElement( "DrunkPassenger" );
                        XmlNode PetTransp   = xDoc.CreateElement( "PetTransportation" );
                        XmlNode Price       = xDoc.CreateElement( "Price" );
                        XmlNode PickUpTime  = xDoc.CreateElement( "PickUpTime" );
                        XmlAttribute OID    = xDoc.CreateAttribute( "orderID" );
                        XmlAttribute RID    = xDoc.CreateAttribute( "ReserveID" );
                        
                        Status      .InnerText      = ORI.Status.ToString();
                        DT_B        .InnerText      = ORI.BeginDate.ToString();
                        DT_E        .InnerText      = ORI.EndDate.ToString();
                        FromAdd     .InnerText      = ORI.FromAddress.ToString();
                        ToAdd       .InnerText      = ORI.ToAddress.ToString();
                        ID_cl       .InnerText      = ORI.ClientID.ToString();
                        ID_dr       .InnerText      = ORI.DriverID.ToString();
                        ChildSeat   .InnerText      = ORI.Additions[0].ToString();
                        Conditioner .InnerText      = ORI.Additions[1].ToString();
                        DrunkPass   .InnerText      = ORI.Additions[2].ToString();
                        PetTransp   .InnerText      = ORI.Additions[3].ToString();
                        Price       .InnerText      = ORI.Price.ToString( "N2" );
                        PickUpTime  .InnerText      = ORI.PickTime.ToString();
                        
                        OID.InnerText = ORI.ID.ToString();
                        RID.InnerText = ORI.OpID.ToString();

                        UserNode.AppendChild( Status );
                        UserNode.AppendChild( DT_B );
                        UserNode.AppendChild( DT_E );
                        UserNode.AppendChild( FromAdd );
                        UserNode.AppendChild( ToAdd );
                        UserNode.AppendChild( ID_cl );
                        UserNode.AppendChild( ID_dr );
                        UserNode.AppendChild( ChildSeat );
                        UserNode.AppendChild( Conditioner );
                        UserNode.AppendChild( DrunkPass );
                        UserNode.AppendChild( PetTransp );
                        UserNode.AppendChild( Price );
                        UserNode.AppendChild( PickUpTime );
                        UserNode.Attributes.Append( OID );
                        UserNode.Attributes.Append( RID );

                        ClassNode.AppendChild( UserNode );
                    }
                    RootNode.AppendChild( ClassNode );
                }

                #endregion

                #region XML Cars

                {
                    XmlElement ClassNode = xDoc.CreateElement( "Cars", "Data" );
                    try
                    {
                        UserData.OpenConnection();
                        SqlCommand command = new SqlCommand( "GetCarsList", UserData.Connection );
                        command.CommandType = System.Data.CommandType.StoredProcedure;
                        command.Parameters.AddWithValue( "ID_driver", -1 );
                        SqlDataReader reader = command.ExecuteReader();
                        if ( reader.HasRows )
                        {
                            while ( reader.Read() )
                            {

                                XmlElement CarNode = xDoc.CreateElement( "car" );

                                XmlNode Model = xDoc.CreateElement( "model" );
                                XmlNode License = xDoc.CreateElement( "license" );
                                XmlNode RegDate = xDoc.CreateElement( "register_date" );
                                XmlNode IDDriver = xDoc.CreateElement( "driverID" );
                                XmlAttribute ID = xDoc.CreateAttribute( "ID" );

                                Model.InnerText = reader["Car_Model"].ToString();
                                License.InnerText = reader["License"].ToString();
                                RegDate.InnerText = reader["Register_Date"].ToString();
                                IDDriver.InnerText = reader["ID_driver"].ToString();
                                ID.InnerText = reader["ID_car"].ToString();

                                CarNode.AppendChild( Model );
                                CarNode.AppendChild( License );
                                CarNode.AppendChild( IDDriver );
                                CarNode.AppendChild( RegDate );
                                CarNode.Attributes.Append( ID );

                                ClassNode.AppendChild( CarNode );
                            }
                        }
                        reader.Close();
                        RootNode.AppendChild( ClassNode );
                    }
                    finally
                    {
                        UserData.OpenConnection();
                    }
                }


                #endregion

                #region XML Questions

                {
                    XmlElement ClassNode = xDoc.CreateElement( "Questions", "Data" );
                    try
                    {
                        UserData.OpenConnection();
                        SqlCommand command = new SqlCommand( "GetQuestions", UserData.Connection );
                        command.CommandType = System.Data.CommandType.StoredProcedure;
                        SqlDataReader reader = command.ExecuteReader();
                        if ( reader.HasRows )
                        {
                            while ( reader.Read() )
                            {

                                XmlElement CarNode = xDoc.CreateElement( "car" );

                                XmlNode Header = xDoc.CreateElement( "header" );
                                XmlNode QText = xDoc.CreateElement( "QText" );
                                XmlNode Answer = xDoc.CreateElement( "answer" );
                                XmlNode UserLogin = xDoc.CreateElement( "UserLogin" );
                                XmlNode ManagerID = xDoc.CreateElement( "ManagerID" );
                                XmlAttribute ID = xDoc.CreateAttribute( "ID" );

                                Header.InnerText = reader["Header"].ToString();
                                QText.InnerText = reader["QText"].ToString();
                                Answer.InnerText = reader["answer"].ToString();
                                UserLogin.InnerText = reader["login"].ToString();
                                ManagerID.InnerText = reader["ManagerID"].ToString();
                                ID.InnerText = reader["ID_question"].ToString();

                                CarNode.AppendChild( Header );
                                CarNode.AppendChild( QText );
                                CarNode.AppendChild( Answer );
                                CarNode.AppendChild( UserLogin );
                                CarNode.AppendChild( ManagerID );
                                CarNode.Attributes.Append( ID );

                                ClassNode.AppendChild( CarNode );
                            }
                        }
                        reader.Close();
                        RootNode.AppendChild( ClassNode );
                    }
                    finally
                    {
                        UserData.OpenConnection();
                    }
                }

                #endregion

                xDoc.AppendChild( RootNode );
                try
                {
                    xDoc.Save( folderDialog.SelectedPath + @"\GLPRDB_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + ".xml" );
                }
                catch ( UnauthorizedAccessException UAE )
                {
                    System.Windows.Forms.MessageBox.Show( "Недостаточно прав для записи файла" );
                    Console.WriteLine( UAE.Message );
                }
            }
        }

        private void WriteXLSX_Click( object sender, RoutedEventArgs e )
        {
            System.Windows.Forms.FolderBrowserDialog folderDialog = new System.Windows.Forms.FolderBrowserDialog()
            {
                SelectedPath = @"C:\"
            };
            if ( folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK && !string.IsNullOrEmpty( folderDialog.SelectedPath ) )
            {
                Excel.Application exApp = new Excel.Application();

                exApp.Visible = false;

                Excel.Workbook exBook = exApp.Workbooks.Add();

                Excel.Worksheet exSheet = exBook.ActiveSheet;

                #region Users

                {
                    exSheet.Name = "Users";
                    int EditableRow = 1;

                    for ( int i = 1; i <= 4; i++ )
                    {
                        bool isOddLine = true;

                        List<UserEntity> TypedUserList = new List<UserEntity>();

                        string UserType = "";
                        switch ( i )
                        {
                            case 1:
                                UserType = "Clients";
                                break;
                            case 2:
                                UserType = "Drivers";
                                break;
                            case 3:
                                UserType = "Managers";
                                break;
                            case 4:
                                UserType = "Administrators";
                                break;
                        }

                        Excel.Range NameRange = exSheet.Range[exSheet.Cells[EditableRow, 1], exSheet.Cells[EditableRow, 5]];
                        NameRange.Merge();
                        NameRange.Font.Italic = true;
                        NameRange.Font.Bold = true;
                        NameRange.Value = UserType;
                        NameRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                        NameRange.Interior.Color = ColorTranslator.ToOle( Color.LightGray );

                        try
                        {
                            UserData.OpenConnection();
                            SqlCommand command = new SqlCommand( "GetUsers", UserData.Connection );
                            command.CommandType = System.Data.CommandType.StoredProcedure;
                            command.Parameters.AddWithValue( "UserType", i );
                            SqlDataReader reader = command.ExecuteReader();
                            if ( reader.HasRows )
                            {
                                while ( reader.Read() )
                                {
                                    UserEntity UE = new UserEntity()
                                    {
                                        ID = Convert.ToInt32( reader["ID"] ),
                                        login = Convert.ToString( reader["login"] )
                                    };

                                    if ( i < 3 )
                                    {
                                        UE.name = Convert.ToString( reader["name"] );
                                        UE.surname = Convert.ToString( reader["surname"] );
                                        UE.phone_number = Convert.ToString( reader["phone_number"] );
                                    }

                                    TypedUserList.Add( UE );
                                }
                            }
                            reader.Close();
                        }
                        finally
                        {
                            UserData.OpenConnection();
                        }

                        foreach ( UserEntity UE in TypedUserList )
                        {
                            EditableRow++;
                            if ( isOddLine )
                            {
                                exSheet.Cells[EditableRow, 1].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                exSheet.Cells[EditableRow, 2].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                exSheet.Cells[EditableRow, 3].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                exSheet.Cells[EditableRow, 4].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                exSheet.Cells[EditableRow, 5].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                isOddLine = false;
                            }
                            else
                            {
                                isOddLine = true;
                            }
                            exSheet.Cells[EditableRow, 1] = UE.ID.ToString();
                            exSheet.Cells[EditableRow, 2] = UE.login;
                            exSheet.Cells[EditableRow, 3] = UE.name;
                            exSheet.Cells[EditableRow, 4] = UE.surname;
                            exSheet.Cells[EditableRow, 5] = UE.phone_number;
                        }

                        Excel.Range Underline = exSheet.Range[exSheet.Cells[EditableRow, 1], exSheet.Cells[EditableRow, 5]];
                        Underline.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                        EditableRow += 3;

                        exSheet.Columns.AutoFit();
                    }

                    Excel.Range RightBorder = exSheet.Range[exSheet.Cells[1, 5], exSheet.Cells[EditableRow - 3, 5]];
                    RightBorder.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

                }
                #endregion

                #region Cars
                {
                    int EditableRow = 1;

                    bool isOddLine = true;

                    exSheet = exBook.Worksheets.Add();
                    exSheet.Activate();
                    exSheet.Name = "Cars";

                    Excel.Range NameRange = exSheet.Range[exSheet.Cells[EditableRow, 1], exSheet.Cells[EditableRow, 5]];
                    NameRange.Font.Italic = true;
                    NameRange.Font.Bold = true;
                    NameRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    NameRange.Interior.Color = ColorTranslator.ToOle( Color.LightGray );
                    exSheet.Cells[EditableRow, 1] = "Модель авто";
                    exSheet.Cells[EditableRow, 2] = "Номерной знак";
                    exSheet.Cells[EditableRow, 3] = "Дата регистрации";
                    exSheet.Cells[EditableRow, 4] = "Код водителя";
                    exSheet.Cells[EditableRow, 5] = "Код авто";

                    try
                    {
                        UserData.OpenConnection();
                        SqlCommand command = new SqlCommand( "GetCarsList", UserData.Connection );
                        command.CommandType = System.Data.CommandType.StoredProcedure;
                        command.Parameters.AddWithValue( "ID_driver", -1 );
                        SqlDataReader reader = command.ExecuteReader();
                        if ( reader.HasRows )
                        {
                            while ( reader.Read() )
                            {
                                EditableRow++;
                                if ( isOddLine )
                                {
                                    exSheet.Cells[EditableRow, 1].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                    exSheet.Cells[EditableRow, 2].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                    exSheet.Cells[EditableRow, 3].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                    exSheet.Cells[EditableRow, 4].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                    exSheet.Cells[EditableRow, 5].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                    isOddLine = false;
                                }
                                else
                                {
                                    isOddLine = true;
                                }
                                exSheet.Cells[EditableRow, 1] = reader["Car_Model"].ToString();
                                exSheet.Cells[EditableRow, 2] = reader["License"].ToString();
                                exSheet.Cells[EditableRow, 3] = reader["Register_Date"].ToString();
                                exSheet.Cells[EditableRow, 4] = reader["ID_driver"].ToString();
                                exSheet.Cells[EditableRow, 5] = reader["ID_car"].ToString();                                                              
                            }
                        }
                        reader.Close();
                    }
                    finally
                    {
                        UserData.OpenConnection();
                    }
                    
                    Excel.Range Underline = exSheet.Range[exSheet.Cells[EditableRow, 1], exSheet.Cells[EditableRow, 5]];
                    Underline.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    EditableRow += 3;
                    
                    Excel.Range RightBorder = exSheet.Range[exSheet.Cells[1, 5], exSheet.Cells[EditableRow - 3, 5]];
                    RightBorder.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

                    exSheet.Columns.AutoFit();
                }
                #endregion

                #region Questions
                {
                    int EditableRow = 1;

                    bool isOddLine = true;

                    exSheet = exBook.Worksheets.Add();
                    exSheet.Activate();
                    exSheet.Name = "Questions";

                    Excel.Range NameRange = exSheet.Range[exSheet.Cells[EditableRow, 1], exSheet.Cells[EditableRow, 6]];
                    NameRange.Font.Italic = true;
                    NameRange.Font.Bold = true;
                    NameRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    NameRange.Interior.Color = ColorTranslator.ToOle( Color.LightGray );
                    exSheet.Cells[EditableRow, 1] = "Код вопроса";
                    exSheet.Cells[EditableRow, 2] = "Заголовок";
                    exSheet.Cells[EditableRow, 3] = "Текст вопроса";
                    exSheet.Cells[EditableRow, 4] = "Ответ";
                    exSheet.Cells[EditableRow, 5] = "Логин пользователя";
                    exSheet.Cells[EditableRow, 6] = "Код менеджера";


                    try
                    {
                        UserData.OpenConnection();
                        SqlCommand command = new SqlCommand( "GetQuestions", UserData.Connection );
                        command.CommandType = System.Data.CommandType.StoredProcedure;
                        SqlDataReader reader = command.ExecuteReader();
                        if ( reader.HasRows )
                        {
                            while ( reader.Read() )
                            {
                                EditableRow++;
                                if ( isOddLine )
                                {
                                    exSheet.Cells[EditableRow, 1].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                    exSheet.Cells[EditableRow, 2].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                    exSheet.Cells[EditableRow, 3].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                    exSheet.Cells[EditableRow, 4].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                    exSheet.Cells[EditableRow, 5].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                    exSheet.Cells[EditableRow, 6].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                    isOddLine = false;
                                }
                                else
                                {
                                    isOddLine = true;
                                }
                                exSheet.Cells[EditableRow, 1] = reader["ID_question"].ToString();
                                exSheet.Cells[EditableRow, 2] = reader["Header"].ToString();
                                exSheet.Cells[EditableRow, 3] = reader["QText"].ToString();
                                exSheet.Cells[EditableRow, 4] = reader["answer"].ToString();
                                exSheet.Cells[EditableRow, 5] = reader["login"].ToString();
                                exSheet.Cells[EditableRow, 6] = reader["ManagerID"].ToString();
                            }
                        }
                        reader.Close();
                    }
                    finally
                    {
                        UserData.OpenConnection();
                    }

                    Excel.Range Underline = exSheet.Range[exSheet.Cells[EditableRow, 1], exSheet.Cells[EditableRow, 6]];
                    Underline.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    EditableRow += 3;

                    Excel.Range RightBorder = exSheet.Range[exSheet.Cells[1, 6], exSheet.Cells[EditableRow - 3, 6]];
                    RightBorder.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

                    exSheet.Columns.AutoFit();
                }
                #endregion

                #region Statistics

                {
                    int EditableRow = 1;

                    bool isOddLine = true;

                    exSheet = exBook.Worksheets.Add();
                    exSheet.Activate();
                    exSheet.Name = "Statistics";

                    Excel.Range NameRange = exSheet.Range[exSheet.Cells[EditableRow, 1], exSheet.Cells[EditableRow++, 2]];
                    NameRange.Font.Italic = true;
                    NameRange.Font.Bold = true;
                    NameRange.Merge();
                    NameRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    NameRange.Interior.Color = ColorTranslator.ToOle( Color.LightGray );
                    NameRange.Value = "Количество вопросов, заданных пользователями";

                    exSheet.Cells[EditableRow, 1] = "Логин";
                    exSheet.Cells[EditableRow, 1].Interior.Color = ColorTranslator.ToOle( Color.Bisque );
                    exSheet.Cells[EditableRow, 2] = "Кол-во вопросов";
                    exSheet.Cells[EditableRow, 2].Interior.Color = ColorTranslator.ToOle( Color.Bisque );
                    
                    try
                    {
                        UserData.OpenConnection();
                        SqlCommand command = new SqlCommand( "GetStatisticData", UserData.Connection );
                        command.CommandType = System.Data.CommandType.StoredProcedure;
                        SqlDataReader reader = command.ExecuteReader();
                        if ( reader.HasRows )
                        {
                            while ( reader.Read() )
                            {
                                EditableRow++;
                                if ( isOddLine )
                                {
                                    exSheet.Cells[EditableRow, 1].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                    exSheet.Cells[EditableRow, 2].Interior.Color = ColorTranslator.ToOle( Color.PaleGreen );
                                    isOddLine = false;
                                }
                                else
                                {
                                    isOddLine = true;
                                }
                                exSheet.Cells[EditableRow, 1] = reader["login"].ToString();
                                exSheet.Cells[EditableRow, 2] = reader["count"].ToString();
                            }
                        }
                        reader.Close();
                    }
                    finally
                    {
                        UserData.OpenConnection();
                    }

                    Excel.Range Underline = exSheet.Range[exSheet.Cells[EditableRow, 1], exSheet.Cells[EditableRow, 2]];
                    Underline.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    EditableRow += 3;

                    Excel.Range RightBorder = exSheet.Range[exSheet.Cells[1, 2], exSheet.Cells[EditableRow - 3, 2]];
                    RightBorder.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

                    exSheet.Columns.AutoFit();
                }

                #endregion

                try
                {
                    string fname = folderDialog.SelectedPath + @"\GLPRDB_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year;
                    exBook.SaveAs( fname, Excel.XlFileFormat.xlOpenXMLWorkbook );
                    System.Windows.Forms.MessageBox.Show( "Отчет успешно создан" );
                }
                catch ( Exception E )
                {
                    System.Windows.Forms.MessageBox.Show( E.Message );
                }
                finally
                {
                    exBook.Close();
                    exApp.Quit();
                }
            }
        }

        #endregion

    }
}