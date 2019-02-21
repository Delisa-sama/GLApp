using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Data.SqlClient;

namespace GLApp
{
    /// <summary>
    /// Interaction logic for TechSupport.xaml
    /// </summary>
    public partial class TechSupport : Window
    {
        public TechSupport()
        {
            InitializeComponent();
            RefreshList();
        }

        List<QuestionItem> QList = new List<QuestionItem>();

        private void RefreshList()
        {
            try
            {
                UserData.OpenConnection();
                SqlCommand command = new SqlCommand( "GetQuestions" , UserData.Connection);
                command.CommandType = System.Data.CommandType.StoredProcedure;
                SqlDataReader reader = command.ExecuteReader();
                QuestionListBox.Items.Clear();
                QList.Clear();
                if ( reader.HasRows )
                {
                    while ( reader.Read() )
                    {
                        QuestionItem QItem = new QuestionItem();
                        QItem.ID =      reader.GetInt32( 0 );
                        QItem.Header =  reader.GetString( 1 );
                        QItem.Text =    reader.GetString( 2 );
                        QItem.Answer = DBNull.Value == reader.GetValue( 3 ) ? "Нет ответа" : reader.GetString( 3 );
                        QItem.UserLogin = reader["login"].ToString();
                        QList.Add( QItem );
                        TreeViewItem TVI = new TreeViewItem()
                        {
                            Tag = QList.IndexOf( QItem ),
                            Header = "Вопрос #" + QItem.ID.ToString() + ". " + QItem.Header
                        };
                        TVI.Items.Add( new TextBlock()
                        {
                            Text = new StringBuilder( QItem.Text + Environment.NewLine + "Ответ: " + QItem.Answer ).ToString()
                        } );
                        QuestionListBox.Items.Add( TVI );
                            
                    }
                }
                else
                {
                    QuestionListBox.Items.Add( new TextBlock()
                        {
                            TextWrapping = TextWrapping.Wrap,
                            Text = "Не найдено ни одного вопроса" 
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

        private void AskQuestionBtn_Click( object sender, RoutedEventArgs e )
        {
            if ( !string.IsNullOrWhiteSpace( HeaderTextBox.Text ) && !string.IsNullOrWhiteSpace( QuestionTextTB.Text ) )
            {
                try
                {
                    UserData.OpenConnection();
                    SqlCommand command = new SqlCommand( "AskQuestion", UserData.Connection );
                    command.CommandType = System.Data.CommandType.StoredProcedure;
                    command.Parameters.AddWithValue( "Text", QuestionTextTB.Text );
                    command.Parameters.AddWithValue( "Header", HeaderTextBox.Text );
                    command.Parameters.AddWithValue( "login", UserData.uname );
                    if ( command.ExecuteNonQuery() != 2 )
                    {
                        MessageBox.Show( "Ошибка размещения вопроса" );
                    }
                    else
                    {
                        MessageBox.Show( "Вопрос успешно размещен" );
                        this.Close();
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

        private void SearchBtn_Click( object sender, RoutedEventArgs e )
        {
            QuestionListBox.Items.Clear();
            foreach ( QuestionItem QI in QList )
            {
                if ( QI.Header.Contains( SearchString.Text ) || QI.ID.ToString().Contains( SearchString.Text ) )
                {
                    TreeViewItem TVI = new TreeViewItem()
                    {
                        Tag = QList.IndexOf( QI ),
                        Header = "Вопрос #" + QI.ID + ". " + QI.Header
                    };
                    TVI.Items.Add( new TextBlock()
                    {
                        Text = new StringBuilder( QI.Text + Environment.NewLine + "Ответ: " + QI.Answer ).ToString()
                    } );
                    QuestionListBox.Items.Add( TVI );
                }
            }
        }
    }
}
