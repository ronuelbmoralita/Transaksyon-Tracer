using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.InteropServices;
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
using ZXing;
using ZXing.Common;
using ZXing.QrCode;
using Brushes = System.Drawing.Brushes;
using Rectangle = System.Drawing.Rectangle;

namespace Transaksyon_Tracer
{
    /// <summary>
    /// Interaction logic for administrator.xaml
    /// </summary>
    public partial class AdminWindow : Window
    {
        //docx
        public Microsoft.Office.Interop.Word.Application wordApp = null;
        public Microsoft.Office.Interop.Word.Document myWordDoc = null;


        readonly string sqliteConnectionString = @"Data Source=//DESKTOP-7DENF7N\Transaksyon Tracer\Database\TransaksyonTracerDatabase.sqlite;Version=3;";

        readonly DispatcherTimer timer = new DispatcherTimer();

        private readonly BackgroundWorker worker = new BackgroundWorker();

        public AdminWindow()
        {
            InitializeComponent();
            //System.Diagnostics.PresentationTraceSources.DataBindingSource.Switch.Level = System.Diagnostics.SourceLevels.Critical;
            timer.Interval = TimeSpan.FromSeconds(10);
            timer.Tick += Timer_Tick;
            timer.Start();
            worker.DoWork += Worker_DoWork;
        }

        /////////////////////////////////////////////////////////////////////////////TIMER
        
        private void Timer_Tick(object sender, EventArgs e)
        {
            worker.RunWorkerAsync();
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                if (db_transaction.Items.Count == 0)
                {
                    return;
                }
                else
                {
                    using (SQLiteConnection source = new SQLiteConnection(@"Data Source=//DESKTOP-7DENF7N\Transaksyon Tracer\Database\TransaksyonTracerDatabase.sqlite;Version=3"))
                    using (SQLiteConnection destination = new SQLiteConnection(@"Data Source=C:\Transaksyon Tracer\Backup\Database\TransaksyonTracerDatabase.sqlite"))
                    {
                        //C
                        source.Open();
                        destination.Open();
                        source.BackupDatabase(destination, "main", "main", -1, null, 0);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /////////////////////////////////////////////////////////////////////////////LOAD/CLOSE FORM

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            staff.Text = LoginWindow.getFirstname;
            CbReason.Items.Add("Accepted");
            CbReason.Items.Add("Incomplete Documents");
            CbReason.Items.Add("Unauthorized Person");
            CbReason.Items.Add("Others");
            cb_transactionType.Items.Add("Registration of Certificate of Live Birth (On-Time)");
            cb_transactionType.Items.Add("Registration of Certificate of Live Birth (Delayed)");
            cb_transactionType.Items.Add("Registration of Certificate of Death (On-Time)");
            cb_transactionType.Items.Add("Registration of Certificate of Death (Delayed)");
            cb_transactionType.Items.Add("Registration of Certificate of Marriage (On-Time)");
            cb_transactionType.Items.Add("Registration of Certificate of Marriage (Delayed)");
            cb_transactionType.Items.Add("Application for Marriage License");
            cb_transactionType.Items.Add("Request for Civil Registry Forms under Act No.3753");
            cb_transactionType.Items.Add("Petitions under RA 9048/ RA 10172");
            cb_transactionType.Items.Add("Petitions under RA 9048/ RA 10172 (COD)");
            cb_transactionType.Items.Add("Petitions under RA 9048/ RA 10172 (COLB)");
            cb_transactionType.Items.Add("Petitions under RA 9048/ RA 10172 (COM)");
            cb_transactionType.Items.Add("Request for Migrant Petition to Other LCRO's");
            cb_transactionType.Items.Add("Request for Endorsement of Authentication/Advance Copy to PSA");
            cb_transactionType.Items.Add("Request for Endorsement OF Second Copy of Registrable Instruments to PSA");
            cb_transactionType.Items.Add("Request for Endorsement of Legitimation to PSA");
            cb_transactionType.Items.Add("Request for Endorsement of Accomplished Civil Registry Forms under RA 3753");
            cb_transactionType.Items.Add("Request for Out-of-Town Reporting of Registrable Instruments to Other LCRO's");
            cb_transactionType.Items.Add("Request for Endorsement of Court Orders and Judicial Decress to PSA");
            DisplayUserData();
            DisplayTransactionData();
            DisplayCode();
            DisplayQrCode();
            DisplayName();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            //Application.Current.Shutdown();
        }

        /////////////////////////////////////////////////////////////////////////////ADMINISTRATOR/DISPLAY DATA

        private void DisplayName()
        {
            try
            {
                if (!Directory.Exists(@"C:\Transaksyon Tracer"))
                {
                    return;
                }
                else
                {
                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        con.Open();
                        using (SQLiteCommand cmd = new SQLiteCommand("Select firstname from user_account where username like @uname", con))
                        {
                            cmd.Parameters.AddWithValue("@uname", staff.Text);
                            using (SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    logFirstname.Text = reader.GetString(0);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        string count_item;

        public string Count_item { get => count_item; set => count_item = value; }

        private void DisplayTransactionData()
        {
            try
            {
                if (!Directory.Exists(@"C:\Transaksyon Tracer"))
                {
                    return;
                }
                else
                {
                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        using (SQLiteDataAdapter sda = new SQLiteDataAdapter("Select * From transactions order by NOS desc", con))
                        {
                            using (DataSet dts = new DataSet())
                            {
                                con.Open();
                                sda.Fill(dts, "transactions");
                                db_transaction.ItemsSource = dts.Tables["transactions"].DefaultView;
                                //count_total.Content = "Total: " + db_transaction.Items.Count.ToString();
                                count_total.Content = db_transaction.Items.Count.ToString();
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void DisplayCode()
        {
            try
            {
                if (db_transaction.Items.Count == 0)
                {
                    return;
                }
                else
                {
                    using (SQLiteDataAdapter sda = new SQLiteDataAdapter("Select transactionCode from transactions ORDER BY NOS DESC", sqliteConnectionString))
                    {
                        DataTable dt_code = new DataTable();
                        sda.Fill(dt_code);
                        retrieveCode.Text = dt_code.Rows[0][0].ToString();
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void DisplayQrCode()
        {
            try
            {
                if (db_transaction.Items.Count == 0)
                {
                    return;
                }
                else
                {
                    var barcodeWriter = new BarcodeWriter
                    {
                        Format = BarcodeFormat.QR_CODE,
                        Options = new EncodingOptions
                        {
                            Height = 200,
                            Width = 200,
                            Margin = 0,
                            //Margin = 4,
                            PureBarcode = false
                        }
                    };


                    //string firstText = "Transaksyon Tracer";
                    //Rectangle rectf = new Rectangle(0, 0, 0, 0);

                    using (var bitmap = barcodeWriter.Write(retrieveCode.Text))
                    using (var stream = new MemoryStream())
                    {
                        /*
                        using (Graphics graphics = Graphics.FromImage(bitmap))
                        {
                            using (Font arialFont = new Font("Consolas", 15))
                            {

                                //graphics
                                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                                /*
                                StringFormat sf = new StringFormat();
                                sf.LineAlignment = StringAlignment.Center;
                                sf.Alignment = StringAlignment.Center;
                               

                                graphics.DrawString(firstText, arialFont, Brushes.Black, rectf);
                                   */


                        //bitmap
                        bitmap.Save(stream, ImageFormat.Png);
                        BitmapImage bi = new BitmapImage();
                        bi.BeginInit();
                        stream.Seek(0, SeekOrigin.Begin);
                        bi.StreamSource = stream;
                        bi.CacheOption = BitmapCacheOption.OnLoad;
                        bi.EndInit();
                        qrImage.Source = bi;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void DisplayAddressID()
        {
            try
            {
                if (db_transaction.Items.Count == 0)
                {
                    return;
                }
                else
                {
                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        con.Open();
                        DataTable dt = new DataTable();
                        using (SQLiteDataAdapter sda = new SQLiteDataAdapter("select distinct address from transactions", con))
                        {
                            DataTable dtidDocType = new DataTable();
                            using (SQLiteDataAdapter da = new SQLiteDataAdapter("select distinct idDocType from transactions", con))
                            {
                                da.Fill(dtidDocType);

                                idDoc.DisplayMemberPath = "idDocType";
                                idDoc.ItemsSource = dtidDocType.DefaultView;
                                con.Close();
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /////////////////////////////////////////////////////////////////////////////OTHER METHOD

        private void WaitCursor()
        {
            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;  
        }

        private void NormalCursor()
        {
            Mouse.OverrideCursor = null;
        }

        private void OnlyText(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("^[a-zA-Z]+$");
            e.Handled = !regex.IsMatch(e.Text);
        }

        private void OnlyNumber(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("^[0-9]*$");
            e.Handled = !regex.IsMatch(e.Text);
        }

        void Clear(DependencyObject obj)
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {

                if (obj is TextBox textbox)
                    textbox.Text = string.Empty;
                if (obj is CheckBox checkbox)
                    checkbox.IsChecked = false;
                if (obj is ComboBox combobox)
                    combobox.Text = string.Empty;
                if (obj is RadioButton Button)
                    Button.IsChecked = false;

                Clear(VisualTreeHelper.GetChild(obj, i));
            }
        }

        private void KillWordApp()
        {
            wordApp.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
            Marshal.FinalReleaseComObject(wordApp);
        }

        /////////////////////////////////////////////////////////////////////////////Button/MOUSELEFTButton

        private void ButtonClose_Click(object sender, RoutedEventArgs e)
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

        private void ButtonRefresh_Click(object sender, RoutedEventArgs e)
        {
            WaitCursor();
            Clear(this);
            DisplayTransactionData();
            DisplayCode();
            NormalCursor();
        }

        private void ButtonMinimize_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void ButtonEdit_Click(object sender, RoutedEventArgs e)
        {
            if (trackId.Text == string.Empty)
            {
                MessageBox.Show("Select valid item!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                MessageBoxResult result = MessageBox.Show("Save changes to Transation?", "Transaksyon Tracer", MessageBoxButton.YesNo, MessageBoxImage.Information, MessageBoxResult.No);
                if (result == MessageBoxResult.Yes)
                {
                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        using (SQLiteCommand cmd = con.CreateCommand())
                        {
                            WaitCursor();
                            con.Open();
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "update transactions  set transactionType=@transactionType,ownerFirstname=@ownerFirstname,ownerMiddlename=@ownerMiddlename,ownerLastname=@ownerLastname,authorizeFirstname=@authorizeFirstname,authorizeMiddlename=@authorizeMiddlename,authorizeLastname=@authorizeLastname,birthday=@birthday,idDocType=@idDocType,idDocNoRef=@idDocNoRef,address=@address,mobileNumber=@mobileNumber,emailAddress=@emailAddress,adviceGiven=@adviceGiven,succeedingAction=@action where NOS=" + trackId.Text;
                            cmd.Parameters.AddWithValue("@transactionType", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(TransType.Text));
                            cmd.Parameters.AddWithValue("@ownerFirstname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(ownerFirstname.Text));
                            cmd.Parameters.AddWithValue("@ownerMiddlename", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(ownerMiddlename.Text));
                            cmd.Parameters.AddWithValue("@ownerLastname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(ownerLastname.Text));
                            cmd.Parameters.AddWithValue("@authorizeFirstname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(ownerFirstname.Text));
                            cmd.Parameters.AddWithValue("@authorizeMiddlename", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(ownerMiddlename.Text));
                            cmd.Parameters.AddWithValue("@authorizeLastname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(ownerLastname.Text));
                            cmd.Parameters.AddWithValue("birthday", birthday.Text); cmd.Parameters.AddWithValue("@idDocType", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(idDoc.Text));
                            cmd.Parameters.AddWithValue("@idDocNoRef", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(idDocNoRef.Text));
                            cmd.Parameters.AddWithValue("@address", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(address.Text));
                            cmd.Parameters.AddWithValue("@mobileNumber", mobileNumber.Text);
                            cmd.Parameters.AddWithValue("@emailAddress", emailAddress.Text);
                            cmd.Parameters.AddWithValue("@status", CbReason.Text);
                            cmd.Parameters.AddWithValue("@adviceGiven", advice.Text);
                            cmd.Parameters.AddWithValue("@action", SucceedingAction.Text);
                            cmd.ExecuteNonQuery();

                            MessageBox.Show("Transaction has been successfully updated!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Question);
                            DisplayTransactionData();
                            Clear(this);
                            NormalCursor();
                        }
                    }

                }
                else
                {
                    return;
                }
            }
        }

        private void ButtonDelete_Click(object sender, RoutedEventArgs e)
        {
            if (trackId.Text == string.Empty)
            {
                MessageBox.Show("Select valid item!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                MessageBoxResult result = MessageBox.Show("Are you sure, you want to delete? This cannot be undone!", "Transaksyon Tracer", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
                if (result == MessageBoxResult.Yes)
                {

                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        using (SQLiteCommand cmd = con.CreateCommand())
                        {
                            con.Open();
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "delete from transactions where NOS=@nos";
                            cmd.Parameters.AddWithValue("@nos", trackId.Text);
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Deleted!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            DisplayTransactionData();
                            Clear(this);
                        }
                    }
                }
                else
                {
                    //DisplayTransactionData();
                    return;
                }
            }
        }      

        private void DeleteAll_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (db_transaction == null || db_transaction.Items.Count == 0)
                {
                    MessageBox.Show("No data found!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    MessageBoxResult delete_all = MessageBox.Show("Are you sure, you want to delete all data in the table? This cannot be undone.", "Transaksyon Tracer", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
                    if (delete_all == MessageBoxResult.Yes)
                    {

                        using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                        {
                            using (SQLiteCommand cmd = con.CreateCommand())
                            {
                                using (SQLiteCommand cmd_clean = con.CreateCommand())
                                {
                                    con.Open();
                                    cmd.CommandType = CommandType.Text;
                                    cmd_clean.CommandType = CommandType.Text;

                                    cmd.CommandText = "delete from transactions";
                                    cmd_clean.CommandText = "DELETE FROM sqlite_sequence WHERE name = 'transactions'";
                                    cmd_clean.CommandText = "edit sqlite_sequence SET seq = 0 WHERE name = 'transactions'";
                                    cmd.ExecuteNonQuery();
                                    cmd_clean.ExecuteNonQuery();
                                    DisplayTransactionData();
                                }
                            }
                        }
                    }
                    else if (delete_all == MessageBoxResult.No)
                    {
                        return;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void ButtonPrint_Click(object sender, RoutedEventArgs e)
        {
            if (db_transaction == null || db_transaction.Items.Count == 0)
            {
                MessageBox.Show("No record found!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else if (trackId.Text == string.Empty)
            {
                MessageBox.Show("Select valid item!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                if (!File.Exists(@"C:\Transaksyon Tracer\Documents\format.docx"))
                {
                    if (!File.Exists(@"C:\Program Files (x86)\OWL\Transaksyon Tracer\Documents\format.docx"))
                    {
                        MessageBox.Show("File not found!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                    else
                    {
                        File.Copy(@"C:\Program Files (x86)\OWL\Transaksyon Tracer\Documents\format.docx", @"C:\Transaksyon Tracer\Documents\format.docx");
                        WaitCursor();
                        CreateWordDocument(@"C:\Transaksyon Tracer\Documents\format.docx", @"C:\Transaksyon Tracer\Documents\mcr.docx");
                        NormalCursor();
                    }
                }
                else
                {
                    WaitCursor();
                    CreateWordDocument(@"C:\Transaksyon Tracer\Documents\format.docx", @"C:\Transaksyon Tracer\Documents\mcr.docx");
                    NormalCursor();
                }
            }
        }

        private void GoMain_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Go to main page?", "Transaksyon Tracer", MessageBoxButton.YesNo, MessageBoxImage.Information, MessageBoxResult.No);
            if (result == MessageBoxResult.Yes)
            {
                WaitCursor();
                MainWindow win = new MainWindow();
                win.Show();
                this.Close();
                NormalCursor();
            }
            else
            {
                return;
            }
        }

        int count = 1;

        private void DeveloperOption_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (count++ <= 1)
            {
                MessageBox.Show("Developer option is turned on!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Information);
                deleteAll.Visibility = Visibility.Visible;
            }
            else
            {
                MessageBox.Show("Developer option are already turned on, this feature will automatically turn off after you close the application!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void TechnicalSupport_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            openTechnical.IsOpen = true;
        }

        private void GoExit_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Are you sure, you want to Exit? ", "Transaksyon Tracer", MessageBoxButton.YesNo, MessageBoxImage.Information, MessageBoxResult.No);
            if (result == MessageBoxResult.Yes)
            {
                Application.Current.Shutdown();
            }
            else
            {
                return;
            }
        }

        /////////////////////////////////////////////////////////////////////////////TEXTCHANGED/SELECTIONCHANGED

        private void Search_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                Regex rx = new Regex("^[a-zA-Z]+$");
                if (rx.IsMatch(search.Text))
                {
                    DataView dv = db_transaction.ItemsSource as DataView;
                    dv.RowFilter = string.Format("ownerLastname LIKE '%{0}%' or authorizeLastname LIKE '{0}%'", search.Text);
                }
                else
                {
                    DataView dv = db_transaction.ItemsSource as DataView;
                    dv.RowFilter = "Convert(transactionCode, 'System.String') like '%" + search.Text + "%'";
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void TrackId_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (trackId.Text != string.Empty)
            {
                SucceedingAction.Visibility = Visibility.Visible;
                TransType.Visibility = Visibility.Visible;
            }
            else
            {
                SucceedingAction.Visibility = Visibility.Collapsed;
                TransType.Visibility = Visibility.Collapsed;
            }
        }

        private void Db_transaction_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                DataGrid gd = (DataGrid)sender;
                if (gd.SelectedItem is DataRowView row_selected)
                {
                    WaitCursor();
                    trackId.Text = row_selected["NOS"].ToString();
                    CbReason.Text = row_selected["status"].ToString();
                    TransType.Text = row_selected["transactionType"].ToString();
                    SucceedingAction.Text = row_selected["succeedingAction"].ToString();

                    advice.Text = row_selected["adviceGiven"].ToString();
                    retrieveCode.Text = row_selected["transactionCode"].ToString();

                    ownerFirstname.Text = row_selected["ownerFirstname"].ToString();
                    ownerMiddlename.Text = row_selected["ownerMiddlename"].ToString();
                    ownerLastname.Text = row_selected["ownerLastname"].ToString();

                    authorizeFirstname.Text = row_selected["authorizeFirstname"].ToString();
                    authorizeMiddlename.Text = row_selected["authorizeMiddlename"].ToString();
                    authorizeLastname.Text = row_selected["authorizeLastname"].ToString();

                    address.Text = row_selected["address"].ToString();
                    idDoc.Text = row_selected["idDocType"].ToString();
                    idDocNoRef.Text = row_selected["idDocNoRef"].ToString();
                    birthday.Text = row_selected["birthday"].ToString();
                    mobileNumber.Text = row_selected["mobileNumber"].ToString();
                    emailAddress.Text = row_selected["emailAddress"].ToString();
                    retrieveCode.Text = row_selected["transactionCode"].ToString();
                    DisplayQrCode();
                    NormalCursor();

                    DataView dv = db_transaction.ItemsSource as DataView;
                    dv.RowFilter = "Convert(transactionCode, 'System.String') like '%" + retrieveCode.Text + "%'"; //where n is a column name of the DataTable
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Please call Technical Support!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CbReason_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (CbReason.SelectedItem)
            {
                case "Accepted":
                    SucceedingAction.Visibility = Visibility.Collapsed;
                    break;
            }
        }

        private void Cb_transactionType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TransType.Text != string.Empty)
            {
                TransType.Visibility = Visibility.Visible;
            }
            else
            {
                TransType.Visibility = Visibility.Hidden;
            }
        }

        /////////////////////////////////////////////////////////////////////////////WORD/EXCEL

        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }

        private void CreateWordDocument(object filename, object SaveAs)
        {
            //save image
            String filePath = @"C:\Transaksyon Tracer\Images\qr.png";
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create((BitmapSource)qrImage.Source));
            using (FileStream stream = new FileStream(filePath, FileMode.Create)) encoder.Save(stream);

            try
            {

                //Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                wordApp = new Microsoft.Office.Interop.Word.Application();
                object missing = Missing.Value;

                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;

                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                           ref missing, ref missing, ref missing,
                                           ref missing, ref missing, ref missing,
                                           ref missing, ref missing, ref missing,
                                           ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();

                    this.FindAndReplace(wordApp, "<firstname>", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(ownerFirstname.Text));
                    this.FindAndReplace(wordApp, "<middlename>", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(ownerMiddlename.Text));
                    this.FindAndReplace(wordApp, "<lastname>", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(ownerLastname.Text));
                    this.FindAndReplace(wordApp, "<code>", retrieveCode.Text);
                    this.FindAndReplace(wordApp, "<reason>", "N/A");
                    this.FindAndReplace(wordApp, "<transaction>", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(TransType.Text));
                    this.FindAndReplace(wordApp, "<date>", DateTime.Now.ToString("MMMM dd, yyyy"));
                    this.FindAndReplace(wordApp, "<code>", retrieveCode.Text);
                    this.FindAndReplace(wordApp, "<number>", mobileNumber.Text);
                    this.FindAndReplace(wordApp, "<email>", emailAddress.Text);
                    this.FindAndReplace(wordApp, "<advice>", advice.Text);

                    switch (CbReason.SelectedItem)
                    {
                        case "Accepted":
                            this.FindAndReplace(wordApp, "<initial>", "Application Accepted");
                            this.FindAndReplace(wordApp, "<succeeding>", SucceedingAction.Text);
                            break;
                        default:
                            if (SucceedingAction.Text == string.Empty)
                            {
                                this.FindAndReplace(wordApp, "<succeeding>", SucceedingAction.Text);
                            }
                            else
                            {
                                foreach (Microsoft.Office.Interop.Word.Range tmpRange in myWordDoc.StoryRanges)
                                {
                                    object findText = "<succeeding>";
                                    object replaceText = SucceedingAction.Text;

                                    if (tmpRange.Find.Execute(ref findText, ref missing, ref missing,
                                    ref missing, ref missing, ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing, ref missing, ref missing,
                                    ref missing, ref missing))
                                    {
                                        tmpRange.Select();
                                        wordApp.Selection.Text = replaceText.ToString();
                                    }
                                }
                            }
                            this.FindAndReplace(wordApp, "<initial>", "Application Denied");
                            break;
                    }
                }
                else
                {
                    MessageBox.Show("File not found!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                //Save as
                myWordDoc.SaveAs(ref SaveAs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);

                var shape = myWordDoc.InlineShapes.AddPicture(@"C:\Transaksyon Tracer\Images\qr.png", false, true);
                shape.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                shape.ScaleWidth = 80;
                shape.ScaleHeight = 80;

                myWordDoc.Close();
                wordApp.Quit();

                MessageBoxResult result = MessageBox.Show("The file is ready click yes to open.", "Transaksyon Tracer", MessageBoxButton.YesNo, MessageBoxImage.Asterisk);
                if (result == MessageBoxResult.Yes)
                {
                    System.Diagnostics.Process.Start(@"C:\Transaksyon Tracer\Documents\mcr.docx");
                }
                else
                {
                    return;
                }
            }
            //catch (Exception ex)
            catch (Exception)
            {
                //MessageBox.Show("Error: " + ex.ToString(), "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
                MessageBox.Show("Word cannot save this file because it is already open elsewhere.", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
                KillWordApp();
            }
        }


        /////////////////////////////////////////////////////////////////////////////USER/DISPLAY USER DATA

        private void DisplayUserData()
        {
            try
            {
                if (!Directory.Exists(@"C:\Transaksyon Tracer"))
                {
                    return;
                }
                else
                {
                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        using (SQLiteDataAdapter sda = new SQLiteDataAdapter("Select * From user_account order by NOS desc", con))
                        {
                            using (DataSet ds = new DataSet())
                            {
                                con.Open();
                                sda.Fill(ds, "user_account");
                                db_userAccount.ItemsSource = ds.Tables["user_account"].DefaultView;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /////////////////////////////////////////////////////////////////////////////USER Button

        private void ButtonRefreshUser_Click(object sender, RoutedEventArgs e)
        {
            WaitCursor();
            DisplayUserData();
            Clear(this);
            NormalCursor();
        }

        private void ButtonSaveUser_Click(object sender, RoutedEventArgs e)
        {
            //var hasNumber = new Regex(@"[0-9]+");
            //var hasUpperChar = new Regex(@"[A-Z]+");
            Regex hasMiniMaxChars = new Regex(@".{8,15}");
            //var hasLowerChar = new Regex(@"[a-z]+");
            Regex hasSymbols = new Regex(@"[!@#$%^&*()_+=\[{\]};:<>|./?,-]");

            if (userFirstname.Text == string.Empty || 
                userLastname.Text == string.Empty ||
                userAddress.Text == string.Empty ||
                userMobileNumber.Text == string.Empty ||
                userName.Text == string.Empty ||
                userPassword.Text == string.Empty)
            {
                MessageBox.Show("All fields are required!", "Transaction Tracker", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else if (!hasMiniMaxChars.IsMatch(userPassword.Text) || !hasMiniMaxChars.IsMatch(userPassword.Text))
            {
                MessageBox.Show("Your password and username must be between 8 characters or more!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else if (hasSymbols.IsMatch(userName.Text) || hasSymbols.IsMatch(userPassword.Text))
            {
                MessageBox.Show("Not accepting special characters!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else if (adminCheckBox.IsChecked == false && standardCheckBox.IsChecked == false)
            {
                MessageBox.Show("Please select atleast 1 user type!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }
            else if (adminCheckBox.IsChecked == true)
            {
                using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                {
                    using (SQLiteCommand cmd = con.CreateCommand())
                    {
                        con.Open();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "insert into user_account(userType, username, password, firstname, lastname, address, mobileNumber) values(@userType ,@username,@password,@firstname,@lastname,@address,@mobileNumber)";
                        cmd.Parameters.AddWithValue("@userType", adminCheckBox.Content);
                        cmd.Parameters.AddWithValue("@username", userName.Text);
                        cmd.Parameters.AddWithValue("@password", userPassword.Text);
                        cmd.Parameters.AddWithValue("@firstname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(userFirstname.Text));
                        cmd.Parameters.AddWithValue("@lastname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(userLastname.Text));
                        cmd.Parameters.AddWithValue("@address", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(userAddress.Text));
                        cmd.Parameters.AddWithValue("@mobileNumber", userMobileNumber.Text);

                        using (SQLiteCommand command = new SQLiteCommand("Select count (*) from user_account where username = '" + userName.Text + "'", con))
                        {
                            using (SQLiteCommand command1 = new SQLiteCommand("Select count (*) from user_account where password = '" + userPassword.Text + "'", con))
                            {
                                var result = command.ExecuteScalar();
                                int i = Convert.ToInt32(result);
                                var result1 = command.ExecuteScalar();
                                int i1 = Convert.ToInt32(result1);
                                if (i != 0 || i1 != 0)
                                {
                                    MessageBox.Show("Username or Password already exist!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning); ;
                                    Clear(this);
                                    return;
                                }
                                else
                                {
                                    cmd.ExecuteNonQuery();
                                    MessageBox.Show("User successfully saved!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Information);
                                    LoginWindow win = new LoginWindow();
                                    win.Show();
                                    this.Close();
                                }
                            }
                        }
                    }
                }
            }
            else if (standardCheckBox.IsChecked == true)
            {
                using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                {
                    using (SQLiteCommand cmd = con.CreateCommand())
                    {
                        con.Open();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "insert into user_account(userType, username, password, firstname, lastname, address, mobileNumber) values(@userType ,@username,@password,@firstname,@lastname,@address,@mobileNumber)";
                        cmd.Parameters.AddWithValue("@userType", standardCheckBox.Content);
                        cmd.Parameters.AddWithValue("@username", userName.Text);
                        cmd.Parameters.AddWithValue("@password", userPassword.Text);
                        cmd.Parameters.AddWithValue("@firstname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(userFirstname.Text));
                        cmd.Parameters.AddWithValue("@lastname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(userLastname.Text));
                        cmd.Parameters.AddWithValue("@address", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(userAddress.Text));
                        cmd.Parameters.AddWithValue("@mobileNumber", userMobileNumber.Text);

                        using (SQLiteCommand command = new SQLiteCommand("Select count (*) from user_account where username = '" + userName.Text + "'", con))
                        {
                            using (SQLiteCommand command1 = new SQLiteCommand("Select count (*) from user_account where password = '" + userPassword.Text + "'", con))
                            {
                                var result = command.ExecuteScalar();
                                int i = Convert.ToInt32(result);
                                var result1 = command.ExecuteScalar();
                                int i1 = Convert.ToInt32(result1);
                                if (i != 0 || i1 != 0)
                                {
                                    MessageBox.Show("Username or Password already exist!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
                                    Clear(this);
                                    return;
                                }
                                else
                                {
                                    cmd.ExecuteNonQuery();
                                    MessageBox.Show("User save in database!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Information);
                                    LoginWindow win = new LoginWindow();
                                    win.Show();
                                    this.Close();
                                }
                            }
                        }
                    }
                }
            }
        }

        private void ButtonEditUser_Click(object sender, RoutedEventArgs e)
        {
            if (userFirstname.Text == string.Empty ||
                userLastname.Text == string.Empty ||
                userAddress.Text == string.Empty ||
                userMobileNumber.Text == string.Empty)
            {
                MessageBox.Show("All fields are required!", "Transaction Tracker", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else if (userId.Text == string.Empty)
            {
                MessageBox.Show("Select valid item!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else
            {
                MessageBoxResult result = MessageBox.Show("Save changes to User Account?", "Transaksyon Tracer", MessageBoxButton.YesNo, MessageBoxImage.Information, MessageBoxResult.No);
                if (result == MessageBoxResult.Yes)
                {
                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        using (SQLiteCommand cmd = con.CreateCommand())
                        {
                            DataRowView drv = (DataRowView)db_userAccount.SelectedItem;
                            con.Open(); cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "update user_account set userType=@userType,firstname=@firstname,lastname=@lastname,username=@username,password=@password,mobileNumber=@mobileNumber where NOS=" + userId.Text;
                            cmd.Parameters.AddWithValue("@firstname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(userFirstname.Text));
                            cmd.Parameters.AddWithValue("@lastname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(userLastname.Text));
                            cmd.Parameters.AddWithValue("@username", userName.Text);
                            cmd.Parameters.AddWithValue("@password", userPassword.Text);
                            cmd.Parameters.AddWithValue("@mobileNumber", userMobileNumber.Text);

                            if (adminCheckBox.IsChecked == true)
                            {
                                cmd.Parameters.AddWithValue("userType", adminCheckBox.Content);
                            }
                            else if (standardCheckBox.IsChecked == true)
                            {
                                cmd.Parameters.AddWithValue("userType", standardCheckBox.Content);
                            }
                            cmd.ExecuteNonQuery();
                            DisplayUserData();
                            MessageBox.Show("User has been successfully updated!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Question);
                            Clear(this);

                            /*
                           using (SQLiteCommand command = new SQLiteCommand("Select count (*) from user_account where username = '" + userName.Text + "'", con))
                           {
                               using (SQLiteCommand command1 = new SQLiteCommand("Select count (*) from user_account where password = '" + userPassword.Text + "'", con))
                               {
                                   var resultSecond = command.ExecuteScalar();
                                   int i = Convert.ToInt32(resultSecond);
                                   var result1 = command.ExecuteScalar();
                                   int i1 = Convert.ToInt32(result1);
                                   if (i != 0 && i1 != 0)
                                   {
                                       MessageBox.Show("Username or Password already exist!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning); ;
                                       Clear(this);
                                       return;
                                   }
                                   else
                                   {
                                       cmd.ExecuteNonQuery();
                                       MessageBox.Show("User successfully saved!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Information);
                                       DisplayUserData();
                                       Clear(this);
                                   }
                               }
                           }
                           */
                        }
                    }
                }
                else
                {
                    return;
                }
            }
        }

        private void ButtonDeleteUser_Click(object sender, RoutedEventArgs e)
        {
            if (db_userAccount.Items.Count == 1)
            {
                MessageBox.Show("User database cannot be empty!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (userId.Text == string.Empty)
            {
                MessageBox.Show("Select valid item!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else
            {
                MessageBoxResult result = MessageBox.Show("Are you sure, you want to delete? This cannot be undone!", "Transaksyon Tracer", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
                if (result == MessageBoxResult.Yes)
                {

                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        using (SQLiteCommand cmd = con.CreateCommand())
                        {
                            con.Open();
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "delete from user_account where NOS=@nos";
                            cmd.Parameters.AddWithValue("@nos", userId.Text);
                            cmd.ExecuteNonQuery();
                            DisplayUserData();
                            Clear(this);
                        }
                    }
                }
                else
                {
                    //DisplayUserData();
                    return;
                }
            }
        }

        /////////////////////////////////////////////////////////////////////////////USER/TEXTCHANGED/SELECTIONCHANGED

        private void SearchUser_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                Regex rx = new Regex("^[a-zA-Z]+$");
                if (rx.IsMatch(searchUser.Text))
                {
                    DataView dv = db_userAccount.ItemsSource as DataView;
                    dv.RowFilter = string.Format("lastname LIKE '%{0}%' or firstname LIKE '{0}%'", searchUser.Text);
                }
                else
                {
                    DataView dv = db_userAccount.ItemsSource as DataView;
                    dv.RowFilter = "Convert(NOS, 'System.String') like '%" + searchUser.Text + "%'";
                }
            }
            catch (Exception)
            {
                MessageBox.Show("No data found!");
            }
        }

        private void Db_userAccount_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                DataGrid gd = (DataGrid)sender;
                if (gd.SelectedItem is DataRowView row_selected)
                {
                    userId.Text = row_selected["NOS"].ToString();
                    userFirstname.Text = row_selected["firstname"].ToString();
                    userLastname.Text = row_selected["lastname"].ToString();
                    userName.Text = row_selected["username"].ToString();
                    userPassword.Text = row_selected["password"].ToString();
                    userAddress.Text = row_selected["address"].ToString();
                    userMobileNumber.Text = row_selected["mobileNumber"].ToString();

                    DataView dv = db_userAccount.ItemsSource as DataView;
                    dv.RowFilter = "Convert(NOS, 'System.String') like '%" + userId.Text + "%'";

                    using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                    {
                        using (SQLiteCommand cmd_username = con.CreateCommand())
                        {
                            con.Open();
                            cmd_username.CommandType = CommandType.Text;
                            cmd_username.CommandText = "Select * from user_account where username = '" + userName.Text.Trim() + "'and password = '" + userPassword.Text.Trim() + "'";
                            SQLiteDataReader sdr_admin;
                            sdr_admin = cmd_username.ExecuteReader();
                            string userRole_admin = string.Empty;

                            while (sdr_admin.Read())
                            {
                                userRole_admin = sdr_admin["userType"].ToString();
                            }
                            if (userRole_admin == "Administrator")
                            {
                                adminCheckBox.IsChecked = true;
                            }
                            else if (userRole_admin == "Standard")
                            {
                                standardCheckBox.IsChecked = true;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Please call Technical Support!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Enable_developer_Click(object sender, RoutedEventArgs e)
        {
            deleteAll.Visibility = Visibility.Visible;
        }

        private void Disable_developer_Click(object sender, RoutedEventArgs e)
        {
            deleteAll.Visibility = Visibility.Collapsed;
        }

        private void UserId_TextChanged(object sender, TextChangedEventArgs e)
        {
            if(userId.Text == string.Empty)
            {
                disableUser.IsHitTestVisible = true;
            }
            else
            {
                disableUser.IsHitTestVisible = false;
            }
        }

        private void ButtonDeveloper_Click(object sender, RoutedEventArgs e)
        {
            WaitCursor();
            MainWindow open = new MainWindow();
            open.Show();
            Hide();
            NormalCursor();
        }

        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private void ButtonGoAdmin_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Login to your Account?", "Transaksyon Tracer", MessageBoxButton.YesNo, MessageBoxImage.Information, MessageBoxResult.No);
            if (result == MessageBoxResult.Yes)
            {
                WaitCursor();
                LoginWindow win = new LoginWindow();
                win.Show();
                this.Close();
                NormalCursor();
            }
            else
            {
                return;
            }
        }
    }
}
