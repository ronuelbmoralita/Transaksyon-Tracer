using System;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using ZXing;
using ZXing.Common;
using Brushes = System.Drawing.Brushes;

namespace Transaksyon_Tracer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
       
        //excel
        public Microsoft.Office.Interop.Excel.Application excel = null;
        public Microsoft.Office.Interop.Excel.Workbook wb = null;
        public Microsoft.Office.Interop.Excel.Worksheet ws = null;

        //docx
        public Microsoft.Office.Interop.Word.Application wordApp = null;
        public Microsoft.Office.Interop.Word.Document myWordDoc = null;


        //sqlite connection
        readonly string sqliteConnectionString = @"Data Source=//DESKTOP-7DENF7N\Transaksyon Tracer\Database\TransaksyonTracerDatabase.sqlite;Version=3;";

        //readonly string connectionString = @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=db_transaction;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";

        readonly DispatcherTimer timer = new DispatcherTimer();

        private readonly BackgroundWorker worker = new BackgroundWorker();

        public MainWindow()
        {
            InitializeComponent();
            //System.Diagnostics.PresentationTraceSources.DataBindingSource.Switch.Level = System.Diagnostics.SourceLevels.Critical;
            /**/
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
            //month.SelectedValue = DateTime.Now.Month;
            //int monthInDigit = DateTime.ParseExact(month.Text, "MMM", CultureInfo.InvariantCulture).Month;

            month.ItemsSource = System.Globalization.CultureInfo.InvariantCulture.DateTimeFormat.MonthNames.Take(12).ToArray();
            //month.ItemsSource = System.Globalization.CultureInfo.InvariantCulture.DateTimeFormat.DaysInMonth.Take(31).ToList();

            int[] days = Enumerable.Range(1, DateTime.DaysInMonth(12, 12)).ToArray();
            day.ItemsSource = days;


            year.ItemsSource = Enumerable.Range(1800, DateTime.Now.Year - 1800 + 1).ToList();
            //year.SelectedItem = "1800";

             /*
            for (int i = 1800; i <= DateTime.Now.Year; i++)
            {
                ComboBoxItem item = new ComboBoxItem
                {
                    Content = i
                };
                year.Items.Add(item);
            }
            */
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

            DisplayTransactionData();
            DisplayCode();
            DisplayQrCode();
            DisplayAddressID();
            DisplayName();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            //Application.Current.Shutdown();
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

        private void DisabledHeader()
        {
            db_transaction.IsHitTestVisible = false;
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
            try
            {
                wordApp.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
                Marshal.FinalReleaseComObject(wordApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.ToString(), "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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

        /////////////////////////////////////////////////////////////////////////////DISPLAY DATA

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

        private string count_item;
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
                                //count_total.Content = "Total: " + db_transaction.Items.Count.ToString(); }
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


        /////////////////////////////////////////////////////////////////////////////Button/MOUSELEFTButtonUP

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

        private void ButtonMinimize_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void ButtonRefresh_Click(object sender, RoutedEventArgs e)
        {
            WaitCursor();
            Clear(this);
            DisplayTransactionData();
            DisplayCode();
            NormalCursor();
        }

        private void ButtonSave_Click(object sender, RoutedEventArgs e)
        {
            if (cb_transactionType.Text == string.Empty ||
                ownerFirstname.Text == string.Empty ||
                ownerMiddlename.Text == string.Empty ||
                ownerLastname.Text == string.Empty ||
                birthday.Text == string.Empty ||
                address.Text == string.Empty ||
                //idDoc.Text == string.Empty ||
                //idDocNoRef.Text == string.Empty ||
                //mobileNumber.Text == string.Empty ||
                advice.Text == string.Empty)
            {
                MessageBox.Show("All fields are required!", "Transaksyon Tracker", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            /*
            else if (CbReason.Text != "Accepted" && SucceedingAction.Text == string.Empty)
            {
                MessageBox.Show("Please enter valid succeeding action for denied documents!", "Transaction Tracker", MessageBoxButton.OK, MessageBoxImage.Warning);
            }   
            */
            else
            {
                try
                {
                    if (!Directory.Exists(@"C:\Transaksyon Tracer"))
                    {
                        MessageBox.Show("Cannot find Transaksyon Tracer directory!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    else
                    {
                        using (SQLiteConnection con = new SQLiteConnection(sqliteConnectionString))
                        {

                            using (SQLiteCommand cmd = con.CreateCommand())
                            {
                                WaitCursor();
                                con.Open();
                                cmd.CommandType = CommandType.Text;
                                cmd.CommandText = "insert into transactions(transactionCode,dateTime, status, transactionType,ownerFirstname,ownerMiddlename,ownerLastname,authorizeFirstname,authorizeMiddlename,authorizeLastname,birthday,idDocType,idDocNoRef,address,mobileNumber,emailAddress,adviceGiven,SucceedingAction,staff)" +
                                " values(@transactionCode,@dateTime,@status,@transactionType,@ownerFirstname,@ownerMiddlename,@ownerLastname,@authorizeFirstname,@authorizeMiddlename,@authorizeLastname,@birthday,@idDocType,@idDocNoRef,@address,@mobileNumber,@emailAddress,@adviceGiven,@SucceedingAction,@staff)";

                                cmd.Parameters.AddWithValue("@transactionCode", DateTime.Now.ToString("yyMMddHHmmssfff"));
                                cmd.Parameters.AddWithValue("@dateTime", DateTime.Now.ToString());
                                cmd.Parameters.AddWithValue("@transactionType", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(cb_transactionType.Text));
                                cmd.Parameters.AddWithValue("@ownerFirstname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(ownerFirstname.Text));
                                cmd.Parameters.AddWithValue("@ownerMiddlename", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(ownerMiddlename.Text));
                                cmd.Parameters.AddWithValue("@ownerLastname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(ownerLastname.Text));
                                cmd.Parameters.AddWithValue("birthday", birthday.Text);

                                //cmd.Parameters.AddWithValue("birthday", day.Text + "/" + month.SelectedIndex + 1 + "/" + year.Text);
                                cmd.Parameters.AddWithValue("@address", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(address.Text));
                                cmd.Parameters.AddWithValue("@emailAddress", emailAddress.Text);
                                cmd.Parameters.AddWithValue("@status", CbReason.Text);
                                cmd.Parameters.AddWithValue("@SucceedingAction", SucceedingAction.Text);
                                cmd.Parameters.AddWithValue("@adviceGiven", advice.Text);
                                cmd.Parameters.AddWithValue("@staff", logFirstname.Text);

                                if (mobileNumber.Text == string.Empty)
                                {
                                    cmd.Parameters.AddWithValue("@mobileNumber", "N/A");
                                }
                                else
                                {
                                    cmd.Parameters.AddWithValue("@mobileNumber", mobileNumber.Text);
                                }

                                if (idDoc.Text == string.Empty || idDocNoRef.Text == string.Empty)
                                {
                                    cmd.Parameters.AddWithValue("@idDocType", "N/A");
                                    cmd.Parameters.AddWithValue("@idDocNoRef", "N/A");
                                }
                                else
                                {
                                    cmd.Parameters.AddWithValue("@idDocType", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(idDoc.Text));
                                    cmd.Parameters.AddWithValue("@idDocNoRef", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(idDocNoRef.Text));
                                }

                                if (authorizeFirstname.Text == string.Empty || authorizeMiddlename.Text == string.Empty || authorizeLastname.Text == string.Empty)
                                {
                                    cmd.Parameters.AddWithValue("@authorizeFirstname", "N/A");
                                    cmd.Parameters.AddWithValue("@authorizeMiddlename", "N/A");
                                    cmd.Parameters.AddWithValue("@authorizeLastname", "N/A");
                                }
                                else
                                {
                                    cmd.Parameters.AddWithValue("@authorizeFirstname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(authorizeFirstname.Text));
                                    cmd.Parameters.AddWithValue("@authorizeMiddlename", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(authorizeMiddlename.Text));
                                    cmd.Parameters.AddWithValue("@authorizeLastname", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(authorizeLastname.Text));
                                }
                                switch (CbReason.SelectedItem)
                                {
                                    case "Accepted":
                                        cmd.Parameters.AddWithValue("@SucceedingAction", "N/A");
                                        break;
                                    default:
                                        cmd.Parameters.AddWithValue("@SucceedingAction", SucceedingAction.Text);
                                        break;
                                }
                                cmd.ExecuteNonQuery();

                                MessageBox.Show("Transaction successfully saved into Database!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Asterisk);

                                DisplayTransactionData();
                                DisplayCode();
                                DisplayQrCode();
                                NormalCursor();


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
                                        Clear(this);
                                    }
                                }
                                else
                                {
                                    WaitCursor();
                                    CreateWordDocument(@"C:\Transaksyon Tracer\Documents\format.docx", @"C:\Transaksyon Tracer\Documents\mcr.docx");
                                    NormalCursor();
                                    Clear(this);
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
        }

        private void ButtonEdit_Click(object sender, RoutedEventArgs e)
        {
            try
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
                                cmd.CommandText = "update transactions set SucceedingAction=@action where NOS=" + trackId.Text; cmd.Parameters.AddWithValue("@action", SucceedingAction.Text);
                                cmd.ExecuteNonQuery();

                                //MessageBox.Show("Transaction has been successfully updated!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Question);
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
            catch (Exception)
            {
                throw;
            }
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (trackId.Text == string.Empty)
                {
                    MessageBox.Show("No item selected, please select into table!", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    string var;
                    var = "Trasansaction Code " + retrieveCode.Text;
                    MessageBoxResult result = MessageBox.Show("Are you sure, you want to delete?  " + var + ".", "Transaction Tracker", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
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
                                DisplayTransactionData();
                                Clear(this);
                            }
                        }
                    }
                    else
                    {
                        DisplayTransactionData();
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
            try
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
                    _ = " Total Items: " + db_transaction.Items.Count.ToString();
                    MessageBoxResult result = MessageBox.Show("You are about to generate a document, Proceed?", "Transaksyon Tracer", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
                    if (result == MessageBoxResult.No)
                    {
                        return;
                    }
                    else
                    {       /*
                    WaitCursor();
                    CreateWordDocument(@"C:\Transaksyon Tracer\Documents\format.docx", @"C:\Transaksyon Tracer\Documents\mcr.docx");
                    NormalCursor();
                    */

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
                        /**/
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void ButtonGoAdmin_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("You need to login first, Continue?", "Transaksyon Tracer", MessageBoxButton.YesNo, MessageBoxImage.Information, MessageBoxResult.No);
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

        private void TechnicalSupport_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            openTechnical.IsOpen = true;
        }

        private void GoAdmin_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("You need to login first, Continue?", "Transaksyon Tracer", MessageBoxButton.YesNo, MessageBoxImage.Information, MessageBoxResult.No);
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

        private void GoExit_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
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
                   
                    this.FindAndReplace(wordApp, "<transaction>", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(TransType.Text));
                    this.FindAndReplace(wordApp, "<date>", DateTime.Now.ToString("MMMM dd, yyyy"));
                    this.FindAndReplace(wordApp, "<code>", retrieveCode.Text);
                    this.FindAndReplace(wordApp, "<number>", mobileNumber.Text);
                    this.FindAndReplace(wordApp, "<email>", emailAddress.Text);
                    this.FindAndReplace(wordApp, "<advice>", advice.Text);
                    this.FindAndReplace(wordApp, "<appname>", "Transaksyon Tracer");

                    switch (CbReason.SelectedItem)
                    {
                        case "Accepted":
                            this.FindAndReplace(wordApp, "<reason>", "N/A");
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
                            this.FindAndReplace(wordApp, "<reason>", CultureInfo.CurrentCulture.TextInfo.ToTitleCase(CbReason.Text));
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

                System.Diagnostics.Process.Start(@"C:\Transaksyon Tracer\Documents\mcr.docx");
                /*
                MessageBoxResult result = MessageBox.Show("The file is ready click yes to open.", "Transaksyon Tracer", MessageBoxButton.YesNo, MessageBoxImage.Asterisk);
                if (result == MessageBoxResult.Yes)
                {
                    System.Diagnostics.Process.Start(@"C:\Transaksyon Tracer\Documents\mcr.docx");
                }
                else
                {
                    return;
                }
                */
            }
            //catch (Exception ex)
            catch (Exception)
            {
                //MessageBox.Show("Error: " + ex.ToString(), "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
                MessageBox.Show("Word cannot save this file because it is already open elsewhere.", "Transaksyon Tracer", MessageBoxButton.OK, MessageBoxImage.Error);
                KillWordApp();
            }
        }

        /////////////////////////////////////////////////////////////////////////////TEXTCHANGED/SELECTIONCHANGED/MOUSEKEYDOWN

        private void Search_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                Regex rx = new Regex("^[a-zA-Z]+$");
                if (rx.IsMatch(search.Text))
                {
                    DataView dv = db_transaction.ItemsSource as DataView;
                    dv.RowFilter = string.Format("ownerLastname LIKE '%{0}%'", search.Text);
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
                ButtonSave.IsEnabled = false;
                TransType.Visibility = Visibility.Visible;
                firstStack.IsHitTestVisible = false;
                secondStack.IsHitTestVisible = false;
                MaterialDesignThemes.Wpf.ButtonProgressAssist.SetIsIndicatorVisible(ButtonPrint, true);
                cb_transactionType.Visibility = Visibility.Collapsed;
            }
            else
            {
                SucceedingAction.Visibility = Visibility.Collapsed;
                ButtonSave.IsEnabled = true;
                TransType.Visibility = Visibility.Hidden;
                firstStack.IsHitTestVisible = true;
                secondStack.IsHitTestVisible = true;
                MaterialDesignThemes.Wpf.ButtonProgressAssist.SetIsIndicatorVisible(ButtonPrint, false);
                cb_transactionType.Visibility = Visibility.Visible;
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
                    SucceedingAction.Text = row_selected["SucceedingAction"].ToString();
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
                    DisplayQrCode();
                    NormalCursor();

                    DataView dv = db_transaction.ItemsSource as DataView;
                    dv.RowFilter = "Convert(transactionCode, 'System.String') like '%" + retrieveCode.Text + "%'"; //where n is a column name of the DataTable
                }
            }
            catch (Exception)
            {
                throw;
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

        private void SucceedingAction_KeyDown(object sender, KeyEventArgs e)
        {
            if (CbReason.Text == "Accepted")
            {
                return;
            }
            else if (!(e.KeyboardDevice.Modifiers == ModifierKeys.Shift) && e.KeyboardDevice.IsKeyDown(Key.Enter))
            {
                if (SucceedingAction.Text != string.Empty)
                {
                    //SucceedingAction.Text += DateTime.Now.ToString() + "\n";
                    SucceedingAction.Text += ", " + DateTime.Now.ToString() + " - ";
                    SucceedingAction.ScrollToEnd();
                    SucceedingAction.Focus();
                    SucceedingAction.Select(0, 0);
                    SucceedingAction.Select(SucceedingAction.Text.Length, 0);
                }
                else
                {
                    //SucceedingAction.Text += DateTime.Now.ToString() + "\n";
                    SucceedingAction.Text += DateTime.Now.ToString() + " - ";
                    SucceedingAction.ScrollToEnd();
                    SucceedingAction.Focus();
                    SucceedingAction.Select(0, 0);
                    SucceedingAction.Select(SucceedingAction.Text.Length, 0);
                }
            }
        }

        private void SucceedingAction_TextChanged(object sender, TextChangedEventArgs e)
        {
            SucceedingAction.ScrollToEnd();
            SucceedingAction.Focus();
            SucceedingAction.Select(0, 0);
            SucceedingAction.Select(SucceedingAction.Text.Length, 0);
        }

        private void CbReason_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (CbReason.Text)
            {
                case "Accepted":
                    SucceedingAction.Visibility = Visibility.Collapsed;
                    break;
            }
        }

        private void ButtonDeveloper_Click(object sender, RoutedEventArgs e)
        {
            WaitCursor();
            AdminWindow open = new AdminWindow();
            open.Show();
            Hide();
            NormalCursor();
        }

        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }
    }
}

