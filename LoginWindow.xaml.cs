using System;
using System.Data.SQLite;
using System.IO;
using System.Windows;

namespace QuanLyBanHang
{
    public partial class LoginWindow : Window
    {
        private readonly string dbPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DataQL.db");

        public LoginWindow()
        {
            InitializeComponent();
            KhoiTaoDatabase();
        }

        private void KhoiTaoDatabase()
        {
            try
            {
                string dir = Path.GetDirectoryName(dbPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                if (!File.Exists(dbPath))
                {
                    SQLiteConnection.CreateFile(dbPath);
                }

                using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();

                    new SQLiteCommand("CREATE TABLE IF NOT EXISTS Users (UserID INTEGER PRIMARY KEY AUTOINCREMENT, Username TEXT UNIQUE, Password TEXT, Role TEXT)", conn).ExecuteNonQuery();
                    try { new SQLiteCommand("ALTER TABLE Users ADD COLUMN AllowedTabs TEXT DEFAULT 'tabKho,tabBan,tabThongKe,tabHeThong'", conn).ExecuteNonQuery(); } catch { }
                    new SQLiteCommand("INSERT OR IGNORE INTO Users (Username, Password, Role, AllowedTabs) VALUES ('admin', 'admin', 'Admin', 'tabKho,tabBan,tabThongKe,tabHeThong')", conn).ExecuteNonQuery();

                    new SQLiteCommand("CREATE TABLE IF NOT EXISTS Inventory (ProductID INTEGER PRIMARY KEY AUTOINCREMENT, ProductName TEXT UNIQUE, Category TEXT, Unit TEXT, Quantity INTEGER DEFAULT 0, ImportPrice REAL DEFAULT 0, SafeLevel INTEGER DEFAULT 0, Importer TEXT, ImportDate TEXT)", conn).ExecuteNonQuery();
                    new SQLiteCommand("CREATE TABLE IF NOT EXISTS Sales (SaleID INTEGER PRIMARY KEY AUTOINCREMENT, ProductID INTEGER, QuantitySold INTEGER, SalePrice REAL, SaleDate TEXT, UserID INTEGER, Unit TEXT, Seller TEXT)", conn).ExecuteNonQuery();
                    new SQLiteCommand("CREATE TABLE IF NOT EXISTS ImportLogs (LogID INTEGER PRIMARY KEY AUTOINCREMENT, ProductID INTEGER, Qty INTEGER, Price REAL, ImportDate TEXT, Importer TEXT, EntryTime TEXT)", conn).ExecuteNonQuery();

                    new SQLiteCommand("CREATE TABLE IF NOT EXISTS DS_Loai (CategoryName TEXT UNIQUE)", conn).ExecuteNonQuery();
                    new SQLiteCommand("CREATE TABLE IF NOT EXISTS DS_SanPham (ProductName TEXT UNIQUE, CategoryName TEXT)", conn).ExecuteNonQuery();
                    new SQLiteCommand("CREATE TABLE IF NOT EXISTS DS_DonVi (UnitName TEXT UNIQUE)", conn).ExecuteNonQuery();

                    if ((long)new SQLiteCommand("SELECT COUNT(*) FROM DS_DonVi", conn).ExecuteScalar() == 0)
                    {
                        new SQLiteCommand("INSERT INTO DS_DonVi VALUES ('Cái'), ('Hộp'), ('Bình'), ('Chiếc')", conn).ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi thiết lập Database ban đầu: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void KiemTraCapsLock()
        {
            // Bật tắt cái bong bóng dựa vào đèn Caps Lock
            if (System.Windows.Input.Keyboard.IsKeyToggled(System.Windows.Input.Key.CapsLock))
                popCapsLock.IsOpen = true;
            else
                popCapsLock.IsOpen = false;
        }

        private void txtPassword_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            KiemTraCapsLock();
        }

        private void txtPassword_GotFocus(object sender, RoutedEventArgs e)
        {
            KiemTraCapsLock();
        }

        // Thêm hàm này để khi khách bấm chuột ra chỗ khác thì cái bong bóng tự ẩn đi cho gọn
        private void txtPassword_LostFocus(object sender, RoutedEventArgs e)
        {
            popCapsLock.IsOpen = false;
        }

        private void btnLogin_Click(object sender, RoutedEventArgs e)
        {
            // TÊN ĐÃ KHỚP HOÀN TOÀN VỚI FILE XAML Ở TRÊN
            string user = txtUsername.Text.Trim();
            string pass = txtPassword.Password;

            if (string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ Tài khoản và Mật khẩu!", "Cảnh báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Tạo các biến để hứng dữ liệu tạm thời
            bool isLoginSuccess = false;
            string role = "";
            string tabs = "";

            try
            {
                using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();

                    // 👉 TUI ĐÃ SỬA DÒNG NÀY: Thêm 'COLLATE NOCASE' vào ngay sau Username
                    string sql = "SELECT Role, AllowedTabs FROM Users WHERE Username = @u COLLATE NOCASE AND Password = @p";

                    using (var cmd = new SQLiteCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@u", user);
                        cmd.Parameters.AddWithValue("@p", pass);

                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                // Chép dữ liệu ra biến tạm
                                isLoginSuccess = true;
                                role = reader["Role"].ToString();
                                tabs = reader["AllowedTabs"]?.ToString() ?? "";
                            }
                        }
                    }
                } // ĐẾN ĐÂY: Kết nối kiểm tra Đăng nhập đã được ngắt hoàn toàn, nhường đường cho MainWindow!
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết nối cơ sở dữ liệu: " + ex.Message, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Xử lý chuyển form SAU KHI đã đóng xong Database
            if (isLoginSuccess)
            {
                MainWindow main = new MainWindow(user, role, tabs);
                main.Show();
                this.Close();
            }
            else
            {
                MessageBox.Show("Tài khoản hoặc mật khẩu không chính xác!", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}