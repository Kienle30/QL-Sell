using ClosedXML.Excel;
using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace QuanLyBanHang
{
    public class CheckedItem { public string Name { get; set; } = ""; public bool IsChecked { get; set; } }
    public class CartItem { public int STT { get; set; } public int ProductID { get; set; } public string TenSP { get; set; } = ""; public string DonVi { get; set; } = ""; public int SoLuong { get; set; } public double DonGia { get; set; } public double ThanhTien { get; set; } }
    public class ImportItem { public int STT { get; set; } public string LoaiSP { get; set; } = ""; public string TenSP { get; set; } = ""; public string DonVi { get; set; } = ""; public int SoLuong { get; set; } public double GiaVon { get; set; } public double ThanhTien { get; set; } public int SafeLevel { get; set; } }
    public class ChartItem { public string MonthLabel { get; set; } = ""; public double Value { get; set; } public string ValueString { get; set; } = ""; public double ChartHeight { get; set; } }

    public partial class MainWindow : Window
    {
        private readonly string dbPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DataQL.db");
        private readonly string templateFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");
        private string currentUsername = ""; private string currentUserRole = ""; private string allowedTabs = "";
        private DataTable dtKhoHangToanCuc = new DataTable(); private DataTable dtBanHangToanCuc = new DataTable();
        private DataTable dtThongKeBanCuc = new DataTable(); private DataTable dtThongKeNhapCuc = new DataTable();
        private ObservableCollection<CartItem> gioHangList = new ObservableCollection<CartItem>(); private ObservableCollection<ImportItem> phieuNhapList = new ObservableCollection<ImportItem>();
        private bool isEditMode = false; private int editProductID = -1; private bool isEditUserMode = false; private string editUserID = "";
        private int editSaleID = -1; private int editSaleProductID = -1; private int editSaleOldQty = 0; private bool isUpdatingThongKeCombos = false;
        private string filterCurrentGridName = ""; private string filterCurrentCol = "";

        private Dictionary<string, Dictionary<string, List<string>>> excelFilters = new Dictionary<string, Dictionary<string, List<string>>>() { { "dgKhoHang", new Dictionary<string, List<string>>() }, { "dgBanHang", new Dictionary<string, List<string>>() }, { "dgThongKe", new Dictionary<string, List<string>>() }, { "dgThongKeNhap", new Dictionary<string, List<string>>() } };

        static MainWindow() { var field = typeof(SystemParameters).GetField("_menuDropAlignment", BindingFlags.NonPublic | BindingFlags.Static); if (field != null) { try { field.SetValue(null, false); } catch { } SystemParameters.StaticPropertyChanged += (s, e) => { if (e.PropertyName == "MenuDropAlignment") try { field.SetValue(null, false); } catch { } }; } }

        public MainWindow(string username, string role, string tabs)
        {
            InitializeComponent();
            currentUsername = username ?? "";
            currentUserRole = role ?? "";
            allowedTabs = tabs ?? "";

            txtWelcome.Text = $"Chào {currentUsername} ({currentUserRole})";
            dpNgayBan.SelectedDate = DateTime.Now;

            EnsureDatabaseSetup();

            PhanQuyen(); LoadDuLieuKho(); LoadCombos(); LoadLichSuBanHang(); InitThongKeCombos();
            if (currentUserRole == "Admin") LoadUsers(); dgGioHang.ItemsSource = gioHangList; dgPhieuNhap.ItemsSource = phieuNhapList;
            KiemTraHangSapHet();


            // THÊM DÒNG NÀY VÀO ĐỂ CHẠY BACKUP KHI MỞ APP
            ThucHienAutoBackup();
            KhoiTaoTemplateMacDinh();
        }

        private void EnsureDatabaseSetup()
        {
            try
            {
                string dir = Path.GetDirectoryName(dbPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
                if (!File.Exists(dbPath)) SQLiteConnection.CreateFile(dbPath);

                using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();
                    new SQLiteCommand("CREATE TABLE IF NOT EXISTS Users (UserID INTEGER PRIMARY KEY AUTOINCREMENT, Username TEXT UNIQUE, Password TEXT, Role TEXT)", conn).ExecuteNonQuery();
                    try { new SQLiteCommand("ALTER TABLE Users ADD COLUMN AllowedTabs TEXT DEFAULT 'tabKho,tabBan,tabThongKe,tabHeThong'", conn).ExecuteNonQuery(); } catch { }
                    new SQLiteCommand("INSERT OR IGNORE INTO Users (Username, Password, Role, AllowedTabs) VALUES ('admin', '123', 'Admin', 'tabKho,tabBan,tabThongKe,tabHeThong')", conn).ExecuteNonQuery();

                    new SQLiteCommand("CREATE TABLE IF NOT EXISTS Inventory (ProductID INTEGER PRIMARY KEY AUTOINCREMENT, ProductName TEXT UNIQUE, Category TEXT, Unit TEXT, Quantity INTEGER DEFAULT 0, ImportPrice REAL DEFAULT 0, SafeLevel INTEGER DEFAULT 0, Importer TEXT, ImportDate TEXT)", conn).ExecuteNonQuery();
                    new SQLiteCommand("CREATE TABLE IF NOT EXISTS Sales (SaleID INTEGER PRIMARY KEY AUTOINCREMENT, ProductID INTEGER, QuantitySold INTEGER, SalePrice REAL, SaleDate TEXT, UserID INTEGER, Unit TEXT, Seller TEXT, EntryTime TEXT)", conn).ExecuteNonQuery();
                    try { new SQLiteCommand("ALTER TABLE Sales ADD COLUMN EntryTime TEXT DEFAULT ''", conn).ExecuteNonQuery(); } catch { }

                    new SQLiteCommand("CREATE TABLE IF NOT EXISTS ImportLogs (LogID INTEGER PRIMARY KEY AUTOINCREMENT, ProductID INTEGER, Qty INTEGER, Price REAL, ImportDate TEXT, Importer TEXT, EntryTime TEXT)", conn).ExecuteNonQuery();
                    new SQLiteCommand("CREATE TABLE IF NOT EXISTS DS_Loai (CategoryName TEXT UNIQUE)", conn).ExecuteNonQuery();
                    new SQLiteCommand("CREATE TABLE IF NOT EXISTS DS_SanPham (ProductName TEXT UNIQUE, CategoryName TEXT)", conn).ExecuteNonQuery();
                    new SQLiteCommand("CREATE TABLE IF NOT EXISTS DS_DonVi (UnitName TEXT UNIQUE)", conn).ExecuteNonQuery();

                    if ((long)new SQLiteCommand("SELECT COUNT(*) FROM DS_DonVi", conn).ExecuteScalar() == 0)
                    {
                        new SQLiteCommand("INSERT INTO DS_DonVi VALUES ('Cái'), ('Hộp'), ('Bình'), ('Chiếc')", conn).ExecuteNonQuery();
                    }
                    try { new SQLiteCommand("ALTER TABLE Inventory ADD COLUMN ProductCode TEXT", conn).ExecuteNonQuery(); } catch { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi thiết lập Database: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LockPlacement_DropDownOpened(object sender, EventArgs e) { if (sender is ComboBox cmb && cmb.Template != null) { if (cmb.Template.FindName("PART_Popup", cmb) is System.Windows.Controls.Primitives.Popup popup) { popup.PlacementTarget = cmb; popup.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom; Application.Current.Dispatcher.BeginInvoke(new Action(() => { popup.HorizontalOffset += 0.01; popup.HorizontalOffset -= 0.01; }), System.Windows.Threading.DispatcherPriority.Render); } } }

        private void PhanQuyen()
        {
            if (currentUserRole == "Admin")
            {
                menuKho.Visibility = Visibility.Visible; menuBan.Visibility = Visibility.Visible; menuThongKePanel.Visibility = Visibility.Visible; menuHeThong.Visibility = Visibility.Visible; gridSafeLevel.Visibility = Visibility.Visible; btnXemLog.Visibility = Visibility.Visible;
                btnMoTaoSanPham.Visibility = Visibility.Visible; btnImportExcel.Visibility = Visibility.Visible; btnExportData.Visibility = Visibility.Visible; btnExportTemplate.Visibility = Visibility.Visible;

                // Hiện 2 nút ở tab Kho đối với Admin
                btnSua.Visibility = Visibility.Visible;
                btnXoa.Visibility = Visibility.Visible;
            }
            else
            {
                menuKho.Visibility = allowedTabs.Contains("tabKho") ? Visibility.Visible : Visibility.Collapsed;
                menuBan.Visibility = allowedTabs.Contains("tabBan") ? Visibility.Visible : Visibility.Collapsed;
                menuThongKePanel.Visibility = allowedTabs.Contains("tabThongKe") ? Visibility.Visible : Visibility.Collapsed;
                menuHeThong.Visibility = allowedTabs.Contains("tabHeThong") ? Visibility.Visible : Visibility.Collapsed;
                gridSafeLevel.Visibility = Visibility.Collapsed; btnXemLog.Visibility = Visibility.Collapsed;
                btnMoTaoSanPham.Visibility = Visibility.Collapsed; btnImportExcel.Visibility = Visibility.Collapsed; btnExportData.Visibility = Visibility.Collapsed; btnExportTemplate.Visibility = Visibility.Collapsed;

                // Ẩn 2 nút ở tab Kho đối với Staff
                btnSua.Visibility = Visibility.Collapsed;
                btnXoa.Visibility = Visibility.Collapsed;
            }
            btnXoaBanHang.Visibility = (currentUserRole == "Admin") ? Visibility.Visible : Visibility.Collapsed;
        }

        private bool IsAdmin() { if (currentUserRole == "Admin") return true; MessageBox.Show("⛔ Chỉ Admin mới được thực hiện tính năng này!", "Bảo mật", MessageBoxButton.OK, MessageBoxImage.Warning); return false; }
        private void btnDangXuat_Click(object sender, RoutedEventArgs e) { new LoginWindow().Show(); this.Close(); }
        private double ParseNumber(string text) { if (string.IsNullOrWhiteSpace(text)) return 0; return double.TryParse(text.Replace(",", "").Replace(".", ""), out double res) ? res : 0; }
        private bool isCalculating = false;

        private void TinhTienNhap_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (isCalculating) return; isCalculating = true;
            try
            {
                TextBox? tb = sender as TextBox; if (tb == txtSoLuong || tb == txtThanhTienNhap) { if (!string.IsNullOrEmpty(tb?.Text)) { double parsedRaw = ParseNumber(tb.Text); tb.Text = parsedRaw.ToString("N0"); tb.SelectionStart = tb.Text.Length; } }
                double sl = ParseNumber(txtSoLuong.Text); double thanhTien = ParseNumber(txtThanhTienNhap.Text); if (sl > 0 && thanhTien > 0) txtGiaNhap.Text = (thanhTien / sl).ToString("N0"); else txtGiaNhap.Text = "0";
            }
            catch { }
            isCalculating = false;
        }

        private void dgKhoHang_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e) { if (e.PropertyName == "ProductID") e.Column.Visibility = Visibility.Collapsed; if (e.PropertyName == "Tồn" || e.PropertyName == "Mức an toàn" || e.PropertyName == "Giá Vốn") { if (e.Column is DataGridTextColumn txtCol) txtCol.Binding.StringFormat = "N0"; } }
        private void dgBanHang_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (e.PropertyName == "SaleID" || e.PropertyName == "ProductID")
                e.Column.Visibility = Visibility.Collapsed;

            // Tự động ẩn cột Thời gian tạo nếu không phải là Admin
            if (e.PropertyName == "Thời gian tạo" && currentUserRole != "Admin")
                e.Column.Visibility = Visibility.Collapsed;

            if (e.PropertyName == "Số lượng" || e.PropertyName == "Đơn giá (VNĐ)" || e.PropertyName == "Thành tiền")
            {
                if (e.Column is DataGridTextColumn txtCol) txtCol.Binding.StringFormat = "N0";
            }
        }
        private void dgThongKe_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e) { if (e.PropertyName == "SL Bán" || e.PropertyName == "Giá vốn" || e.PropertyName == "Giá bán" || e.PropertyName == "Doanh thu" || e.PropertyName == "Lãi Lỗ" || e.PropertyName == "Số lượng" || e.PropertyName == "Giá nhập" || e.PropertyName == "Thành tiền") { if (e.Column is DataGridTextColumn txtCol) txtCol.Binding.StringFormat = "N0"; } }

        // ==============================================================
        // BỘ LỌC EXCEL VÀ ĐÁNH SỐ STT THÔNG MINH
        // ==============================================================
        public static T? FindVisualParent<T>(DependencyObject child) where T : DependencyObject { DependencyObject parentObject = VisualTreeHelper.GetParent(child); if (parentObject == null) return null; T? parent = parentObject as T; return parent ?? FindVisualParent<T>(parentObject); }

        private DataTable? GetDataTableForGrid(string gridName)
        {
            if (gridName == "dgKhoHang") return dtKhoHangToanCuc; if (gridName == "dgBanHang") return dtBanHangToanCuc;
            if (gridName == "dgThongKe") return dtThongKeBanCuc; if (gridName == "dgThongKeNhap") return dtThongKeNhapCuc; return null;
        }

        private void btnOpenFilter_Click(object sender, RoutedEventArgs e)
        {
            Button? btn = sender as Button; if (btn == null || btn.Tag == null) return; filterCurrentCol = btn.Tag.ToString() ?? ""; DataGrid? dg = FindVisualParent<DataGrid>(btn); if (dg == null) return; filterCurrentGridName = dg.Name;
            DataTable? dt = GetDataTableForGrid(filterCurrentGridName); if (dt == null) return;
            var uniqueVals = dt.AsEnumerable().Select(r => r[filterCurrentCol]?.ToString() ?? "").Distinct().OrderBy(x => x).ToList(); List<CheckedItem> items = new List<CheckedItem>(); List<string>? allowed = null;
            if (excelFilters.ContainsKey(filterCurrentGridName) && excelFilters[filterCurrentGridName].ContainsKey(filterCurrentCol)) { allowed = excelFilters[filterCurrentGridName][filterCurrentCol]; }
            foreach (var val in uniqueVals) { bool isChk = (allowed == null || allowed.Contains(val)); items.Add(new CheckedItem() { Name = val, IsChecked = isChk }); }
            lbFilterItems.ItemsSource = items; chkSelectAllFilter.IsChecked = items.All(x => x.IsChecked); txtFilterSearch.Text = ""; FilterPopup.PlacementTarget = btn; FilterPopup.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom; FilterPopup.IsOpen = true; Application.Current.Dispatcher.BeginInvoke(new Action(() => { FilterPopup.HorizontalOffset += 0.01; FilterPopup.HorizontalOffset -= 0.01; }), System.Windows.Threading.DispatcherPriority.Render);
        }
        private void txtFilterSearch_TextChanged(object sender, TextChangedEventArgs e) { string search = txtFilterSearch.Text?.ToLower() ?? ""; ICollectionView view = CollectionViewSource.GetDefaultView(lbFilterItems.ItemsSource); if (view != null) { view.Filter = (obj) => { if (obj is CheckedItem item) return item.Name.ToLower().Contains(search); return false; }; } }
        private void chkSelectAllFilter_Click(object sender, RoutedEventArgs e) { bool chk = chkSelectAllFilter.IsChecked == true; if (lbFilterItems.ItemsSource is List<CheckedItem> items) { foreach (var item in items) item.IsChecked = chk; lbFilterItems.Items.Refresh(); } }
        private void btnCancelFilter_Click(object sender, RoutedEventArgs e) { FilterPopup.IsOpen = false; }

        private void btnClearAllFilters_Click(object sender, RoutedEventArgs e)
        {
            string gridName = "";
            if (tabKho.IsSelected) gridName = "dgKhoHang";
            else if (tabBan.IsSelected) gridName = "dgBanHang";
            else if (tabTKBan.IsSelected) gridName = "dgThongKe";
            else if (tabTKNhap.IsSelected) gridName = "dgThongKeNhap";

            if (!string.IsNullOrEmpty(gridName)) { excelFilters[gridName].Clear(); ApplyExcelFilters(gridName); }
        }

        private void btnApplyFilter_Click(object sender, RoutedEventArgs e)
        {
            FilterPopup.IsOpen = false; if (!(lbFilterItems.ItemsSource is List<CheckedItem> items)) return; var allowed = items.Where(x => x.IsChecked).Select(x => x.Name).ToList();
            if (allowed.Count == items.Count) { if (excelFilters[filterCurrentGridName].ContainsKey(filterCurrentCol)) excelFilters[filterCurrentGridName].Remove(filterCurrentCol); } else { excelFilters[filterCurrentGridName][filterCurrentCol] = allowed; }
            ApplyExcelFilters(filterCurrentGridName);
        }

        private void ApplyExcelFilters(string gridName)
        {
            DataTable? dt = GetDataTableForGrid(gridName); if (dt == null) return; List<string> filterParts = new List<string>();
            foreach (var kvp in excelFilters[gridName])
            {
                string col = kvp.Key; List<string> allowed = kvp.Value; if (allowed.Count == 0) { filterParts.Add("1 = 0"); } else { List<string> orClauses = new List<string>(); foreach (var val in allowed) { if (string.IsNullOrEmpty(val)) orClauses.Add($"IsNull([{col}], '') = ''"); else orClauses.Add($"Convert([{col}], 'System.String') = '{val.Replace("'", "''")}'"); } filterParts.Add("(" + string.Join(" OR ", orClauses) + ")"); }
            }
            dt.DefaultView.RowFilter = string.Join(" AND ", filterParts);

            if (dt.Columns.Contains("STT")) { for (int i = 0; i < dt.DefaultView.Count; i++) dt.DefaultView[i]["STT"] = i + 1; }

            if (gridName == "dgKhoHang") TinhTongGiaTriKho();
            else if (gridName == "dgThongKe")
            {
                double tongDoanhThuThang = 0; double tongLaiLoThang = 0; foreach (DataRowView r in dt.DefaultView) { tongDoanhThuThang += Convert.ToDouble(r["Doanh thu"]); tongLaiLoThang += Convert.ToDouble(r["Lãi Lỗ"]); }
                txtTongDoanhThuTK.Text = tongDoanhThuThang.ToString("N0") + " VNĐ"; txtTongLaiLoTK.Text = tongLaiLoThang.ToString("N0") + " VNĐ"; txtTongLaiLoTK.Foreground = tongLaiLoThang >= 0 ? new SolidColorBrush(Color.FromRgb(46, 125, 50)) : new SolidColorBrush(Color.FromRgb(211, 47, 47));
            }
            else if (gridName == "dgThongKeNhap")
            {
                double tongTienNhapThang = 0; foreach (DataRowView r in dt.DefaultView) { tongTienNhapThang += Convert.ToDouble(r["Thành tiền"]); }
                txtTongTienNhapTK.Text = tongTienNhapThang.ToString("N0") + " VNĐ";
            }
        }

        // ==============================================================
        // BÁN HÀNG VÀ XỬ LÝ GIỎ HÀNG 
        // ==============================================================
        private void btnMoTaoDonBan_Click(object sender, RoutedEventArgs e)
        {
            gioHangList.Clear();
            TinhTongDonHang();

            txtSLBan.Text = "";
            txtThanhTien.Text = "";
            txtDonGiaBan.Text = "0";
            txtDonViBan.Text = "";
            cmbLoaiBan.SelectedIndex = 0;
            dpNgayBan.SelectedDate = DateTime.Now;

            OverlayBanHang.Visibility = Visibility.Visible;
        }

        private void btnHuyDonBan_Click(object sender, RoutedEventArgs e)
        {
            if (gioHangList.Count > 0 && MessageBox.Show("Bạn có chắc chắn muốn hủy giỏ hàng đang tạo?", "Cảnh báo", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
                return;

            OverlayBanHang.Visibility = Visibility.Collapsed;
        }

        private void cmbLoaiBan_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (cmbLoaiBan.SelectedItem is DataRowView rv)
                {
                    LoadProductsByCategoryBan(rv["CategoryName"]?.ToString() ?? "");
                }
                else if (cmbLoaiBan.SelectedItem != null)
                {
                    LoadProductsByCategoryBan(cmbLoaiBan.SelectedItem.ToString() ?? "");
                }
            }
            catch { }
        }

        private void cmbLoaiBan_LostFocus(object sender, RoutedEventArgs e)
        {
            try
            {
                LoadProductsByCategoryBan(cmbLoaiBan.Text ?? "");
            }
            catch { }
        }

        private void LoadProductsByCategoryBan(string cat)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                DataTable dt = new DataTable();
                string sql = (string.IsNullOrEmpty(cat) || cat == "Tất cả")
                    ? "SELECT ProductID, ProductName FROM Inventory WHERE Quantity > 0 GROUP BY ProductName"
                    : "SELECT ProductID, ProductName FROM Inventory WHERE Category=@c AND Quantity > 0 GROUP BY ProductName";

                var cmd = new SQLiteCommand(sql, conn);
                if (!string.IsNullOrEmpty(cat) && cat != "Tất cả")
                    cmd.Parameters.AddWithValue("@c", cat);

                new SQLiteDataAdapter(cmd).Fill(dt);
                cmbChonSPBan.ItemsSource = dt.DefaultView;
                cmbChonSPBan.DisplayMemberPath = "ProductName";
                cmbChonSPBan.SelectedValuePath = "ProductID";
            }
        }

        private void cmbChonSPBan_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbChonSPBan.SelectedValue != null)
            {
                int id = Convert.ToInt32(cmbChonSPBan.SelectedValue);
                using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();
                    string? unit = new SQLiteCommand($"SELECT Unit FROM Inventory WHERE ProductID={id}", conn).ExecuteScalar()?.ToString();
                    txtDonViBan.Text = unit ?? "";
                }
            }
        }

        private void TinhTienBan_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (isCalculating) return;
            isCalculating = true;

            try
            {
                TextBox? tb = sender as TextBox;
                if (tb == txtSLBan || tb == txtThanhTien)
                {
                    if (!string.IsNullOrEmpty(tb?.Text))
                    {
                        double parsedRaw = ParseNumber(tb.Text);
                        tb.Text = parsedRaw.ToString("N0");
                        tb.SelectionStart = tb.Text.Length;
                    }
                }

                double sl = ParseNumber(txtSLBan.Text);
                double thanhTien = ParseNumber(txtThanhTien.Text);

                if (sl > 0 && thanhTien > 0)
                    txtDonGiaBan.Text = (thanhTien / sl).ToString("N0");
                else
                    txtDonGiaBan.Text = "0";
            }
            catch { }
            finally { isCalculating = false; }
        }

        private void CapNhatSTTGioHang()
        {
            for (int i = 0; i < gioHangList.Count; i++)
                gioHangList[i].STT = i + 1;

            dgGioHang.Items.Refresh();
        }

        private void btnThemVaoGio_Click(object sender, RoutedEventArgs e)
        {
            if (cmbChonSPBan.SelectedValue == null || string.IsNullOrWhiteSpace(txtSLBan.Text)) { MessageBox.Show("Vui lòng chọn sản phẩm và nhập số lượng!"); return; }
            int id = Convert.ToInt32(cmbChonSPBan.SelectedValue); string tenSP = cmbChonSPBan.Text; int sl = (int)ParseNumber(txtSLBan.Text); double thanhTien = ParseNumber(txtThanhTien.Text); string donVi = txtDonViBan.Text.Trim(); double donGiaTB = (sl > 0) ? (thanhTien / sl) : 0;
            if (sl <= 0) { MessageBox.Show("Số lượng phải lớn hơn 0!"); return; }
            if (thanhTien <= 0) { MessageBox.Show("Vui lòng nhập Thành tiền!"); return; }
            int soLuongDaTrongGio = gioHangList.Where(x => x.ProductID == id).Sum(x => x.SoLuong);
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;")) { conn.Open(); int tonHienTai = Convert.ToInt32(new SQLiteCommand($"SELECT Quantity FROM Inventory WHERE ProductID={id}", conn).ExecuteScalar() ?? 0); if (tonHienTai < (sl + soLuongDaTrongGio)) { MessageBox.Show($"❌ KHÔNG ĐỦ HÀNG TRONG KHO!\nTổng tồn: {tonHienTai.ToString("N0")}\nĐã có trong giỏ: {soLuongDaTrongGio.ToString("N0")}"); return; } }
            gioHangList.Add(new CartItem { STT = gioHangList.Count + 1, ProductID = id, TenSP = tenSP, DonVi = donVi, SoLuong = sl, DonGia = donGiaTB, ThanhTien = thanhTien });
            TinhTongDonHang(); txtSLBan.Text = ""; txtThanhTien.Text = ""; txtDonGiaBan.Text = "0"; txtDonViBan.Text = ""; cmbChonSPBan.SelectedIndex = -1; cmbChonSPBan.Focus();
        }

        private void TinhTongDonHang() { double tong = gioHangList.Sum(x => x.ThanhTien); txtTongTienDonHang.Text = tong.ToString("N0") + " VNĐ"; }
        private void btnXoaKhoiGio_Click(object sender, RoutedEventArgs e) { Button? btn = sender as Button; if (btn != null && btn.DataContext is CartItem item) { gioHangList.Remove(item); CapNhatSTTGioHang(); TinhTongDonHang(); } }

        private void btnLuuDonBan_Click(object sender, RoutedEventArgs e)
        {
            if (gioHangList.Count == 0)
            {
                MessageBox.Show("Giỏ hàng đang trống! Vui lòng thêm sản phẩm trước khi lưu.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string ngayBan = dpNgayBan.SelectedDate.HasValue ? dpNgayBan.SelectedDate.Value.ToString("dd/MM/yyyy") : DateTime.Now.ToString("dd/MM/yyyy");
            string entryTime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var tr = conn.BeginTransaction())
                {
                    try
                    {
                        foreach (var item in gioHangList)
                        {
                            // 1. CẬP NHẬT TỒN KHO & GIÁ VỐN (Logic: Nếu tồn về 0 thì giá vốn về 0)
                            string sqlUpdateInventory = @"
                        UPDATE Inventory 
                        SET Quantity = Quantity - @qty,
                            ImportPrice = CASE 
                                            WHEN (Quantity - @qty) <= 0 THEN 0 
                                            ELSE ImportPrice 
                                          END
                        WHERE ProductID = @pid";

                            using (var cmdUp = new SQLiteCommand(sqlUpdateInventory, conn, tr))
                            {
                                cmdUp.Parameters.AddWithValue("@qty", item.SoLuong);
                                cmdUp.Parameters.AddWithValue("@pid", item.ProductID);
                                cmdUp.ExecuteNonQuery();
                            }

                            // 2. Thêm vào lịch sử bán hàng (Giữ nguyên cấu trúc của bạn)
                            using (var cmdIn = new SQLiteCommand("INSERT INTO Sales (ProductID, QuantitySold, SalePrice, SaleDate, UserID, Unit, Seller, EntryTime) VALUES (@id, @q, @p, @d, 1, @u, @seller, @et)", conn, tr))
                            {
                                cmdIn.Parameters.AddWithValue("@id", item.ProductID);
                                cmdIn.Parameters.AddWithValue("@q", item.SoLuong);
                                cmdIn.Parameters.AddWithValue("@p", item.DonGia);
                                cmdIn.Parameters.AddWithValue("@d", ngayBan);
                                cmdIn.Parameters.AddWithValue("@u", item.DonVi);
                                cmdIn.Parameters.AddWithValue("@seller", currentUsername);
                                cmdIn.Parameters.AddWithValue("@et", entryTime);
                                cmdIn.ExecuteNonQuery();
                            }
                        }
                        tr.Commit();

                        MessageBox.Show("✅ Đã xuất bán thành công toàn bộ đơn hàng!", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);

                        // 3. Xóa sạch giỏ hàng và cập nhật giao diện
                        gioHangList.Clear();
                        txtTongTienDonHang.Text = "0 VNĐ";

                        OverlayBanHang.Visibility = Visibility.Collapsed;
                        LoadDuLieuKho(); // Làm mới bảng kho để thấy giá vốn về 0
                        LoadLichSuBanHang();
                        UpdateThongKeData();
                    }
                    catch (Exception ex)
                    {
                        tr.Rollback();
                        MessageBox.Show("❌ Có lỗi xảy ra trong quá trình lưu.\n" + ex.Message, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private void GhiLogSuaDon(string actionDetail) { try { string logPath = Path.Combine(Path.GetDirectoryName(dbPath) ?? "", "SaleEditLogs.txt"); File.AppendAllText(logPath, $"[{DateTime.Now:dd/MM/yyyy HH:mm:ss}] - User [{currentUsername}] ({currentUserRole}):\n{actionDetail}\n-------------------------------------------------------\n"); } catch { } }
        private void btnXemLog_Click(object sender, RoutedEventArgs e) { if (!IsAdmin()) return; string logPath = Path.Combine(Path.GetDirectoryName(dbPath) ?? "", "SaleEditLogs.txt"); if (File.Exists(logPath)) { System.Diagnostics.Process.Start("notepad.exe", logPath); } else { MessageBox.Show("Chưa có lịch sử sửa đơn nào trong hệ thống!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information); } }

        private void btnMoSuaDonBan_Click(object sender, RoutedEventArgs e)
        {
            if (dgBanHang.SelectedItem is DataRowView r)
            {
                string seller = r["Người bán"].ToString() ?? ""; string saleDateStr = r["Ngày bán"].ToString() ?? ""; string todayStr = DateTime.Now.ToString("dd/MM/yyyy");
                if (currentUserRole != "Admin") { if (seller != currentUsername) { MessageBox.Show("⛔ Bạn chỉ được sửa những đơn hàng do CHÍNH BẠN bán ra!", "Từ chối", MessageBoxButton.OK, MessageBoxImage.Error); return; } if (!saleDateStr.Contains(todayStr)) { MessageBox.Show("⛔ Bạn chỉ được phép sửa đơn hàng của NGÀY HÔM NAY. Vui lòng báo Admin nếu cần sửa đơn cũ!", "Từ chối", MessageBoxButton.OK, MessageBoxImage.Error); return; } }
                editSaleID = Convert.ToInt32(r["SaleID"]); editSaleProductID = Convert.ToInt32(r["ProductID"]); editSaleOldQty = Convert.ToInt32(r["Số lượng"]); txtEditBanSP.Text = r["Tên SP"].ToString(); if (DateTime.TryParseExact(saleDateStr, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime pDate)) dpEditNgayBan.SelectedDate = pDate; else dpEditNgayBan.SelectedDate = DateTime.Now; txtEditBanSL.Text = editSaleOldQty.ToString("N0"); txtEditBanThanhTien.Text = ParseNumber(r["Thành tiền"].ToString() ?? "0").ToString("N0"); dpEditNgayBan.IsEnabled = (currentUserRole == "Admin"); OverlayEditBanHang.Visibility = Visibility.Visible;
            }
            else { MessageBox.Show("Vui lòng click chọn 1 dòng ở bảng để Sửa!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning); }
        }

        private void TinhTienEditBan_TextChanged(object sender, TextChangedEventArgs e) { if (isCalculating) return; isCalculating = true; try { TextBox? tb = sender as TextBox; if (tb == txtEditBanSL || tb == txtEditBanThanhTien) { if (!string.IsNullOrEmpty(tb?.Text)) { double parsedRaw = ParseNumber(tb.Text); tb.Text = parsedRaw.ToString("N0"); tb.SelectionStart = tb.Text.Length; } } double sl = ParseNumber(txtEditBanSL.Text); double thanhTien = ParseNumber(txtEditBanThanhTien.Text); if (sl > 0 && thanhTien > 0) txtEditBanDonGia.Text = (thanhTien / sl).ToString("N0"); else txtEditBanDonGia.Text = "0"; } catch { } isCalculating = false; }
        private void btnHuyEditBan_Click(object sender, RoutedEventArgs e) { OverlayEditBanHang.Visibility = Visibility.Collapsed; }

        private void btnLuuEditBan_Click(object sender, RoutedEventArgs e)
        {
            int newQty = (int)ParseNumber(txtEditBanSL.Text); double newThanhTien = ParseNumber(txtEditBanThanhTien.Text); double newPrice = (newQty > 0) ? (newThanhTien / newQty) : 0; string newDate = dpEditNgayBan.SelectedDate.HasValue ? dpEditNgayBan.SelectedDate.Value.ToString("dd/MM/yyyy") : DateTime.Now.ToString("dd/MM/yyyy");
            if (newQty <= 0 || newThanhTien <= 0) { MessageBox.Show("Vui lòng nhập Số lượng và Thành tiền lớn hơn 0!"); return; }
            int chênhLệchSL = newQty - editSaleOldQty;
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open(); using (var tr = conn.BeginTransaction())
                {
                    try
                    {
                        if (chênhLệchSL > 0) { int tonHienTai = Convert.ToInt32(new SQLiteCommand($"SELECT Quantity FROM Inventory WHERE ProductID={editSaleProductID}", conn).ExecuteScalar() ?? 0); if (tonHienTai < chênhLệchSL) { MessageBox.Show($"❌ KHO KHÔNG ĐỦ ĐỂ TĂNG THÊM!\nTồn hiện tại chỉ còn: {tonHienTai.ToString("N0")}", "Lỗi Kho", MessageBoxButton.OK, MessageBoxImage.Error); return; } }
                        new SQLiteCommand($"UPDATE Inventory SET Quantity = Quantity - {chênhLệchSL} WHERE ProductID = {editSaleProductID}", conn).ExecuteNonQuery(); using (var cmd = new SQLiteCommand("UPDATE Sales SET QuantitySold=@q, SalePrice=@p, SaleDate=@d WHERE SaleID=@id", conn)) { cmd.Parameters.AddWithValue("@q", newQty); cmd.Parameters.AddWithValue("@p", newPrice); cmd.Parameters.AddWithValue("@d", newDate); cmd.Parameters.AddWithValue("@id", editSaleID); cmd.ExecuteNonQuery(); }
                        tr.Commit(); GhiLogSuaDon($"- Sửa [SaleID: {editSaleID}] | Tên SP: {txtEditBanSP.Text}\n- Số lượng: [Cũ: {editSaleOldQty}] -> [Mới: {newQty}]\n- Thành tiền mới: {newThanhTien:N0} VNĐ | Ngày sửa thành: {newDate}"); MessageBox.Show("✅ Cập nhật đơn bán thành công!", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information); OverlayEditBanHang.Visibility = Visibility.Collapsed; LoadDuLieuKho(); LoadLichSuBanHang(); UpdateThongKeData();
                    }
                    catch (Exception ex) { tr.Rollback(); MessageBox.Show("Lỗi: " + ex.Message); }
                }
            }
        }

        private void btnXoaBanHang_Click(object sender, RoutedEventArgs e)
        {
            if (!IsAdmin()) return;
            if (dgBanHang.SelectedItem is DataRowView r)
            {
                if (MessageBox.Show("Xóa giao dịch này?\nSố lượng bán sẽ được HOÀN LẠI vào kho.", "Xác nhận", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                {
                    int saleId = Convert.ToInt32(r["SaleID"]); int productId = Convert.ToInt32(r["ProductID"]); int qty = Convert.ToInt32(r["Số lượng"]);
                    using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;")) { conn.Open(); new SQLiteCommand($"UPDATE Inventory SET Quantity = Quantity + {qty} WHERE ProductID = {productId}", conn).ExecuteNonQuery(); new SQLiteCommand($"DELETE FROM Sales WHERE SaleID = {saleId}", conn).ExecuteNonQuery(); }
                    GhiLogSuaDon($"- XÓA TOÀN BỘ ĐƠN [SaleID: {saleId}]\n- Tên SP: {r["Tên SP"]}\n- Số lượng hoàn kho: {qty}\n- Tổng tiền bị xóa: {r["Thành tiền"]}"); LoadDuLieuKho(); LoadLichSuBanHang(); UpdateThongKeData(); MessageBox.Show("✅ Đã xóa và hoàn kho thành công!");
                }
            }
            else MessageBox.Show("Vui lòng click chọn 1 dòng ở bảng để xóa!");
        }

        private void LoadLichSuBanHang()
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                dtBanHangToanCuc = new DataTable();
                // Thêm S.EntryTime as [Thời gian tạo] vào cuối lệnh SELECT
                new SQLiteDataAdapter(@"SELECT S.SaleID, S.ProductID, S.SaleDate as [Ngày bán], I.Category as [Loại SP], I.ProductName as [Tên SP], S.Unit as [Đơn vị], S.QuantitySold as [Số lượng], S.SalePrice as [Đơn giá (VNĐ)], (S.QuantitySold * S.SalePrice) as [Thành tiền], S.Seller as [Người bán], S.EntryTime as [Thời gian tạo] FROM Sales S INNER JOIN Inventory I ON S.ProductID = I.ProductID ORDER BY S.SaleID DESC LIMIT 300", conn).Fill(dtBanHangToanCuc);

                dtBanHangToanCuc.Columns.Add("STT", typeof(int)).SetOrdinal(0);
                // Vòng lặp gán số thứ tự để cột STT không bị trống
                for (int i = 0; i < dtBanHangToanCuc.Rows.Count; i++) dtBanHangToanCuc.Rows[i]["STT"] = i + 1;

                dgBanHang.ItemsSource = dtBanHangToanCuc.DefaultView;
            }
            ApplyExcelFilters("dgBanHang");
        }

        // ==============================================================
        // TÍNH NĂNG THỐNG KÊ (HỖ TRỢ TẤT CẢ CÁC NĂM)
        // ==============================================================
        private void InitThongKeCombos()
        {
            isUpdatingThongKeCombos = true;
            int currentYear = DateTime.Now.Year;
            cmbNamTK.Items.Add("Tất cả"); cmbNamTKNhap.Items.Add("Tất cả");
            for (int i = currentYear - 3; i <= currentYear + 1; i++) { cmbNamTK.Items.Add(i.ToString()); cmbNamTKNhap.Items.Add(i.ToString()); }

            cmbThangTK.Items.Add("-"); cmbThangTKNhap.Items.Add("-");
            for (int i = 1; i <= 12; i++) { cmbThangTK.Items.Add("Tháng " + i.ToString()); cmbThangTKNhap.Items.Add("Tháng " + i.ToString()); }

            cmbNamTK.SelectedItem = currentYear.ToString(); cmbNamTKNhap.SelectedItem = currentYear.ToString();
            cmbThangTK.SelectedItem = "Tháng " + DateTime.Now.Month.ToString(); cmbThangTKNhap.SelectedItem = "Tháng " + DateTime.Now.Month.ToString();
            isUpdatingThongKeCombos = false; UpdateThongKeData();
        }

        private void cmbThongKe_SelectionChanged(object sender, SelectionChangedEventArgs e) { if (!isUpdatingThongKeCombos) UpdateThongKeData(); }

        private void UpdateThongKeData()
        {
            if (isUpdatingThongKeCombos) return;
            excelFilters["dgThongKe"].Clear(); excelFilters["dgThongKeNhap"].Clear();

            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();

                // --- 1. DỮ LIỆU BÁN HÀNG ---
                string selYearBanStr = cmbNamTK.SelectedItem?.ToString() ?? DateTime.Now.Year.ToString();
                bool isAllYearsBan = selYearBanStr == "Tất cả";
                int selYearBan = isAllYearsBan ? -1 : int.Parse(selYearBanStr);

                string selMonthStrBan = cmbThangTK.SelectedItem?.ToString() ?? "";
                bool isYearBanOnly = selMonthStrBan == "-";
                int selMonthBan = isYearBanOnly ? -1 : int.Parse(selMonthStrBan.Replace("Tháng ", ""));

                txtChartTitle.Text = isAllYearsBan ? "Biểu đồ Doanh Thu Các Năm" : $"Biểu đồ Doanh Thu Cả Năm {selYearBan}";

                DataTable dtRawBan = new DataTable();
                new SQLiteDataAdapter(@"SELECT S.SaleDate as [Ngày bán], I.ProductName as [Sản phẩm], S.QuantitySold as [SL Bán], I.ImportPrice as [Giá vốn], S.SalePrice as [Giá bán], (S.QuantitySold * S.SalePrice) as [Doanh thu], ((S.SalePrice * S.QuantitySold) - (I.ImportPrice * S.QuantitySold)) as [Lãi Lỗ], S.Seller as [Người bán],S.EntryTime as [Thời gian tạo] FROM Sales S INNER JOIN Inventory I ON S.ProductID = I.ProductID ORDER BY S.SaleID DESC", conn).Fill(dtRawBan);

                dtRawBan.Columns.Add("STT", typeof(int)).SetOrdinal(0);
                dtThongKeBanCuc = dtRawBan.Clone();

                double[] monthlyRevenue = new double[12];
                Dictionary<int, double> yearlyRevenue = new Dictionary<int, double>();
                int currentYear = DateTime.Now.Year;
                for (int i = currentYear - 3; i <= currentYear + 1; i++) yearlyRevenue[i] = 0;

                foreach (DataRow row in dtRawBan.Rows)
                {
                    string dateStr = row["Ngày bán"].ToString() ?? "";
                    if (DateTime.TryParseExact(dateStr, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime saleDate))
                    {
                        if (isAllYearsBan || saleDate.Year == selYearBan)
                        {
                            if (isAllYearsBan || isYearBanOnly || saleDate.Month == selMonthBan)
                            {
                                dtThongKeBanCuc.ImportRow(row);
                            }
                            if (isAllYearsBan)
                            {
                                if (yearlyRevenue.ContainsKey(saleDate.Year)) yearlyRevenue[saleDate.Year] += Convert.ToDouble(row["Doanh thu"]);
                            }
                            else
                            {
                                if (saleDate.Year == selYearBan) monthlyRevenue[saleDate.Month - 1] += Convert.ToDouble(row["Doanh thu"]);
                            }
                        }
                    }
                }
                dgThongKe.ItemsSource = dtThongKeBanCuc.DefaultView;

                List<ChartItem> chartData = new List<ChartItem>();
                if (isAllYearsBan)
                {
                    double maxVal = yearlyRevenue.Values.Count > 0 ? yearlyRevenue.Values.Max() : 0; if (maxVal == 0) maxVal = 1;
                    foreach (var kvp in yearlyRevenue)
                    {
                        chartData.Add(new ChartItem { MonthLabel = kvp.Key.ToString(), Value = kvp.Value, ValueString = kvp.Value.ToString("N0") + " đ", ChartHeight = (kvp.Value / maxVal) * 220 });
                    }
                }
                else
                {
                    double maxVal = monthlyRevenue.Max(); if (maxVal == 0) maxVal = 1;
                    for (int i = 0; i < 12; i++)
                    {
                        chartData.Add(new ChartItem { MonthLabel = "T" + (i + 1), Value = monthlyRevenue[i], ValueString = monthlyRevenue[i].ToString("N0") + " đ", ChartHeight = (monthlyRevenue[i] / maxVal) * 220 });
                    }
                }
                icChart.ItemsSource = chartData;

                // --- 2. DỮ LIỆU NHẬP HÀNG ---
                string selYearNhapStr = cmbNamTKNhap.SelectedItem?.ToString() ?? DateTime.Now.Year.ToString();
                bool isAllYearsNhap = selYearNhapStr == "Tất cả";
                int selYearNhap = isAllYearsNhap ? -1 : int.Parse(selYearNhapStr);

                string selMonthStrNhap = cmbThangTKNhap.SelectedItem?.ToString() ?? "";
                bool isYearNhapOnly = selMonthStrNhap == "-";
                int selMonthNhap = isYearNhapOnly ? -1 : int.Parse(selMonthStrNhap.Replace("Tháng ", ""));

                txtChartTitleNhap.Text = isAllYearsNhap ? "Biểu đồ Nhập Hàng Các Năm" : $"Biểu đồ Nhập Hàng Cả Năm {selYearNhap}";

                DataTable dtRawNhap = new DataTable();
                new SQLiteDataAdapter(@"SELECT L.ImportDate as [Ngày Nhập],  I.ProductName as [Sản phẩm], L.Qty as [Số lượng], L.Price as [Giá nhập], (L.Qty * L.Price) as [Thành tiền], L.Importer as [Người nhập], L.EntryTime as [Thời gian tạo] FROM ImportLogs L INNER JOIN Inventory I ON L.ProductID = I.ProductID ORDER BY L.LogID DESC", conn).Fill(dtRawNhap);

                dtRawNhap.Columns.Add("STT", typeof(int)).SetOrdinal(0);
                dtThongKeNhapCuc = dtRawNhap.Clone();

                double[] monthlyImport = new double[12];
                Dictionary<int, double> yearlyImport = new Dictionary<int, double>();
                for (int i = currentYear - 3; i <= currentYear + 1; i++) yearlyImport[i] = 0;

                foreach (DataRow row in dtRawNhap.Rows)
                {
                    string dateStr = row["Thời gian tạo"].ToString() ?? "";
                    if (DateTime.TryParseExact(dateStr, "dd/MM/yyyy HH:mm:ss", null, System.Globalization.DateTimeStyles.None, out DateTime impDate) || DateTime.TryParseExact(dateStr, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out impDate))
                    {
                        if (isAllYearsNhap || impDate.Year == selYearNhap)
                        {
                            if (isAllYearsNhap || isYearNhapOnly || impDate.Month == selMonthNhap)
                            {
                                dtThongKeNhapCuc.ImportRow(row);
                            }
                            if (isAllYearsNhap)
                            {
                                if (yearlyImport.ContainsKey(impDate.Year)) yearlyImport[impDate.Year] += Convert.ToDouble(row["Thành tiền"]);
                            }
                            else
                            {
                                if (impDate.Year == selYearNhap) monthlyImport[impDate.Month - 1] += Convert.ToDouble(row["Thành tiền"]);
                            }
                        }
                    }
                }
                dgThongKeNhap.ItemsSource = dtThongKeNhapCuc.DefaultView;

                List<ChartItem> chartDataNhap = new List<ChartItem>();
                if (isAllYearsNhap)
                {
                    double maxVal = yearlyImport.Values.Count > 0 ? yearlyImport.Values.Max() : 0; if (maxVal == 0) maxVal = 1;
                    foreach (var kvp in yearlyImport)
                    {
                        chartDataNhap.Add(new ChartItem { MonthLabel = kvp.Key.ToString(), Value = kvp.Value, ValueString = kvp.Value.ToString("N0") + " đ", ChartHeight = (kvp.Value / maxVal) * 220 });
                    }
                }
                else
                {
                    double maxVal = monthlyImport.Max(); if (maxVal == 0) maxVal = 1;
                    for (int i = 0; i < 12; i++)
                    {
                        chartDataNhap.Add(new ChartItem { MonthLabel = "T" + (i + 1), Value = monthlyImport[i], ValueString = monthlyImport[i].ToString("N0") + " đ", ChartHeight = (monthlyImport[i] / maxVal) * 220 });
                    }
                }
                icChartNhap.ItemsSource = chartDataNhap;
            }

            ApplyExcelFilters("dgThongKe");
            ApplyExcelFilters("dgThongKeNhap");
        }

        private void btnExportThongKeBan_Click(object sender, RoutedEventArgs e)
        {
            if (!IsAdmin() || dtThongKeBanCuc == null || dtThongKeBanCuc.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Logic lấy tên tháng năm của bồ (Rất hay!)
            string monthPart = cmbThangTK.Text == "-" ? "CaNam" : cmbThangTK.Text.Replace(" ", "");
            string yearPart = cmbNamTK.Text == "Tất cả" ? "TatCaCacNam" : cmbNamTK.Text;

            // Lấy dữ liệu đang hiển thị (đã áp dụng bộ lọc)
            DataTable dtXuat = dtThongKeBanCuc.DefaultView.ToTable();

            // Gọi hàm xuất vào form mẫu
            XuatBaoCaoRaForm(dtXuat, "Template_ThongKeBan.xlsx", $"DoanhThu_{monthPart}_{yearPart}");
        }

        private void btnExportThongKeNhap_Click(object sender, RoutedEventArgs e)
        {
            if (!IsAdmin() || dtThongKeNhapCuc == null || dtThongKeNhapCuc.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string monthPart = cmbThangTKNhap.Text == "-" ? "CaNam" : cmbThangTKNhap.Text.Replace(" ", "");
            string yearPart = cmbNamTKNhap.Text == "Tất cả" ? "TatCaCacNam" : cmbNamTKNhap.Text;

            // Lấy dữ liệu đang hiển thị
            DataTable dtXuat = dtThongKeNhapCuc.DefaultView.ToTable();

            // Trỏ vào form mẫu Thống kê Nhập
            XuatBaoCaoRaForm(dtXuat, "Template_ThongKeNhap.xlsx", $"LichSuNhap_{monthPart}_{yearPart}");
        }

        // ==============================================================
        // TOÀN BỘ PHẦN TẠO SẢN PHẨM & NHẬP KHO
        // ==============================================================

        // ----- HÀM NHẬP DỮ LIỆU TỪ POPUP NHỎ -----
        private string ShowInput(string t, string d = "")
        {
            Window w = new Window() { Width = 350, Height = 160, Title = t, WindowStartupLocation = WindowStartupLocation.CenterScreen, WindowStyle = WindowStyle.ToolWindow };
            StackPanel s = new StackPanel() { Margin = new Thickness(15) };
            TextBox x = new TextBox() { Height = 35, Margin = new Thickness(0, 10, 0, 15), Text = d };
            Button b = new Button() { Content = "LƯU LẠI", IsDefault = true, Height = 35, Background = new SolidColorBrush(Color.FromRgb(0, 123, 255)), Foreground = Brushes.White, FontWeight = FontWeights.Bold };
            b.Click += (se, ev) => { w.DialogResult = true; w.Close(); };
            s.Children.Add(new TextBlock() { Text = t, FontWeight = FontWeights.Bold });
            s.Children.Add(x);
            s.Children.Add(b);
            w.Content = s;
            x.Focus();
            return w.ShowDialog() == true ? x.Text.Trim() : "";
        }

        // ----- LUỒNG TẠO SẢN PHẨM MỚI -----
        private void btnMoTaoSanPham_Click(object sender, RoutedEventArgs e)
        {
            if (!IsAdmin()) return;

            // Reset các ô nhập
            cmbNewLoai.Text = "";
            txtNewTenSP.Text = "";
            cmbNewDonVi.Text = "";
            txtNewSafeLevel.Text = "0";
            txtNewMaSP.Text = ""; // Để trống mã cho đến khi chọn loại hàng

            OverlayTaoSanPham.Visibility = Visibility.Visible;
        }

        private void btnHuyTaoSanPham_Click(object sender, RoutedEventArgs e) { OverlayTaoSanPham.Visibility = Visibility.Collapsed; }

        private void btnAddCategory_Click(object sender, RoutedEventArgs e)
        {
            if (IsAdmin())
            {
                string s = ShowInput("Loại Sản Phẩm mới:");
                if (!string.IsNullOrEmpty(s))
                {
                    RunSQL("INSERT OR IGNORE INTO DS_Loai VALUES (@c)", c => c.Parameters.AddWithValue("@c", s));
                    LoadCombos();
                    cmbNewLoai.Text = s;
                }
            }
        }

        private void btnRemoveCategory_Click(object sender, RoutedEventArgs e)
        {
            if (IsAdmin())
            {
                if (MessageBox.Show("Xóa Loại SP này?", "Xác nhận", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    RunSQL("DELETE FROM DS_Loai WHERE CategoryName=@c", c => c.Parameters.AddWithValue("@c", cmbNewLoai.Text));
                    LoadCombos();
                    cmbNewLoai.Text = "";
                }
            }
        }

        private void btnAddDonViKho_Click(object sender, RoutedEventArgs e)
        {
            if (IsAdmin())
            {
                string s = ShowInput("Đơn vị tính mới:");
                if (!string.IsNullOrEmpty(s))
                {
                    RunSQL("INSERT OR IGNORE INTO DS_DonVi VALUES (@v)", c => c.Parameters.AddWithValue("@v", s));
                    LoadCombos();
                    cmbNewDonVi.Text = s;
                }
            }
        }

        private void btnRemoveDonViKho_Click(object sender, RoutedEventArgs e)
        {
            if (IsAdmin())
            {
                if (MessageBox.Show("Xóa Đơn vị tính này?", "Xác nhận", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    RunSQL("DELETE FROM DS_DonVi WHERE UnitName=@v", c => c.Parameters.AddWithValue("@v", cmbNewDonVi.Text));
                    LoadCombos();
                    cmbNewDonVi.Text = "";
                }
            }
        }

        private void btnLuuSanPhamMoi_Click(object sender, RoutedEventArgs e)
        {
            // 1. Lấy dữ liệu từ giao diện
            string maSP = txtNewMaSP.Text; // Mã tự động đã sinh ra ở ô mới thêm
            string loai = cmbNewLoai.Text.Trim();
            string ten = txtNewTenSP.Text.Trim();
            string donVi = cmbNewDonVi.Text.Trim();
            int safeLvl = (int)ParseNumber(txtNewSafeLevel.Text);

            // 2. Kiểm tra bỏ trống
            if (string.IsNullOrEmpty(loai) || string.IsNullOrEmpty(ten) || string.IsNullOrEmpty(donVi))
            {
                MessageBox.Show("Vui lòng nhập đủ các trường Loại SP, Tên SP và Đơn vị!", "Cảnh báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();

                // 3. Kiểm tra trùng tên hoặc trùng mã sản phẩm
                string sqlCheck = "SELECT COUNT(*) FROM Inventory WHERE ProductName = @name OR ProductCode = @code";
                using (var cmdCheck = new SQLiteCommand(sqlCheck, conn))
                {
                    cmdCheck.Parameters.AddWithValue("@name", ten);
                    cmdCheck.Parameters.AddWithValue("@code", maSP);
                    int count = Convert.ToInt32(cmdCheck.ExecuteScalar());
                    if (count > 0)
                    {
                        MessageBox.Show("Sản phẩm hoặc Mã này đã tồn tại trong hệ thống!", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }

                // 4. Cập nhật các bảng danh mục phụ (Dùng Parameter để chống lỗi dấu nháy)
                using (var cmdRef = new SQLiteCommand("INSERT OR IGNORE INTO DS_Loai VALUES (@l)", conn)) { cmdRef.Parameters.AddWithValue("@l", loai); cmdRef.ExecuteNonQuery(); }
                using (var cmdRef = new SQLiteCommand("INSERT OR IGNORE INTO DS_SanPham (ProductName, CategoryName) VALUES (@n, @l)", conn)) { cmdRef.Parameters.AddWithValue("@n", ten); cmdRef.Parameters.AddWithValue("@l", loai); cmdRef.ExecuteNonQuery(); }
                using (var cmdRef = new SQLiteCommand("INSERT OR IGNORE INTO DS_DonVi VALUES (@d)", conn)) { cmdRef.Parameters.AddWithValue("@d", donVi); cmdRef.ExecuteNonQuery(); }

                // 5. LƯU VÀO BẢNG CHÍNH (Đã thêm ProductCode)
                string now = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                string sqlInsert = @"INSERT INTO Inventory (ProductCode, ProductName, Category, Unit, Quantity, ImportPrice, SafeLevel, Importer, ImportDate) 
                             VALUES (@code, @name, @cat, @unit, 0, 0, @safe, @user, @date)";

                using (var cmd = new SQLiteCommand(sqlInsert, conn))
                {
                    cmd.Parameters.AddWithValue("@code", maSP); // Lưu mã SP tự động
                    cmd.Parameters.AddWithValue("@name", ten);
                    cmd.Parameters.AddWithValue("@cat", loai);
                    cmd.Parameters.AddWithValue("@unit", donVi);
                    cmd.Parameters.AddWithValue("@safe", safeLvl);
                    cmd.Parameters.AddWithValue("@user", currentUsername);
                    cmd.Parameters.AddWithValue("@date", now);

                    cmd.ExecuteNonQuery();
                }
            }

            // 6. Hoàn tất và dọn dẹp
            MessageBox.Show("✅ Đã tạo Danh mục sản phẩm mới thành công!", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
            OverlayTaoSanPham.Visibility = Visibility.Collapsed;

            // Reset các ô nhập liệu cho lần sau
            txtNewTenSP.Text = "";
            txtNewMaSP.Text = "";

            LoadCombos();
            LoadDuLieuKho();
        }

        // ----- LUỒNG QUẢN LÝ KHO -----
        private void LoadDuLieuKho()
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                dtKhoHangToanCuc = new DataTable();

                // THÊM: ProductCode as [Mã SP] vào ngay sau ProductID
                string sql = @"SELECT ProductID, 
                       ProductCode as [Mã SP], 
                       Category as [Loại SP], 
                       ProductName as [Tên SP], 
                       Unit as [Đơn vị], 
                       Quantity as [Tồn], 
                       SafeLevel as [Mức an toàn], 
                       CASE 
                            WHEN Quantity <= 0 THEN '❌ HẾT' 
                            WHEN Quantity < SafeLevel THEN '⚠️ SẮP HẾT' 
                            ELSE '✅ OK' 
                       END as [Trạng thái], 
                       ImportPrice as [Giá Vốn] 
                       FROM Inventory ORDER BY ProductID DESC";

                new SQLiteDataAdapter(sql, conn).Fill(dtKhoHangToanCuc);

                // Cấp số thứ tự (STT) cho bảng
                if (!dtKhoHangToanCuc.Columns.Contains("STT"))
                {
                    dtKhoHangToanCuc.Columns.Add("STT", typeof(int)).SetOrdinal(0);
                }
                for (int i = 0; i < dtKhoHangToanCuc.Rows.Count; i++)
                {
                    dtKhoHangToanCuc.Rows[i]["STT"] = i + 1;
                }

                dgKhoHang.ItemsSource = dtKhoHangToanCuc.DefaultView;
            }
            ApplyExcelFilters("dgKhoHang");
        }

        private void TinhTongGiaTriKho() { if (dtKhoHangToanCuc == null || txtTongGiaTriKho == null) return; double tong = 0; foreach (DataRowView r in dtKhoHangToanCuc.DefaultView) { double ton = ParseNumber(r["Tồn"].ToString() ?? "0"); double gia = ParseNumber(r["Giá Vốn"].ToString() ?? "0"); if (ton > 0) tong += (ton * gia); } txtTongGiaTriKho.Text = tong.ToString("N0") + " VNĐ"; }

        private void btnHuyPopup_Click(object sender, RoutedEventArgs e) { if (!isEditMode && phieuNhapList.Count > 0 && MessageBox.Show("Hủy phiếu nhập đang tạo?", "Cảnh báo", MessageBoxButton.YesNo) == MessageBoxResult.No) return; OverlayKho.Visibility = Visibility.Collapsed; }

        private void btnThem_Click(object sender, RoutedEventArgs e)
        {


            // 1. Đặt mặc định ô Loại SP là "Tất cả"
            cmbLoai.Text = "Tất cả";

            // 2. Reset các ô nhập liệu khác về trống
            cmbTenSP.Text = "";
            txtDonViKho.Text = "";
            txtSoLuong.Text = "";
            txtThanhTienNhap.Text = "";
            txtGiaNhap.Text = "0";
            txtSafeLevel.Text = "0";
            dpNgayNhapKho.SelectedDate = DateTime.Now;

            // 3. Gọi hàm tải toàn bộ sản phẩm ngay lập tức
            LoadProductsByCategory("Tất cả");

            // Khóa ô Mức an toàn (chỉ mở khi Sửa dòng)
            txtSafeLevel.IsReadOnly = true;
            txtSafeLevel.Background = (System.Windows.Media.SolidColorBrush)new System.Windows.Media.BrushConverter().ConvertFrom("#E9ECEF");

            txtPopupTitle.Text = "📥 TẠO PHIẾU NHẬP KHO";
            btnLuuPhieuNhap.Content = "💾 LƯU PHIẾU NHẬP";
            isEditMode = false;

            dgPhieuNhap.Visibility = Visibility.Visible;
            btnThemVaoPhieu.Visibility = Visibility.Visible;
            pnlTongTienNhap.Visibility = Visibility.Visible;

            OverlayKho.Visibility = Visibility.Visible;
        }

        private void btnSua_Click(object sender, RoutedEventArgs e)
        {
            if (!IsAdmin()) return;
            if (dgKhoHang.SelectedItem is DataRowView r)
            {
                isEditMode = true;
                editProductID = Convert.ToInt32(r["ProductID"]);
                txtPopupTitle.Text = "📝 CẬP THÔNG TIN NHẬP KHO";
                btnLuuPhieuNhap.Content = "💾 CẬP NHẬT DÒNG";
                cmbLoai.Text = r["Loại SP"].ToString();
                LoadProductsByCategory(cmbLoai.Text);
                cmbTenSP.Text = r["Tên SP"].ToString();
                txtDonViKho.Text = r["Đơn vị"].ToString();
                double ton = ParseNumber(r["Tồn"].ToString() ?? "0");
                double giaVon = ParseNumber(r["Giá Vốn"].ToString() ?? "0");
                txtSoLuong.Text = ton.ToString("N0");
                txtThanhTienNhap.Text = (ton * giaVon).ToString("N0");
                txtSafeLevel.Text = r["Mức an toàn"].ToString();

                using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();
                    string dStr = new SQLiteCommand($"SELECT ImportDate FROM Inventory WHERE ProductID = {editProductID}", conn).ExecuteScalar()?.ToString() ?? "";
                    if (DateTime.TryParseExact(dStr, "dd/MM/yyyy HH:mm:ss", null, System.Globalization.DateTimeStyles.None, out DateTime pDate) || DateTime.TryParseExact(dStr, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out pDate))
                        dpNgayNhapKho.SelectedDate = pDate;
                    else
                        dpNgayNhapKho.SelectedDate = DateTime.Now;
                }

                txtSafeLevel.IsReadOnly = false;
                txtSafeLevel.Background = System.Windows.Media.Brushes.White;

                dgPhieuNhap.Visibility = Visibility.Collapsed;
                btnThemVaoPhieu.Visibility = Visibility.Collapsed;
                pnlTongTienNhap.Visibility = Visibility.Collapsed;
                OverlayKho.Visibility = Visibility.Visible;
            }
            else
            {
                MessageBox.Show("Chọn 1 dòng trong bảng kho để sửa!");
            }
        }

        private void AutoFillProductInfo()
        {
            string tenSP = cmbTenSP.Text.Trim();
            if (string.IsNullOrEmpty(tenSP)) { txtDonViKho.Text = ""; txtSafeLevel.Text = ""; return; }
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                var cmd = new SQLiteCommand("SELECT Category, Unit, SafeLevel FROM Inventory WHERE ProductName=@n LIMIT 1", conn);
                cmd.Parameters.AddWithValue("@n", tenSP);
                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        txtDonViKho.Text = reader["Unit"].ToString();
                        txtSafeLevel.Text = reader["SafeLevel"].ToString();
                        string cat = reader["Category"].ToString();
                        if (!string.IsNullOrEmpty(cat)) cmbLoai.Text = cat;
                    }
                    else
                    {
                        txtDonViKho.Text = ""; txtSafeLevel.Text = "";
                    }
                }
            }
        }

        private void cmbLoaiKho_SelectionChanged(object sender, SelectionChangedEventArgs e) { try { if (cmbLoai.SelectedItem is DataRowView r) LoadProductsByCategory(r["CategoryName"].ToString()); else if (cmbLoai.SelectedItem != null) LoadProductsByCategory(cmbLoai.SelectedItem.ToString() ?? ""); AutoFillProductInfo(); } catch { } }
        private void cmbLoaiKho_LostFocus(object sender, RoutedEventArgs e) { try { LoadProductsByCategory(cmbLoai.Text); AutoFillProductInfo(); } catch { } }

        private void CapNhatSTTPhieuNhap() { for (int i = 0; i < phieuNhapList.Count; i++) phieuNhapList[i].STT = i + 1; dgPhieuNhap.Items.Refresh(); }

        private void btnThemVaoPhieu_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(cmbTenSP.Text) || string.IsNullOrWhiteSpace(txtSoLuong.Text) || string.IsNullOrWhiteSpace(txtThanhTienNhap.Text)) { MessageBox.Show("Vui lòng nhập Tên Sản Phẩm, Số lượng và Thành tiền!", "Cảnh báo", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
            if (string.IsNullOrWhiteSpace(txtDonViKho.Text)) { MessageBox.Show("Sản phẩm này chưa có trong hệ thống!\nVui lòng sử dụng nút [TẠO SP MỚI] để khai báo danh mục trước khi nhập kho.", "Từ chối", MessageBoxButton.OK, MessageBoxImage.Error); return; }

            string tenSP = cmbTenSP.Text.Trim(); string loai = cmbLoai.Text.Trim(); string donVi = txtDonViKho.Text.Trim(); int slMoi = (int)ParseNumber(txtSoLuong.Text); double thanhTienMoi = ParseNumber(txtThanhTienNhap.Text); double giaMoi = ParseNumber(txtGiaNhap.Text); int safeLvl = (int)ParseNumber(txtSafeLevel.Text);

            var existingItem = phieuNhapList.FirstOrDefault(x => x.TenSP == tenSP && x.LoaiSP == loai);
            if (existingItem != null) { existingItem.SoLuong += slMoi; existingItem.ThanhTien += thanhTienMoi; existingItem.GiaVon = existingItem.ThanhTien / existingItem.SoLuong; dgPhieuNhap.Items.Refresh(); } else { phieuNhapList.Add(new ImportItem { STT = phieuNhapList.Count + 1, LoaiSP = loai, TenSP = tenSP, DonVi = donVi, SoLuong = slMoi, GiaVon = giaMoi, ThanhTien = thanhTienMoi, SafeLevel = safeLvl }); }
            CapNhatSTTPhieuNhap(); TinhTongPhieuNhap(); txtSoLuong.Text = ""; txtThanhTienNhap.Text = ""; txtGiaNhap.Text = "0"; cmbTenSP.Focus();
        }

        private void TinhTongPhieuNhap() { double tong = phieuNhapList.Sum(x => x.ThanhTien); txtTongTienNhap.Text = tong.ToString("N0") + " VNĐ"; }
        private void btnXoaKhoiPhieuNhap_Click(object sender, RoutedEventArgs e) { Button? btn = sender as Button; if (btn != null && btn.DataContext is ImportItem item) { phieuNhapList.Remove(item); CapNhatSTTPhieuNhap(); TinhTongPhieuNhap(); } }

        private void btnLuuPopup_Click(object sender, RoutedEventArgs e)
        {
            string customDate = dpNgayNhapKho.SelectedDate.HasValue ? dpNgayNhapKho.SelectedDate.Value.ToString("dd/MM/yyyy") : DateTime.Now.ToString("dd/MM/yyyy");
            string entryTime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                if (isEditMode)
                {
                    if (string.IsNullOrWhiteSpace(cmbTenSP.Text) || string.IsNullOrWhiteSpace(txtDonViKho.Text) || string.IsNullOrWhiteSpace(txtSoLuong.Text) || string.IsNullOrWhiteSpace(txtThanhTienNhap.Text))
                    {
                        MessageBox.Show("Nhập đủ các trường (*)", "Cảnh báo"); return;
                    }

                    int safeLvl = string.IsNullOrWhiteSpace(txtSafeLevel.Text) ? 0 : (int)ParseNumber(txtSafeLevel.Text);
                    int qty = (int)ParseNumber(txtSoLuong.Text);
                    double price = ParseNumber(txtGiaNhap.Text);

                    // LOGIC: Nếu tồn về 0 thì giá vốn ép về 0
                    if (qty <= 0) price = 0;

                    using (var cmd = new System.Data.SQLite.SQLiteCommand("UPDATE Inventory SET Quantity=@q, ImportPrice=@p, SafeLevel=@s, Importer=@u, ImportDate=@d WHERE ProductID=@id", conn))
                    {
                        cmd.Parameters.AddWithValue("@q", qty);
                        cmd.Parameters.AddWithValue("@p", price);
                        cmd.Parameters.AddWithValue("@s", safeLvl);
                        cmd.Parameters.AddWithValue("@u", currentUsername);
                        cmd.Parameters.AddWithValue("@d", customDate);
                        cmd.Parameters.AddWithValue("@id", editProductID);
                        cmd.ExecuteNonQuery();
                    }
                    MessageBox.Show("✅ Đã cập nhật thành công!");
                }
                else
                {
                    if (phieuNhapList.Count == 0) { MessageBox.Show("Phiếu nhập đang trống!", "Cảnh báo", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
                    using (var tr = conn.BeginTransaction())
                    {
                        try
                        {
                            foreach (var item in phieuNhapList)
                            {
                                int oldID = -1; int slCu = 0; double giaCu = 0;
                                using (var cmdCheck = new System.Data.SQLite.SQLiteCommand("SELECT ProductID, Quantity, ImportPrice FROM Inventory WHERE ProductName=@n", conn))
                                {
                                    cmdCheck.Parameters.AddWithValue("@n", item.TenSP);
                                    using (var rd = cmdCheck.ExecuteReader())
                                    {
                                        if (rd.Read()) { oldID = Convert.ToInt32(rd["ProductID"]); slCu = Convert.ToInt32(rd["Quantity"]); giaCu = Convert.ToDouble(rd["ImportPrice"]); }
                                    }
                                }

                                if (oldID != -1)
                                {
                                    int tongSL = slCu + item.SoLuong;
                                    double giaTB = tongSL > 0 ? ((giaCu * slCu) + (item.GiaVon * item.SoLuong)) / tongSL : 0;

                                    // Check lại một lần nữa cho chắc chắn nếu tổng tồn <= 0
                                    if (tongSL <= 0) giaTB = 0;

                                    using (var cmd = new System.Data.SQLite.SQLiteCommand("UPDATE Inventory SET Quantity=@q, ImportPrice=@p, ImportDate=@d, Importer=@u WHERE ProductID=@id", conn))
                                    {
                                        cmd.Parameters.AddWithValue("@q", tongSL);
                                        cmd.Parameters.AddWithValue("@p", Math.Round(giaTB, 2));
                                        cmd.Parameters.AddWithValue("@d", customDate);
                                        cmd.Parameters.AddWithValue("@u", currentUsername);
                                        cmd.Parameters.AddWithValue("@id", oldID);
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                                else
                                {
                                    using (var cmd = new System.Data.SQLite.SQLiteCommand("INSERT INTO Inventory (ProductName,Category,Unit,Quantity,ImportPrice,SafeLevel,Importer,ImportDate) VALUES (@n,@c,@unit,@q,@p,@s,@u,@d)", conn))
                                    {
                                        cmd.Parameters.AddWithValue("@n", item.TenSP);
                                        cmd.Parameters.AddWithValue("@c", item.LoaiSP);
                                        cmd.Parameters.AddWithValue("@unit", item.DonVi);
                                        cmd.Parameters.AddWithValue("@q", item.SoLuong);
                                        cmd.Parameters.AddWithValue("@p", item.GiaVon);
                                        cmd.Parameters.AddWithValue("@s", item.SafeLevel);
                                        cmd.Parameters.AddWithValue("@u", currentUsername);
                                        cmd.Parameters.AddWithValue("@d", customDate);
                                        cmd.ExecuteNonQuery();
                                        oldID = (int)conn.LastInsertRowId;
                                    }
                                }
                                using (var logCmd = new System.Data.SQLite.SQLiteCommand("INSERT INTO ImportLogs (ProductID, Qty, Price, ImportDate, Importer, EntryTime) VALUES (@pid, @q, @p, @d, @u, @et)", conn))
                                {
                                    logCmd.Parameters.AddWithValue("@pid", oldID);
                                    logCmd.Parameters.AddWithValue("@q", item.SoLuong);
                                    logCmd.Parameters.AddWithValue("@p", item.GiaVon);
                                    logCmd.Parameters.AddWithValue("@d", customDate);
                                    logCmd.Parameters.AddWithValue("@u", currentUsername);
                                    logCmd.Parameters.AddWithValue("@et", entryTime);
                                    logCmd.ExecuteNonQuery();
                                }
                            }
                            tr.Commit();
                            MessageBox.Show("✅ Đã nhập kho thành công toàn bộ phiếu!", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);

                            // Xóa sạch danh sách phiếu nhập sau khi lưu thành công
                            phieuNhapList.Clear();
                        }
                        catch (Exception ex) { tr.Rollback(); MessageBox.Show("Lỗi trong quá trình lưu: " + ex.Message, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error); return; }
                    }
                }
            }

            // XÓA SẠCH CÁC Ô NHẬP LIỆU TRÊN GIAO DIỆN
            cmbTenSP.SelectedIndex = -1;
            cmbTenSP.Text = "";
            txtDonViKho.Clear();
            txtSoLuong.Clear();
            txtGiaNhap.Clear();
            txtSafeLevel.Clear();
            txtThanhTienNhap.Text = "0";

            OverlayKho.Visibility = Visibility.Collapsed;
            LoadDuLieuKho();
            LoadCombos();
            UpdateThongKeData();
        }

        private void btnXoa_Click(object sender, RoutedEventArgs e) { if (!IsAdmin()) return; if (dgKhoHang.SelectedItem is DataRowView r) { if (MessageBox.Show("Xóa sản phẩm này?", "Cảnh báo", MessageBoxButton.YesNo) == MessageBoxResult.Yes) { RunSQL("DELETE FROM Inventory WHERE ProductID=@id", c => c.Parameters.AddWithValue("@id", r["ProductID"])); LoadDuLieuKho(); } } }

        private void LoadCombos()
        {
            try
            {
                using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();

                    // 1. Tải danh sách Loại Sản Phẩm
                    System.Data.DataTable dtLoai = new System.Data.DataTable();
                    new System.Data.SQLite.SQLiteDataAdapter("SELECT CategoryName FROM DS_Loai ORDER BY CategoryName", conn).Fill(dtLoai);

                    System.Collections.Generic.List<string> listLoai = new System.Collections.Generic.List<string>();
                    listLoai.Add("Tất cả"); // Thêm tùy chọn "Tất cả" lên đầu

                    foreach (System.Data.DataRow row in dtLoai.Rows)
                    {
                        listLoai.Add(row["CategoryName"].ToString());
                    }

                    cmbLoai.ItemsSource = listLoai;
                    cmbLoaiBan.ItemsSource = listLoai;
                    cmbNewLoai.ItemsSource = listLoai;

                    // 2. Tải danh sách Đơn Vị Tính
                    System.Data.DataTable dtDonVi = new System.Data.DataTable();
                    new System.Data.SQLite.SQLiteDataAdapter("SELECT UnitName FROM DS_DonVi ORDER BY UnitName", conn).Fill(dtDonVi);

                    System.Collections.Generic.List<string> listDonVi = new System.Collections.Generic.List<string>();
                    foreach (System.Data.DataRow row in dtDonVi.Rows)
                    {
                        listDonVi.Add(row["UnitName"].ToString());
                    }

                    cmbNewDonVi.ItemsSource = listDonVi;
                }
            }
            catch { }
        }

        private void LoadProductsByCategory(string cat)
        {
            try
            {
                using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();
                    System.Data.DataTable dt = new System.Data.DataTable();

                    // Nếu cat rỗng HOẶC cat là "Tất cả", SELECT tất cả sản phẩm
                    string sql = (string.IsNullOrWhiteSpace(cat) || cat == "Tất cả")
                        ? "SELECT ProductID, ProductName FROM Inventory GROUP BY ProductName"
                        : "SELECT ProductID, ProductName FROM Inventory WHERE Category=@c GROUP BY ProductName";

                    var cmd = new System.Data.SQLite.SQLiteCommand(sql, conn);

                    // Chỉ truyền tham số khi có chọn loại cụ thể
                    if (!string.IsNullOrWhiteSpace(cat) && cat != "Tất cả")
                    {
                        cmd.Parameters.AddWithValue("@c", cat);
                    }

                    new System.Data.SQLite.SQLiteDataAdapter(cmd).Fill(dt);
                    cmbTenSP.ItemsSource = dt.DefaultView;
                    cmbTenSP.DisplayMemberPath = "ProductName";
                    cmbTenSP.SelectedValuePath = "ProductID";
                }
            }
            catch { }
        }

        // 1. NÚT XUẤT BÁO CÁO KHO HÀNG (Sử dụng hàm đa năng)
        private void btnExportData_Click(object sender, RoutedEventArgs e)
        {
            if (!IsAdmin() || dtKhoHangToanCuc == null) return;

            // DefaultView.ToTable() giúp xuất chính xác những gì đang hiển thị trên màn hình (đã lọc/tìm kiếm)
            DataTable dtXuat = dtKhoHangToanCuc.DefaultView.ToTable();

            // Xóa cột ProductID đi để không in ra báo cáo làm rối mắt
            if (dtXuat.Columns.Contains("ProductID")) dtXuat.Columns.Remove("ProductID");

            // Gọi hàm xuất form siêu xịn đã viết ở bước trước
            XuatBaoCaoRaForm(dtXuat, "Template_Kho.xlsx", "Bao_Cao_Kho");
        }

        // 2. NÚT XUẤT FORM MẪU ĐỂ NHẬP LIỆU (Đồng bộ Excel)
        private void btnExportTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (!IsAdmin()) return;

            SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook (*.xlsx)|*.xlsx", FileName = "Form_Mau_Dong_Bo" };
            if (sfd.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var ws = workbook.Worksheets.Add("Template");

                        // Tiêu đề cột
                        string[] headers = { "Loại SP", "Tên SP", "Đơn vị", "Số lượng", "Thành tiền", "Ngày nhập" };
                        for (int i = 0; i < headers.Length; i++)
                        {
                            var cell = ws.Cell(1, i + 1);
                            cell.Value = headers[i];
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#28A745"); // Nền xanh lá cây (chuẩn form nhập)
                            cell.Style.Font.FontColor = XLColor.White;
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        }

                        // Dòng ví dụ mẫu (người dùng sẽ dựa vào đây để điền)
                        ws.Cell(2, 1).Value = "Điện tử";
                        ws.Cell(2, 2).Value = "Laptop Dell";
                        ws.Cell(2, 3).Value = "Cái";
                        ws.Cell(2, 4).Value = 10;
                        ws.Cell(2, 5).Value = 15000000;
                        ws.Cell(2, 6).Value = DateTime.Now.ToString("dd/MM/yyyy");

                        // Định dạng In nghiêng và bôi màu xám cho dòng ví dụ để dễ phân biệt
                        var exampleRow = ws.Range("A2:F2");
                        exampleRow.Style.Font.Italic = true;
                        exampleRow.Style.Font.FontColor = XLColor.Gray;

                        // Kẻ bảng và tự động giãn cột
                        ws.Range("A1:F2").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Range("A1:F2").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                        ws.Columns().AdjustToContents();

                        workbook.SaveAs(sfd.FileName);
                    }
                    MessageBox.Show("✅ Đã xuất Form mẫu nhập liệu thành công!", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error); }
            }
        }

        private void btnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            if (!IsAdmin()) return;

            OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls" };
            if (ofd.ShowDialog() == true)
            {
                try
                {
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                    using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataTable dtExcel = reader.AsDataSet(new ExcelDataSetConfiguration() { ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true } }).Tables[0];

                            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                            {
                                conn.Open();
                                int countIn = 0, countUp = 0;
                                string entryTime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

                                foreach (DataRow row in dtExcel.Rows)
                                {
                                    string loai = row["Loại SP"]?.ToString().Trim() ?? "";
                                    string ten = row["Tên SP"]?.ToString().Trim() ?? "";
                                    if (string.IsNullOrEmpty(loai) || string.IsNullOrEmpty(ten)) continue;

                                    string donVi = row.Table.Columns.Contains("Đơn vị") ? (row["Đơn vị"]?.ToString().Trim() ?? "") : "";
                                    int slMoi = (int)ParseNumber(row.Table.Columns.Contains("Số lượng") ? (row["Số lượng"]?.ToString() ?? "0") : "0");
                                    double thanhTien = ParseNumber(row.Table.Columns.Contains("Thành tiền") ? (row["Thành tiền"]?.ToString() ?? "0") : "0");
                                    double giaMoi = slMoi > 0 ? (thanhTien / slMoi) : 0;

                                    string docDate = row.Table.Columns.Contains("Ngày nhập") ? (row["Ngày nhập"]?.ToString().Trim() ?? "") : "";
                                    if (string.IsNullOrEmpty(docDate)) docDate = DateTime.Now.ToString("dd/MM/yyyy");
                                    else if (DateTime.TryParse(docDate, out DateTime pd)) docDate = pd.ToString("dd/MM/yyyy");

                                    // Đảm bảo Danh mục tồn tại
                                    new SQLiteCommand("INSERT OR IGNORE INTO DS_Loai VALUES (@c)", conn) { Parameters = { new SQLiteParameter("@c", loai) } }.ExecuteNonQuery();
                                    new SQLiteCommand("INSERT OR IGNORE INTO DS_SanPham (ProductName, CategoryName) VALUES (@n, @c)", conn) { Parameters = { new SQLiteParameter("@n", ten), new SQLiteParameter("@c", loai) } }.ExecuteNonQuery();
                                    if (!string.IsNullOrEmpty(donVi)) new SQLiteCommand("INSERT OR IGNORE INTO DS_DonVi VALUES (@v)", conn) { Parameters = { new SQLiteParameter("@v", donVi) } }.ExecuteNonQuery();

                                    int oldID = -1; int slCu = 0; double giaCu = 0;
                                    using (var cmdCheck = new SQLiteCommand("SELECT ProductID, Quantity, ImportPrice FROM Inventory WHERE ProductName=@n", conn))
                                    {
                                        cmdCheck.Parameters.AddWithValue("@n", ten);
                                        using (var rd = cmdCheck.ExecuteReader())
                                        {
                                            if (rd.Read()) { oldID = Convert.ToInt32(rd["ProductID"]); slCu = Convert.ToInt32(rd["Quantity"]); giaCu = Convert.ToDouble(rd["ImportPrice"]); }
                                        }
                                    }

                                    if (oldID != -1)
                                    {
                                        // CẬP NHẬT HÀNG CŨ
                                        int tongSL = slCu + slMoi;
                                        double giaTB = tongSL > 0 ? ((giaCu * slCu) + thanhTien) / tongSL : 0;
                                        new SQLiteCommand("UPDATE Inventory SET Quantity=@q, ImportPrice=@p, ImportDate=@d, Importer=@u, Unit=@unit WHERE ProductID=@id", conn) { Parameters = { new SQLiteParameter("@q", tongSL), new SQLiteParameter("@p", Math.Round(giaTB, 2)), new SQLiteParameter("@d", docDate), new SQLiteParameter("@u", currentUsername), new SQLiteParameter("@unit", donVi), new SQLiteParameter("@id", oldID) } }.ExecuteNonQuery();
                                        countUp++;
                                    }
                                    else
                                    {
                                        // =========================================================
                                        // LOGIC TỰ NHẢY MÃ SP CHUẨN THEO YÊU CẦU CỦA BRO
                                        // =========================================================
                                        string prefix = GetPrefixLoaiSP(loai);
                                        string searchPrefix = prefix + "-";
                                        string maSPMoi = searchPrefix + "0001";

                                        // Truy vấn mã lớn nhất hiện có với tiền tố này
                                        string queryMax = "SELECT ProductCode FROM Inventory WHERE ProductCode LIKE @pre ORDER BY ProductCode DESC LIMIT 1";
                                        using (var cmdMax = new SQLiteCommand(queryMax, conn))
                                        {
                                            cmdMax.Parameters.AddWithValue("@pre", searchPrefix + "%");
                                            var lastMa = cmdMax.ExecuteScalar()?.ToString();
                                            if (!string.IsNullOrEmpty(lastMa) && lastMa.StartsWith(searchPrefix))
                                            {
                                                string numberPart = lastMa.Substring(searchPrefix.Length);
                                                if (int.TryParse(numberPart, out int lastNum))
                                                {
                                                    maSPMoi = searchPrefix + (lastNum + 1).ToString("D4");
                                                }
                                            }
                                        }

                                        // THÊM MỚI VÀO DATABASE
                                        string sqlInsert = "INSERT INTO Inventory (ProductCode, ProductName, Category, Unit, Quantity, ImportPrice, ImportDate, SafeLevel, Importer) VALUES (@code, @n, @c, @unit, @q, @p, @d, 0, @u)";
                                        using (var cmdIn = new SQLiteCommand(sqlInsert, conn))
                                        {
                                            cmdIn.Parameters.AddWithValue("@code", maSPMoi);
                                            cmdIn.Parameters.AddWithValue("@n", ten);
                                            cmdIn.Parameters.AddWithValue("@c", loai);
                                            cmdIn.Parameters.AddWithValue("@unit", donVi);
                                            cmdIn.Parameters.AddWithValue("@q", slMoi);
                                            cmdIn.Parameters.AddWithValue("@p", giaMoi);
                                            cmdIn.Parameters.AddWithValue("@d", docDate);
                                            cmdIn.Parameters.AddWithValue("@u", currentUsername);
                                            cmdIn.ExecuteNonQuery();
                                        }
                                        oldID = (int)conn.LastInsertRowId;
                                        countIn++;
                                    }

                                    // Lưu Log nhập kho (áp dụng cho cả cũ và mới)
                                    using (var logCmd = new SQLiteCommand("INSERT INTO ImportLogs (ProductID, Qty, Price, ImportDate, Importer, EntryTime) VALUES (@pid, @q, @p, @d, @u, @et)", conn))
                                    {
                                        logCmd.Parameters.AddWithValue("@pid", oldID);
                                        logCmd.Parameters.AddWithValue("@q", slMoi);
                                        logCmd.Parameters.AddWithValue("@p", giaMoi);
                                        logCmd.Parameters.AddWithValue("@d", docDate);
                                        logCmd.Parameters.AddWithValue("@u", currentUsername);
                                        logCmd.Parameters.AddWithValue("@et", entryTime);
                                        logCmd.ExecuteNonQuery();
                                    }
                                }
                                MessageBox.Show($"Đồng bộ hoàn tất!\nThêm mới: {countIn} mã (Mã tự nhảy chuẩn)\nCộng dồn: {countUp} mã", "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);
                            }
                        }
                    }
                    LoadDuLieuKho(); LoadCombos(); UpdateThongKeData();
                }
                catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi Định Dạng", MessageBoxButton.OK, MessageBoxImage.Error); }
            }
        }

        private void LoadUsers() { using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;")) { conn.Open(); DataTable dt = new DataTable(); new SQLiteDataAdapter("SELECT UserID as [Mã], Username as [Tài khoản], Password as [Mật khẩu], Role as [Quyền], AllowedTabs as [Tab] FROM Users", conn).Fill(dt); dgUsers.ItemsSource = dt.DefaultView; } }
        private string GetSelectedTabs() { var tabs = new List<string>(); if (chkTabKho.IsChecked == true) tabs.Add("tabKho"); if (chkTabBan.IsChecked == true) tabs.Add("tabBan"); if (chkTabThongKe.IsChecked == true) tabs.Add("tabThongKe"); if (chkTabHeThong.IsChecked == true) tabs.Add("tabHeThong"); return string.Join(",", tabs); }
        private void btnAddUser_Click(object sender, RoutedEventArgs e) { if (!IsAdmin()) return; isEditUserMode = false; OverlayHeThong.Visibility = Visibility.Visible; }
        private void btnEditUser_Click(object sender, RoutedEventArgs e) { if (!IsAdmin()) return; if (dgUsers.SelectedItem is DataRowView r) { isEditUserMode = true; editUserID = r["Mã"].ToString(); txtUserName.Text = r["Tài khoản"].ToString(); txtPassword.Text = r["Mật khẩu"].ToString(); cmbRole.Text = r["Quyền"].ToString(); OverlayHeThong.Visibility = Visibility.Visible; } }
        private void btnDeleteUser_Click(object sender, RoutedEventArgs e) { if (!IsAdmin()) return; if (dgUsers.SelectedItem is DataRowView r) { RunSQL("DELETE FROM Users WHERE UserID=@id", c => c.Parameters.AddWithValue("@id", r["Mã"])); LoadUsers(); } }
        private void btnHuyUserPopup_Click(object sender, RoutedEventArgs e) { OverlayHeThong.Visibility = Visibility.Collapsed; }
        private void btnLuuUserPopup_Click(object sender, RoutedEventArgs e) { RunSQL(isEditUserMode ? "UPDATE Users SET Password=@p, Role=@r, AllowedTabs=@t WHERE UserID=@id" : "INSERT INTO Users (Username, Password, Role, AllowedTabs) VALUES (@u, @p, @r, @t)", c => { c.Parameters.AddWithValue("@p", txtPassword.Text); c.Parameters.AddWithValue("@r", cmbRole.Text); c.Parameters.AddWithValue("@t", GetSelectedTabs()); if (isEditUserMode) c.Parameters.AddWithValue("@id", editUserID); else c.Parameters.AddWithValue("@u", txtUserName.Text); }); OverlayHeThong.Visibility = Visibility.Collapsed; LoadUsers(); }
        private void RunSQL(string sql, Action<SQLiteCommand> p) { SQLiteConnection.ClearAllPools(); using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;")) { conn.Open(); using (var cmd = new SQLiteCommand(sql, conn)) { p(cmd); cmd.ExecuteNonQuery(); } } }
        // ==============================================================
        // TÍNH NĂNG TÌM KIẾM TOÀN CỤC (CTRL + F) - CHUẨN EXCEL
        // ==============================================================
        // Tạo class để lưu đúng vị trí Ô chứa kết quả
        // =========================================================
        // HỆ THỐNG TÌM KIẾM (CTRL + F) TỪNG Ô GIỐNG EXCEL
        // =========================================================

        private class SearchMatch
        {
            public object RowItem { get; set; }
            public string ColumnName { get; set; } = "";
        }
        private System.Collections.Generic.List<SearchMatch> currentSearchMatches = new System.Collections.Generic.List<SearchMatch>();
        private int currentSearchIndex = -1;

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                // Xử lý khi bấm Ctrl + F
                if (e.Key == Key.F && Keyboard.Modifiers == ModifierKeys.Control)
                {
                    e.Handled = true;
                    if (SearchBox != null && txtSearchInput != null)
                    {
                        SearchBox.Visibility = Visibility.Visible;
                        txtSearchInput.Focus();
                        txtSearchInput.SelectAll();
                    }
                }
                // Bấm phím ESC để đóng nhanh hộp tìm kiếm
                else if (e.Key == Key.Escape)
                {
                    CloseSearchBox();
                }
            }
            catch { }
        }

        private DataGrid GetActiveDataGrid()
        {
            if (tabKho.IsSelected) return dgKhoHang;
            if (tabBan.IsSelected) return dgBanHang;
            if (tabTKBan.IsSelected) return dgThongKe;
            if (tabTKNhap.IsSelected) return dgThongKeNhap;
            if (tabHeThong.IsSelected) return dgUsers;
            return null;
        }

        private void ExecuteSearch()
        {
            string query = txtSearchInput.Text.Trim().ToLower();
            currentSearchMatches.Clear();
            currentSearchIndex = -1;
            txtSearchMatchCount.Text = "0/0";

            DataGrid dg = GetActiveDataGrid();
            if (dg == null || dg.ItemsSource == null || string.IsNullOrEmpty(query))
            {
                if (dg != null) dg.SelectedCells.Clear();
                return;
            }

            // Bật chế độ chọn Cell (Ô) để giống hệt Excel
            dg.SelectionUnit = DataGridSelectionUnit.CellOrRowHeader;

            // BÍ KÍP HIGHLIGHT: Ép màu vàng rực
            Style cellStyle = new Style(typeof(DataGridCell));
            Trigger selectedTrigger = new Trigger { Property = DataGridCell.IsSelectedProperty, Value = true };
            selectedTrigger.Setters.Add(new Setter(BackgroundProperty, new SolidColorBrush(Color.FromRgb(255, 215, 0))));
            selectedTrigger.Setters.Add(new Setter(ForegroundProperty, Brushes.Black));
            selectedTrigger.Setters.Add(new Setter(FontWeightProperty, FontWeights.Bold));
            cellStyle.Triggers.Add(selectedTrigger);
            dg.CellStyle = cellStyle;

            // Quét từng ô trong bảng
            foreach (var item in dg.ItemsSource)
            {
                if (item is System.Data.DataRowView drv)
                {
                    foreach (System.Data.DataColumn col in drv.Row.Table.Columns)
                    {
                        object value = drv.Row[col];
                        if (value != null && value.ToString().ToLower().Contains(query))
                        {
                            currentSearchMatches.Add(new SearchMatch { RowItem = item, ColumnName = col.ColumnName });
                        }
                    }
                }
            }

            if (currentSearchMatches.Count > 0)
            {
                currentSearchIndex = 0;
                NavigateSearch(0);
            }
        }

        private void NavigateSearch(int step)
        {
            if (currentSearchMatches.Count == 0) return;

            currentSearchIndex += step;
            if (currentSearchIndex < 0) currentSearchIndex = currentSearchMatches.Count - 1;
            if (currentSearchIndex >= currentSearchMatches.Count) currentSearchIndex = 0;

            txtSearchMatchCount.Text = $"{currentSearchIndex + 1}/{currentSearchMatches.Count}";

            DataGrid dg = GetActiveDataGrid();
            if (dg != null)
            {
                var match = currentSearchMatches[currentSearchIndex];

                DataGridColumn targetCol = null;
                foreach (var c in dg.Columns)
                {
                    if ((c.Header != null && c.Header.ToString() == match.ColumnName) || c.SortMemberPath == match.ColumnName)
                    {
                        targetCol = c;
                        break;
                    }
                }

                dg.SelectedCells.Clear();

                if (targetCol != null)
                {
                    dg.UpdateLayout();
                    DataGridCellInfo cellInfo = new DataGridCellInfo(match.RowItem, targetCol);
                    dg.CurrentCell = cellInfo;
                    dg.SelectedCells.Add(cellInfo);
                    dg.ScrollIntoView(match.RowItem, targetCol);
                }
                else
                {
                    dg.SelectedItem = match.RowItem;
                    dg.ScrollIntoView(match.RowItem);
                }
            }
        }

        // HÀM QUAN TRỌNG ĐÃ ĐƯỢC FIX LỖI "TỊT NGÒI"
        private void txtSearchInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                // Phải có dòng này thì phần mềm mới bắt đầu quét dữ liệu khi bạn gõ chữ
                ExecuteSearch();
            }
            catch { }
        }

        private void txtSearchInput_KeyDown(object sender, KeyEventArgs e)
        {
            try { if (e.Key == Key.Enter) NavigateSearch(1); } catch { }
        }

        private void btnSearchNext_Click(object sender, RoutedEventArgs e)
        {
            try { NavigateSearch(1); } catch { }
        }

        private void btnSearchPrev_Click(object sender, RoutedEventArgs e)
        {
            try { NavigateSearch(-1); } catch { }
        }

        private void btnSearchClose_Click(object sender, RoutedEventArgs e)
        {
            try { CloseSearchBox(); } catch { }
        }

        private void CloseSearchBox()
        {
            if (SearchBox != null) SearchBox.Visibility = Visibility.Collapsed;
            if (txtSearchInput != null) txtSearchInput.Clear();

            DataGrid dg = GetActiveDataGrid();
            if (dg != null)
            {
                dg.SelectedCells.Clear();
                dg.CellStyle = null;
                // Trả lại chế độ chọn cả dòng để dễ bấm các nút sửa/xóa
                dg.SelectionUnit = DataGridSelectionUnit.FullRow;
            }
        }
        // ==============================================================
        // TÍNH NĂNG TỰ ĐỘNG SAO LƯU DỮ LIỆU (AUTO BACKUP)
        // ==============================================================
        private void ThucHienAutoBackup()
        {
            try
            {
                // 1. Tạo thư mục Backup nằm ngay cạnh file Database gốc
                string? dir = Path.GetDirectoryName(dbPath);
                string backupFolder = Path.Combine(dir ?? AppDomain.CurrentDomain.BaseDirectory, "Backups");

                if (!Directory.Exists(backupFolder))
                    Directory.CreateDirectory(backupFolder);

                // 2. Tạo tên file backup theo ngày (Ví dụ: Backup_2026_05_02.db)
                string fileName = $"Backup_{DateTime.Now:yyyy_MM_dd}.db";
                string destPath = Path.Combine(backupFolder, fileName);

                // 3. Kiểm tra: Nếu hôm nay chưa có bản backup thì mới copy tạo mới
                if (!File.Exists(destPath))
                {
                    if (File.Exists(dbPath))
                    {
                        File.Copy(dbPath, destPath, true);

                        // 4. Tự động dọn dẹp các bản backup cũ hơn 30 ngày cho nhẹ máy
                        XoaBackupCu(backupFolder, 30);
                    }
                }
            }
            catch (Exception ex)
            {
                // Chỉ ghi log ẩn, không hiện MessageBox làm phiền người dùng lúc khởi động
                Console.WriteLine("Lỗi sao lưu tự động: " + ex.Message);
            }
        }

        private void XoaBackupCu(string folder, int days)
        {
            try
            {
                string[] files = Directory.GetFiles(folder, "Backup_*.db");
                foreach (string file in files)
                {
                    FileInfo fi = new FileInfo(file);
                    if (fi.CreationTime < DateTime.Now.AddDays(-days))
                    {
                        fi.Delete();
                    }
                }
            }
            catch { }
        }
        private void btnRestore_Click(object sender, RoutedEventArgs e)
        {
            // Chỉ Admin mới được quyền phục hồi dữ liệu
            if (currentUserRole != "Admin")
            {
                MessageBox.Show("⛔ Chỉ Admin mới có quyền phục hồi dữ liệu!", "Bảo mật");
                return;
            }

            // Mở hộp thoại chọn file
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Backups");
            ofd.Filter = "Database Files (*.db)|*.db";
            ofd.Title = "Chọn bản sao lưu để phục hồi";

            if (ofd.ShowDialog() == true)
            {
                var result = MessageBox.Show("⚠️ CẢNH BÁO: Toàn bộ dữ liệu hiện tại sẽ bị thay thế bằng bản sao lưu này.\nBạn có chắc chắn muốn tiếp tục?",
                                             "Xác nhận phục hồi", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        // Phải ngắt kết nối hoàn toàn trước khi ghi đè file
                        SQLiteConnection.ClearAllPools();

                        // Ghi đè file backup vào file chính
                        File.Copy(ofd.FileName, dbPath, true);

                        MessageBox.Show("✅ Đã phục hồi dữ liệu thành công! Phần mềm sẽ khởi động lại.", "Hoàn tất");

                        // Khởi động lại app để nạp dữ liệu mới
                        System.Diagnostics.Process.Start(Application.ResourceAssembly.Location);
                        Application.Current.Shutdown();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Lỗi khi phục hồi: " + ex.Message);
                    }
                }
            }
        }
        private void KhoiTaoTemplateMacDinh()
        {
            try
            {
                // 1. Tạo thư mục Templates nếu chưa tồn tại
                if (!Directory.Exists(templateFolder))
                    Directory.CreateDirectory(templateFolder);

                var templates = new Dictionary<string, string[]>
        {
            { "Template_Kho.xlsx", new[] { "STT", "Tên SP", "Loại SP", "Đơn vị", "Tồn", "Giá Vốn", "Mức an toàn", "Người nhập", "Ngày nhập" } },
            { "Template_ThongKeBan.xlsx", new[] { "STT", "Ngày bán", "Loại SP", "Tên SP", "Đơn vị", "Số lượng", "Đơn giá (VNĐ)", "Thành tiền", "Người bán", "Thời gian tạo" } },
            { "Template_ThongKeNhap.xlsx", new[] { "STT", "Tên SP", "Loại SP", "Số lượng", "Doanh thu", "Lợi nhuận" } }
        };

                foreach (var item in templates)
                {
                    string path = Path.Combine(templateFolder, item.Key);

                    if (!File.Exists(path))
                    {
                        using (var workbook = new XLWorkbook())
                        {
                            var ws = workbook.Worksheets.Add("Sheet1");
                            ws.Cell("B2").Value = "HỆ THỐNG QUẢN LÝ BÁN HÀNG PRO";
                            ws.Cell("B2").Style.Font.Bold = true;
                            ws.Cell("B2").Style.Font.FontSize = 16;

                            ws.Cell("B3").Value = "BÁO CÁO CHI TIẾT - " + item.Key.Replace(".xlsx", "").Replace("Template_", "");
                            ws.Cell("B3").Style.Font.Italic = true;

                            ws.Cell("B4").Value = "Ngày xuất:";
                            ws.Cell("B5").Value = "Người xuất:";

                            // Ghi Tiêu đề ở Dòng 9
                            for (int i = 0; i < item.Value.Length; i++)
                            {
                                var cell = ws.Cell(9, 2 + i);
                                cell.Value = item.Value[i];
                                cell.Style.Font.Bold = true;
                                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#007BFF");
                                cell.Style.Font.FontColor = XLColor.White;
                                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            }
                            ws.Columns().AdjustToContents();
                            workbook.SaveAs(path);
                        }
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine("Lỗi tạo template: " + ex.Message); }
        }
        // ==============================================================
        // HÀM XUẤT EXCEL THEO FORM MẪU (DÙNG CHUNG CHO TẤT CẢ CÁC BẢNG)
        // ==============================================================
        private void XuatBaoCaoRaForm(DataTable dtDuLieu, string tenFileTemplate, string tenFileXuatRa)
        {
            if (dtDuLieu == null || dtDuLieu.Rows.Count == 0)
            {
                MessageBox.Show("Bảng không có dữ liệu để xuất!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                // 1. TRỎ THẲNG VÀO THƯ MỤC TEMPLATES
                // Đảm bảo bồ đã khai báo: private readonly string templateFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");
                string templatePath = System.IO.Path.Combine(templateFolder, tenFileTemplate);

                if (!System.IO.File.Exists(templatePath))
                {
                    MessageBox.Show($"Không tìm thấy form mẫu: '{tenFileTemplate}' trong thư mục Templates!\nHệ thống sẽ tự tạo file mẫu cơ bản, hãy chỉnh sửa lại sau.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    KhoiTaoTemplateMacDinh(); // Tự tạo lại nếu lỡ tay xóa mất
                    return;
                }

                using (var workbook = new XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheet(1); // Mở Sheet đầu tiên

                    // 2. Điền thông tin chung (C4: Thời gian, C5: Người xuất)
                    worksheet.Cell("C4").Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
                    worksheet.Cell("C5").Value = currentUsername;

                    // 3. ĐỔ DỮ LIỆU ĐỘNG: Bắt đầu từ Dòng 10, Cột B (Cột 2)
                    // Lệnh InsertData sẽ tự phun toàn bộ DataTable xuống dưới
                    var tableRange = worksheet.Cell(10, 2).InsertData(dtDuLieu.AsEnumerable());

                    // 4. Tự động kẻ khung (Borders) cho vùng dữ liệu vừa phun ra
                    // Tính toán vùng bao phủ: từ dòng 10 đến dòng cuối cùng của dữ liệu
                    var dataRange = worksheet.Range(10, 2, 10 + dtDuLieu.Rows.Count - 1, 2 + dtDuLieu.Columns.Count - 1);

                    dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    dataRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    dataRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // 5. Cho người dùng chọn nơi lưu file kết quả
                    SaveFileDialog sfd = new SaveFileDialog
                    {
                        Filter = "Excel Files|*.xlsx",
                        FileName = $"{tenFileXuatRa}_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
                    };

                    if (sfd.ShowDialog() == true)
                    {
                        workbook.SaveAs(sfd.FileName);
                        MessageBox.Show("✅ Đã xuất báo cáo thành công!", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);

                        // Mở file lên ngay lập tức sau khi xuất
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xuất Form: " + ex.Message, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        // 1. Hàm lấy 2 chữ cái đầu theo logic của bạn
        private string RemoveVietnameseTone(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;
            string result = text.ToLower();
            result = System.Text.RegularExpressions.Regex.Replace(result, "à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ", "a");
            result = System.Text.RegularExpressions.Regex.Replace(result, "è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ", "e");
            result = System.Text.RegularExpressions.Regex.Replace(result, "ì|í|ị|ỉ|ĩ", "i");
            result = System.Text.RegularExpressions.Regex.Replace(result, "ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ", "o");
            result = System.Text.RegularExpressions.Regex.Replace(result, "ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ", "u");
            result = System.Text.RegularExpressions.Regex.Replace(result, "ỳ|ý|ỵ|ỷ|ỹ", "y");
            result = System.Text.RegularExpressions.Regex.Replace(result, "đ", "d");
            return result.ToUpper();
        }

        private string GetPrefixLoaiSP(string loaiSP)
        {
            if (string.IsNullOrWhiteSpace(loaiSP)) return "SP";
            string cleanLoaiSP = RemoveVietnameseTone(loaiSP);
            string[] words = cleanLoaiSP.Trim().Split(new char[] { ' ' }, System.StringSplitOptions.RemoveEmptyEntries);
            if (words.Length == 1)
            {
                if (words[0].Length >= 2) return words[0].Substring(0, 2);
                return words[0] + "X";
            }
            return words[0].Substring(0, 1) + words[1].Substring(0, 1);
        }

        // 2. Hàm thực hiện việc truy vấn database và nhảy số -xxxx
        private void TuDongTaoMaSP()
        {
            try
            {
                string loaiSP = cmbNewLoai.Text;
                if (string.IsNullOrWhiteSpace(loaiSP)) return;

                string prefix = GetPrefixLoaiSP(loaiSP);
                string searchPrefix = prefix + "-"; // Kết quả sẽ có dạng TI- hoặc HA-

                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();
                    // Tìm mã lớn nhất có bắt đầu bằng tiền tố đó
                    string query = "SELECT ProductCode FROM Inventory WHERE ProductCode LIKE @prefix ORDER BY ProductCode DESC LIMIT 1";
                    System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(query, conn);
                    cmd.Parameters.AddWithValue("@prefix", searchPrefix + "%");

                    var lastMaSP = cmd.ExecuteScalar()?.ToString();

                    if (string.IsNullOrEmpty(lastMaSP))
                    {
                        txtNewMaSP.Text = searchPrefix + "0001";
                    }
                    else
                    {
                        // Tách lấy phần số sau dấu gạch ngang
                        string numberPart = lastMaSP.Substring(searchPrefix.Length);
                        if (int.TryParse(numberPart, out int lastNumber))
                        {
                            txtNewMaSP.Text = searchPrefix + (lastNumber + 1).ToString("D4");
                        }
                        else
                        {
                            txtNewMaSP.Text = searchPrefix + "0001";
                        }
                    }
                }
            }
            catch { }
        }

        // 3. Các sự kiện để kích hoạt việc nhảy mã
        private void cmbNewLoai_DropDownClosed(object sender, EventArgs e) => TuDongTaoMaSP();
        private void cmbNewLoai_LostFocus(object sender, RoutedEventArgs e) => TuDongTaoMaSP();
        private void cmbTenSPKho_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbTenSP.SelectedValue != null)
            {
                int id = Convert.ToInt32(cmbTenSP.SelectedValue);
                using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();
                    string sql = $"SELECT Unit, ImportPrice, SafeLevel FROM Inventory WHERE ProductID={id}";
                    using (var cmd = new System.Data.SQLite.SQLiteCommand(sql, conn))
                    {
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                txtDonViKho.Text = reader["Unit"]?.ToString() ?? "";
                                txtGiaNhap.Text = reader["ImportPrice"]?.ToString() ?? "0";
                                txtSafeLevel.Text = reader["SafeLevel"]?.ToString() ?? "0";
                            }
                        }
                    }
                }
            }
        }
        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Cực kỳ quan trọng: Phải kiểm tra đúng sự kiện của TabControl, 
            // vì nếu không, cứ mỗi lần bạn chọn 1 dòng trong ComboBox hay DataGrid nó cũng sẽ chạy vào đây gây lỗi.
            if (e.OriginalSource == MainTabControl)
            {
                if (SearchBox != null && SearchBox.Visibility == Visibility.Visible)
                {
                    // Tắt hộp tìm kiếm
                    SearchBox.Visibility = Visibility.Collapsed;

                    try
                    {
                        // Xóa trống text để reset tìm kiếm (ẩn highlight ở các bảng)
                        txtSearchInput.Text = "";
                    }
                    catch
                    {
                        // Bắt lỗi an toàn đề phòng DataGrid cũ chưa kịp thoát khỏi bộ nhớ
                    }
                }
            }
        }
        private void ClearPhieuNhap()
        {
            // Reset ComboBox chọn sản phẩm
            cmbTenSP.SelectedIndex = -1;
            cmbTenSP.Text = "";

            // Xóa trắng các ô TextBox thông tin
            txtDonViKho.Clear();
            txtGiaNhap.Clear();
            txtSoLuong.Clear(); // Ô nhập số lượng mới
            txtSafeLevel.Clear();

            // Nếu bạn có ô tính Tổng tiền hoặc Ghi chú thì cũng clear luôn
            // txtThanhTienNhap.Text = "0"; 

            // Đưa con trỏ chuột về ô chọn Sản phẩm để bắt đầu nhập món mới luôn
            cmbTenSP.Focus();
        }
        // Class phụ để hứng dữ liệu hiển thị lên bảng cảnh báo
        // Class chuẩn 3 dữ liệu bro yêu cầu
        // 1. Class cấu trúc dữ liệu cho bảng cảnh báo
        public class CanhBaoItem
        {
            public string TenSP { get; set; }
            public int TonKho { get; set; }
            public string TrangThai { get; set; }
            public string MauTrangThai { get; set; }
        }

        // 2. Hàm kiểm tra và bật Overlay
        private void KiemTraHangSapHet()
        {
            try
            {
                var listCanhBao = new System.Collections.Generic.List<CanhBaoItem>();

                using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();
                    // Quét thẳng vào DB, tìm món nào Tồn <= Mức an toàn
                    string sql = "SELECT ProductName, Quantity, SafeLevel FROM Inventory WHERE Quantity <= SafeLevel";

                    using (var cmd = new System.Data.SQLite.SQLiteCommand(sql, conn))
                    {
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                int ton = Convert.ToInt32(reader["Quantity"]);

                                // Gán chính xác Trạng thái và Màu sắc theo tình trạng tồn
                                string trangThai = (ton <= 0) ? "❌ HẾT" : "⚠️ SẮP HẾT";
                                string mau = (ton <= 0) ? "Red" : "#f39c12";

                                listCanhBao.Add(new CanhBaoItem
                                {
                                    TenSP = reader["ProductName"].ToString(),
                                    TonKho = ton,
                                    TrangThai = trangThai,
                                    MauTrangThai = mau
                                });
                            }
                        }
                    }
                }

                // Bơm dữ liệu vào DataGrid và làm nó hiện hình
                if (listCanhBao.Count > 0)
                {
                    dgHangSapHet.ItemsSource = listCanhBao;
                    OverlayCanhBao.Visibility = Visibility.Visible;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Lỗi hiển thị cảnh báo: " + ex.Message);
            }
        }

        // 3. Xử lý tắt Overlay (Click nút)
        private void btnDongCanhBao_Click(object sender, RoutedEventArgs e)
        {
            OverlayCanhBao.Visibility = Visibility.Collapsed;
        }

        // 4. Xử lý tắt Overlay (Click nền đen)
        private void btnDongCanhBao_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            OverlayCanhBao.Visibility = Visibility.Collapsed;
        }
    }
}