using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QuanLyNhaHang
{
    public partial class QuanLyNhapkho: Form
    {
        string connectionString = ConfigurationManager.ConnectionStrings["QuanLyNhaHang.Properties.Settings.RestaurantManagementConnectionString"].ConnectionString;
        SqlConnection conn;
        public QuanLyNhapkho()
        {
            InitializeComponent();
        }
        private void QuanLyNhapKho_Load(object sender, EventArgs e)
        {
            // Code xử lý khi form load
        }
        private void FmNhapKho_Load(object sender, EventArgs e)
        {
            conn = new SqlConnection(connectionString);
            LoadNhapKho();
            LoadNhanVien();
            LoadNguyenLieu();
        }
        private void LoadNhapKho()
        {
            string query = "SELECT * FROM NhapKhos";
            SqlDataAdapter da = new SqlDataAdapter(query, conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dgvNhapKho.DataSource = dt;
        }
        private void LoadNhanVien()
        {
            string query = "SELECT MaNhanVien, TenNhanVien FROM NhanViens";
            SqlDataAdapter da = new SqlDataAdapter(query, conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            cbxNhanVien.DataSource = dt;
            cbxNhanVien.DisplayMember = "TenNhanVien";
            cbxNhanVien.ValueMember = "MaNhanVien";
        }
        private void LoadNguyenLieu()
        {
            string query = "SELECT MaNguyenLieu, TenNguyenLieu FROM NguyenLieus";
            SqlDataAdapter da = new SqlDataAdapter(query, conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            cbxNguyenLieu.DataSource = dt;
            cbxNguyenLieu.DisplayMember = "TenNguyenLieu";
            cbxNguyenLieu.ValueMember = "MaNguyenLieu";
        } private void btnLuu_Click(object sender, EventArgs e)
        {
            string query = "INSERT INTO NhapKhos (MaNhanVien, MaNguyenLieu, SoLuong, NgayNhap, TongTien, SoNgayHetHan) VALUES (@MaNV, @MaNL, @SoLuong, @NgayNhap, @TongTien, @SoNgayHetHan)";
            SqlCommand cmd = new SqlCommand(query, conn);
            cmd.Parameters.AddWithValue("@MaNV", cbxNhanVien.SelectedValue);
            cmd.Parameters.AddWithValue("@MaNL", cbxNguyenLieu.SelectedValue);
            cmd.Parameters.AddWithValue("@SoLuong", int.Parse(txtSoLuong.Text));
            cmd.Parameters.AddWithValue("@NgayNhap", dtPkNgayNhap.Value);
            cmd.Parameters.AddWithValue("@TongTien", float.Parse(txtTongTien.Text));
            cmd.Parameters.AddWithValue("@SoNgayHetHan", int.Parse(txtSoNgayHetHan.Text));

            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();

            LoadNhapKho();
            MessageBox.Show("Đã thêm phiếu nhập kho!");
        }
        private void btnSua_Click(object sender, EventArgs e)
        {
            string query = "UPDATE NhapKhos SET MaNhanVien=@MaNV, MaNguyenLieu=@MaNL, SoLuong=@SoLuong, " +
                           "NgayNhap=@NgayNhap, TongTien=@TongTien, SoNgayHetHan=@SoNgayHetHan WHERE MaNhapKho=@MaNK";
            SqlCommand cmd = new SqlCommand(query, conn);
            cmd.Parameters.AddWithValue("@MaNK", int.Parse(txtMaNhapKho.Text));
            cmd.Parameters.AddWithValue("@MaNV", cbxNhanVien.SelectedValue);
            cmd.Parameters.AddWithValue("@MaNL", cbxNguyenLieu.SelectedValue);
            cmd.Parameters.AddWithValue("@SoLuong", int.Parse(txtSoLuong.Text));
            cmd.Parameters.AddWithValue("@NgayNhap", dtPkNgayNhap.Value);
            cmd.Parameters.AddWithValue("@TongTien", decimal.Parse(txtTongTien.Text));
            cmd.Parameters.AddWithValue("@SoNgayHetHan", int.Parse(txtSoNgayHetHan.Text));

            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();

            LoadNhapKho();
            MessageBox.Show("Đã sửa phiếu nhập kho!");
        }
        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc muốn xóa phiếu nhập kho này?", "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                string query = "DELETE FROM NhapKhos WHERE MaNhapKho = @MaNK";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@MaNK", int.Parse(txtMaNhapKho.Text));

                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();

                LoadNhapKho();
                MessageBox.Show("Đã xóa phiếu nhập kho!");
            }
        }
        private void btnReload_Click(object sender, EventArgs e)
        {
            LoadNhapKho();
        }
        private void btnThemMoi_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtSoLuong.Text) ||
                    string.IsNullOrWhiteSpace(txtTongTien.Text) ||
                    string.IsNullOrWhiteSpace(txtSoNgayHetHan.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string query = "INSERT INTO NhapKhos (MaNhanVien, MaNguyenLieu, SoLuong, NgayNhap, TongTien, SoNgayHetHan) " +
                               "VALUES (@MaNV, @MaNL, @SoLuong, @NgayNhap, @TongTien, @SoNgayHetHan)";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@MaNV", cbxNhanVien.SelectedValue);
                    cmd.Parameters.AddWithValue("@MaNL", cbxNguyenLieu.SelectedValue);
                    cmd.Parameters.AddWithValue("@SoLuong", int.Parse(txtSoLuong.Text));
                    cmd.Parameters.AddWithValue("@NgayNhap", dtPkNgayNhap.Value);
                    cmd.Parameters.AddWithValue("@TongTien", decimal.Parse(txtTongTien.Text));
                    cmd.Parameters.AddWithValue("@SoNgayHetHan", int.Parse(txtSoNgayHetHan.Text));

                    conn.Open();
                    int rowsAffected = cmd.ExecuteNonQuery();
                    conn.Close();

                    if (rowsAffected > 0)
                    {
                        LoadNhapKho();
                        MessageBox.Show("Thêm đơn hàng thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Thêm đơn hàng thất bại!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc muốn thoát?", "Xác nhận", MessageBoxButtons.OKCancel) == DialogResult.OK)
                this.Close();
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click_1(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void cbxNhanVien_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void dgvNhapKho_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Kiểm tra nếu dòng được chọn hợp lệ
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dgvNhapKho.Rows[e.RowIndex];

                // Gán giá trị từ DataGridView vào các TextBox
                txtMaNhapKho.Text = row.Cells["MaNhapKho"].Value?.ToString() ?? "";
                cbxNhanVien.SelectedValue = row.Cells["MaNhanVien"].Value;
                cbxNguyenLieu.SelectedValue = row.Cells["MaNguyenLieu"].Value;
                txtSoLuong.Text = row.Cells["SoLuong"].Value?.ToString() ?? "";
                dtPkNgayNhap.Value = Convert.ToDateTime(row.Cells["NgayNhap"].Value);
                txtTongTien.Text = row.Cells["TongTien"].Value?.ToString() ?? "";
                txtSoNgayHetHan.Text = row.Cells["SoNgayHetHan"].Value?.ToString() ?? "";
            }
        }


        private void txtTongTien_TextChanged(object sender, EventArgs e)
        {

        }

        private void cbxNguyenLieu_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void btnXuatThongTin_Click(object sender, EventArgs e)
        {
            // Câu truy vấn để lấy thông tin từ bảng NhapKhos
            string query = "SELECT NK.MaNhapKho, NV.TenNhanVien, NL.TenNguyenLieu, NK.SoLuong, NK.NgayNhap, NK.TongTien, NK.SoNgayHetHan " +
                           "FROM NhapKhos NK " +
                           "JOIN NhanViens NV ON NK.MaNhanVien = NV.MaNhanVien " +
                           "JOIN NguyenLieus NL ON NK.MaNguyenLieu = NL.MaNguyenLieu";

            SqlDataAdapter da = new SqlDataAdapter(query, conn);
            DataTable dt = new DataTable();
            da.Fill(dt);

            // Xuất dữ liệu ra DataGridView (hoặc có thể xuất ra file Excel, PDF, etc.)
            dgvNhapKho.DataSource = dt;

            // Tùy chọn: Xuất ra file Excel hoặc PDF (Dùng thư viện như ClosedXML hoặc iTextSharp)
            ExportToExcel(dt);  // Hàm ExportToExcel bạn có thể tự xây dựng
        }

        private void ExportToExcel(DataTable dt)
        {
            // Sử dụng thư viện ClosedXML để xuất dữ liệu ra Excel
            var workbook = new ClosedXML.Excel.XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Danh Sách Nhập Kho");

            // Ghi dữ liệu vào Excel
            worksheet.Cell(1, 1).Value = "Mã Nhập Kho";
            worksheet.Cell(1, 2).Value = "Tên Nhân Viên";
            worksheet.Cell(1, 3).Value = "Tên Nguyên Liệu";
            worksheet.Cell(1, 4).Value = "Số Lượng";
            worksheet.Cell(1, 5).Value = "Ngày Nhập";
            worksheet.Cell(1, 6).Value = "Tổng Tiền";
            worksheet.Cell(1, 7).Value = "Số Ngày Hết Hạn";

            // Lặp qua dữ liệu và điền vào Excel
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                worksheet.Cell(i + 2, 1).Value = dt.Rows[i]["MaNhapKho"].ToString();
                worksheet.Cell(i + 2, 2).Value = dt.Rows[i]["TenNhanVien"].ToString();
                worksheet.Cell(i + 2, 3).Value = dt.Rows[i]["TenNguyenLieu"].ToString();
                worksheet.Cell(i + 2, 4).Value = dt.Rows[i]["SoLuong"].ToString();
                worksheet.Cell(i + 2, 5).Value = dt.Rows[i]["NgayNhap"].ToString();
                worksheet.Cell(i + 2, 6).Value = dt.Rows[i]["TongTien"].ToString();
                worksheet.Cell(i + 2, 7).Value = dt.Rows[i]["SoNgayHetHan"].ToString();

            }

            // Lưu file Excel
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel files|*.xlsx";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(saveFileDialog.FileName);
                MessageBox.Show("Đã xuất thông tin nhập kho ra Excel!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

    }
}
