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
    public partial class QuanLyXuatKho: Form
    {
        string connectionString = ConfigurationManager.ConnectionStrings["QuanLyNhaHang.Properties.Settings.RestaurantManagementConnectionString"].ConnectionString;
        SqlConnection conn;
        public QuanLyXuatKho()
        {
            InitializeComponent();

        }

        private void QuanLyXuatKho_Load(object sender, EventArgs e)
        {
            conn = new SqlConnection(connectionString);
            LoadXuatKho();
            LoadNhanVien();
            LoadNguyenLieu();

            // Ẩn cột TenNhanVien nếu không cần hiển thị
            dgvXuatKho.Columns["TenNhanVien"].Visible = false;  // Ẩn tên nhân viên
            dgvXuatKho.Columns["TenNguyenLieu"].Visible = false;  // Ẩn tên nguyên liệu
        }

        private void LoadXuatKho()
        {
            string query = "SELECT XK.MaXuatKho, NV.MaNhanVien, NV.TenNhanVien, NL.MaNguyenLieu, NL.TenNguyenLieu, XK.SoLuong, XK.NgayXuat, XK.NguyenNhanXuatKho, XK.MaLuuTru " +
                           "FROM XuatKhos XK " +
                           "JOIN NhanViens NV ON XK.MaNhanVien = NV.MaNhanVien " +
                           "JOIN NguyenLieus NL ON XK.MaNguyenLieu = NL.MaNguyenLieu";
            SqlDataAdapter da = new SqlDataAdapter(query, conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dgvXuatKho.DataSource = dt;
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
        }
        private void btnLuu_Click(object sender, EventArgs e)
        {
            string query = "INSERT INTO XuatKhos (MaNhanVien, MaNguyenLieu, SoLuong, NgayXuat, NguyenNhanXuatKho, MaLuuTru) " +
                           "VALUES (@MaNV, @MaNL, @SoLuong, @NgayXuat, @NguyenNhanXuatKho, @MaLuuTru)";
            SqlCommand cmd = new SqlCommand(query, conn);
            cmd.Parameters.AddWithValue("@MaNV", cbxNhanVien.SelectedValue);
            cmd.Parameters.AddWithValue("@MaNL", cbxNguyenLieu.SelectedValue);
            cmd.Parameters.AddWithValue("@SoLuong", int.Parse(txtSoLuong.Text));
            cmd.Parameters.AddWithValue("@NgayXuat", dtpNgayXuat.Value);
            cmd.Parameters.AddWithValue("@NguyenNhanXuatKho", txtNguyenNhan.Text);
            cmd.Parameters.AddWithValue("@MaLuuTru", Convert.ToInt32(txtMaLuuTru.Text));

            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();

            LoadXuatKho();
            MessageBox.Show("Đã thêm xuất kho thành công!");
        }
        private void btnSua_Click(object sender, EventArgs e)
        {
            string query = "UPDATE XuatKhos SET MaNhanVien=@MaNV, MaNguyenLieu=@MaNL, SoLuong=@SoLuong, " +
                           "NgayXuat=@NgayXuat, NguyenNhanXuatKho=@NguyenNhanXuatKho, MaLuuTru=@MaLuuTru WHERE MaXuatKho=@MaXK";
            SqlCommand cmd = new SqlCommand(query, conn);
            cmd.Parameters.AddWithValue("@MaXK", int.Parse(txtMaXuatKho.Text));
            cmd.Parameters.AddWithValue("@MaNV", cbxNhanVien.SelectedValue);
            cmd.Parameters.AddWithValue("@MaNL", cbxNguyenLieu.SelectedValue);
            cmd.Parameters.AddWithValue("@SoLuong", int.Parse(txtSoLuong.Text));
            cmd.Parameters.AddWithValue("@NgayXuat", dtpNgayXuat.Value);
            cmd.Parameters.AddWithValue("@NguyenNhanXuatKho", txtNguyenNhan.Text);
            cmd.Parameters.AddWithValue("@MaLuuTru", Convert.ToInt32(txtMaLuuTru.Text));

            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();

            LoadXuatKho();
            MessageBox.Show("Đã sửa xuất kho thành công!");
        }
        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc muốn xóa xuất kho này?", "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                string query = "DELETE FROM XuatKhos WHERE MaXuatKho = @MaXK";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@MaXK", int.Parse(txtMaXuatKho.Text));

                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();

                LoadXuatKho();
                MessageBox.Show("Đã xóa xuất kho thành công!");
            }
        }
        private void btnReload_Click(object sender, EventArgs e)
        {
            LoadXuatKho();
        }
        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc muốn thoát?", "Xác nhận", MessageBoxButtons.OKCancel) == DialogResult.OK)
                this.Close();
        }
        private void dgvXuatKho_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dgvXuatKho.Rows[e.RowIndex];
                txtMaXuatKho.Text = row.Cells["MaXuatKho"].Value?.ToString();
                cbxNhanVien.SelectedValue = row.Cells["MaNhanVien"].Value;
                cbxNguyenLieu.SelectedValue = row.Cells["MaNguyenLieu"].Value;
                txtSoLuong.Text = row.Cells["SoLuong"].Value?.ToString();
                dtpNgayXuat.Value = Convert.ToDateTime(row.Cells["NgayXuat"].Value);
                txtNguyenNhan.Text = row.Cells["NguyenNhanXuatKho"].Value?.ToString();
                txtMaLuuTru.Text = row.Cells["MaLuuTru"].Value?.ToString();
            }
        }

        private void dgvNhapKho_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }// Button Thêm: Thêm dữ liệu mới vào bảng XuatKhos
        private void btnThem_Click(object sender, EventArgs e)
        {
            try
            {
                // Kiểm tra các trường nhập liệu
                if (string.IsNullOrWhiteSpace(txtSoLuong.Text) ||
                    string.IsNullOrWhiteSpace(txtNguyenNhan.Text) ||
                    string.IsNullOrWhiteSpace(txtMaLuuTru.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Chuẩn bị câu truy vấn INSERT để thêm dữ liệu vào XuatKhos
                string query = "INSERT INTO XuatKhos (MaNhanVien, MaNguyenLieu, SoLuong, NgayXuat, NguyenNhanXuatKho, MaLuuTru) " +
                               "VALUES (@MaNV, @MaNL, @SoLuong, @NgayXuat, @NguyenNhanXuatKho, @MaLuuTru)";

                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@MaNV", cbxNhanVien.SelectedValue);  // Lấy mã nhân viên từ ComboBox
                cmd.Parameters.AddWithValue("@MaNL", cbxNguyenLieu.SelectedValue);  // Lấy mã nguyên liệu từ ComboBox
                cmd.Parameters.AddWithValue("@SoLuong", int.Parse(txtSoLuong.Text));  // Lấy số lượng từ TextBox
                cmd.Parameters.AddWithValue("@NgayXuat", dtpNgayXuat.Value);  // Lấy ngày xuất từ DateTimePicker
                cmd.Parameters.AddWithValue("@NguyenNhanXuatKho", txtNguyenNhan.Text);  // Lý do xuất kho từ TextBox
                cmd.Parameters.AddWithValue("@MaLuuTru", Convert.ToInt32(txtMaLuuTru.Text));  // Mã lưu trữ từ TextBox

                // Mở kết nối và thực thi câu lệnh INSERT
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();

                // Cập nhật lại danh sách trong DataGridView
                LoadXuatKho();

                // Thông báo thêm thành công
                MessageBox.Show("Đã thêm xuất kho thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // Hiển thị lỗi nếu có
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void txtNguyenNhan_TextChanged(object sender, EventArgs e)
        {

        }
        private void btnXuatThongTinXuatKho_Click(object sender, EventArgs e)
        {
            string query = "SELECT XK.MaXuatKho, NV.TenNhanVien, NL.TenNguyenLieu, XK.SoLuong, XK.NgayXuat, XK.NguyenNhanXuatKho, XK.MaLuuTru " +
                           "FROM XuatKhos XK " +
                           "JOIN NhanViens NV ON XK.MaNhanVien = NV.MaNhanVien " +
                           "JOIN NguyenLieus NL ON XK.MaNguyenLieu = NL.MaNguyenLieu";

            SqlDataAdapter da = new SqlDataAdapter(query, conn);
            DataTable dt = new DataTable();
            da.Fill(dt);

            // Gọi hàm ExportToExcelForXuatKho để xuất thông tin ra Excel
            ExportToExcelForXuatKho(dt);
        }
        private void ExportToExcelForXuatKho(DataTable dt)
        {
            // Sử dụng thư viện ClosedXML để xuất dữ liệu ra Excel
            var workbook = new ClosedXML.Excel.XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Danh Sách Xuất Kho");

            // Ghi tiêu đề cột vào Excel
            worksheet.Cell(1, 1).Value = "Mã Xuất Kho";
            worksheet.Cell(1, 2).Value = "Tên Nhân Viên";
            worksheet.Cell(1, 3).Value = "Tên Nguyên Liệu";
            worksheet.Cell(1, 4).Value = "Số Lượng";
            worksheet.Cell(1, 5).Value = "Ngày Xuất";
            worksheet.Cell(1, 6).Value = "Nguyên Nhân Xuất Kho";
            worksheet.Cell(1, 7).Value = "Mã Lưu Trữ";

            // Lặp qua dữ liệu và điền vào Excel
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                worksheet.Cell(i + 2, 1).Value = dt.Rows[i]["MaXuatKho"].ToString();
                worksheet.Cell(i + 2, 2).Value = dt.Rows[i]["TenNhanVien"].ToString();
                worksheet.Cell(i + 2, 3).Value = dt.Rows[i]["TenNguyenLieu"].ToString();
                worksheet.Cell(i + 2, 4).Value = dt.Rows[i]["SoLuong"].ToString();
                worksheet.Cell(i + 2, 5).Value = dt.Rows[i]["NgayXuat"].ToString();
                worksheet.Cell(i + 2, 6).Value = dt.Rows[i]["NguyenNhanXuatKho"].ToString();
                worksheet.Cell(i + 2, 7).Value = dt.Rows[i]["MaLuuTru"].ToString();
            }

            // Lưu file Excel
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel files|*.xlsx";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(saveFileDialog.FileName);
                MessageBox.Show("Đã xuất thông tin xuất kho ra Excel!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

    }
}
