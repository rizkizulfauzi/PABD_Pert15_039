using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using ExcelDataReader;

namespace CRUDMahasiswaADO
{
    public partial class Form1 : Form
    {
        DAL dbLogic = new DAL();


        private BindingSource bindingSource = new BindingSource();
        private DataTable dtMahasiswa = new DataTable();
        private readonly SqlConnection conn;
        private readonly String connectionString = "Data Source=NJOELL18\\MRZULFAUZI;Initial Catalog=DBAkademikADO;User ID=sa;Password=Mrzxxx18";

        public Form1()
        {
            InitializeComponent();
            conn = new SqlConnection(connectionString);
        }

        private void LoadData()
        {
            try
            {
                bindingSource.DataSource = dbLogic.GetMhs();
                dataGridView1.DataSource = bindingSource;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        private void BindControls()
        {
            txtNIM.DataBindings.Clear();
            txtNama.DataBindings.Clear();
            cmbJK.DataBindings.Clear();
            dtpTanggalLahir.DataBindings.Clear();
            txtAlamat.DataBindings.Clear();
            txtKodeProdi.DataBindings.Clear();

            txtNIM.DataBindings.Add("Text", bindingSource, "NIM");
            txtNama.DataBindings.Add("Text", bindingSource, "Nama");
            cmbJK.DataBindings.Add("Text", bindingSource, "JenisKelamin");
            dtpTanggalLahir.DataBindings.Add("Value", bindingSource, "TanggalLahir");
            txtAlamat.DataBindings.Add("Text", bindingSource, "Alamat");
            txtKodeProdi.DataBindings.Add("Text", bindingSource, "KodeProdi");

        }

        private void HitungTotal()
        {
            try
            {
                int total = (dbLogic.CountMhs().Equals(DBNull.Value)) ? 0 : dbLogic.CountMhs();
                lblTotal.Text = "Total Mahasiswa : " + total;
            }
            catch (Exception ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("Gagal load data: " + ex.Message);
            }
        }

        private void FormMahasiswa_Load(object sender, EventArgs e)
        {

            cmbJK.DataSource = new string[] { "L", "P" };

            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = false;
            dataGridView1.ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;


            bindingNavigator1.BindingSource = bindingSource;
            LoadData();
        }


        private void btnConnect_Click(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(DAL.GetConnectionString()))
                {
                    conn.Open();
                    MessageBox.Show("Koneksi berhasil");
                }
            }
            catch (SqlException ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("SQL Error :" + ex.Message);
            }
            catch (Exception ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("General Error :" + ex.Message);
            }
        }

        private void ConnectDatabase()
        {
            try
            {
                if (conn.State == System.Data.ConnectionState.Closed)
                {
                    conn.Open();
                }
                MessageBox.Show("Koneksi berhasil");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Koneksi gagal: " + ex.Message);
            }
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] ConvertImageToBytes(PictureBox pb)
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        pb.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                        return ms.ToArray();
                    }
                }

                byte[] imgBytes = ConvertImageToBytes(fotoMhs);
                dbLogic.UpdateMhs(txtNIM.Text, txtNama.Text, txtAlamat.Text, cmbJK.Text,
                                  dtpTanggalLahir.Value.Date, txtKodeProdi.Text, imgBytes);

                MessageBox.Show("Data mahasiswa berhasil diubah!");
                ClearForm();
                btnLoad.PerformClick();
            }
            catch (SqlException ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("SQL Error :" + ex.Message);
            }
            catch (Exception ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("General Error :" + ex.Message);
            }
        }


        private void btnInsert_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] ConvertImageToBytes(PictureBox pb)
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        pb.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                        return ms.ToArray();
                    }
                }

                byte[] imgBytes = ConvertImageToBytes(fotoMhs);
                dbLogic.InsertMhs(txtNIM.Text, txtNama.Text, txtAlamat.Text, cmbJK.Text,
                                  dtpTanggalLahir.Value.Date, txtKodeProdi.Text, imgBytes);

                MessageBox.Show("Data mahasiswa berhasil ditambahkan!");
                ClearForm();
                LoadData();
            }
            catch (SqlException ex)
            {
                simpanLog("Rollback Insert :" + ex.Message);
                MessageBox.Show("SQL Error :" + ex.Message);
            }
            catch (Exception ex)
            {
                simpanLog("General Error :" + ex.Message);
                MessageBox.Show("General Error :" + ex.Message);
            }
        }


        private void SimpanLog(string pesan)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = @"INSERT INTO LogError (waktu, pesan_error)
                         VALUES(GETDATE(), @pesan)";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@pesan", pesan);

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void simpanLog(string message)
        {
            dbLogic.InsertLog(message);
        }



        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dg = MessageBox.Show(
                    "Yakin ingin menghapus data?",
                    "Konfirmasi",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (dg == DialogResult.Yes)
                {
                    dbLogic.DeleteMhs(txtNIM.Text);
                    MessageBox.Show("Data mahasiswa berhasil dihapus");
                    ClearForm();
                    btnLoad.PerformClick();
                }
            }
            catch (SqlException ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("SQL Error : " + ex.Message);
            }
            catch (Exception ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("General Error : " + ex.Message);
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataRow row = ((DataRowView)bindingSource[e.RowIndex]).Row;

                // isi textbox sesuai urutan kolom di DataTable
                txtNIM.Text = row[0].ToString();
                txtNama.Text = row[1].ToString();
                cmbJK.Text = row[2].ToString();   // kalau di modul kadang cmbKota, sesuaikan dengan field kamu
                dtpTanggalLahir.Value = Convert.ToDateTime(row[3]);
                txtAlamat.Text = row[4].ToString();

                // kolom 5 = Foto (BLOB)
                if (row[5] != DBNull.Value)
                {
                    byte[] imgBytes = (byte[])row[5];
                    using (MemoryStream ms = new MemoryStream(imgBytes))
                    {
                        fotoMhs.Image = Image.FromStream(ms);
                        fotoMhs.SizeMode = PictureBoxSizeMode.StretchImage;
                    }
                }
                else
                {
                    fotoMhs.Image = null;
                }

                txtKodeProdi.Text = row[6].ToString();

                // disable NIM supaya tidak bisa diubah saat update
                txtNIM.Enabled = false;
            }
        }

        private void ClearForm()
        {
            txtNIM.Enabled = true;
            txtNIM.Clear();
            txtNama.Clear();
            cmbJK.SelectedIndex = -1;
            txtAlamat.Clear();
            txtKodeProdi.Clear();
            dtpTanggalLahir.Value = DateTime.Now;
            fotoMhs.Image = null;
            txtNIM.Focus();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cmbJK.Items.Clear();
            cmbJK.Items.Add("L");
            cmbJK.Items.Add("P");


            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = false;
            dataGridView1.ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            dataGridView1.CellContentClick += dataGridView1_CellContentClick;
        }

        private void txtNIM_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtNama_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'dBAkademikADODataSet.Mahasiswa' table. You can move, or remove it, as needed.
            this.mahasiswaTableAdapter.Fill(this.dBAkademikADODataSet.Mahasiswa);

        }

        private void btnSimpan_Click(object sender, EventArgs e)
        {
            byte[] fotoBytes = null;

            // cek dulu apakah PictureBox ada gambar
            if (fotoMhs.Image != null)
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    fotoMhs.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                    fotoBytes = ms.ToArray();
                }
            }
            else
            {
                MessageBox.Show("Foto belum dipilih!");
            }

            // kirim ke DAL (kalau fotoBytes null, bisa disimpan sebagai DBNull.Value)
            dbLogic.InsertMhs(txtNIM.Text, txtNama.Text, txtAlamat.Text, cmbJK.Text, dtpTanggalLahir.MaxDate, txtKodeProdi.Text, fotoBytes);
        }

        

        private void btnResetData_Click(object sender, EventArgs e)
        {
            try
            {
                dbLogic.resetData();
                MessageBox.Show("Data berhasil direset");
                LoadData();
            }
            catch (SqlException ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("SQL Error : " + ex.Message);
            }
            catch (Exception ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("General Error : " + ex.Message);
            }
        }



        private void btnTestInjection_Click(object sender, EventArgs e)
        {
            {
                try
                {
                    dbLogic.testInject(txtNIM.Text);
                    LoadData();
                }
                catch (SqlException ex)
                {
                    if (ex.Message.Contains("safe"))
                    {
                        simpanLog(ex.Message);
                        MessageBox.Show("SQL Error : Unsafe UPDATE operation not allowed");
                    }
                    else
                    {
                        simpanLog(ex.Message);
                        MessageBox.Show("SQL Error : " + ex.Message);
                    }
                }
                catch (Exception ex)
                {
                    simpanLog(ex.Message);
                    MessageBox.Show("General Error : " + ex.Message);
                }

            }
        }

        private void lblTotal_Click(object sender, EventArgs e)
        {

        }

        private void RekapMahasiswa_Click(object sender, EventArgs e)
        {
            report fm3 = new report();
            fm3.Show();
            this.Hide();
        }

        private void bindingNavigator1_RefreshItems(object sender, EventArgs e)
        {

        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                fotoMhs.Image = Image.FromFile(ofd.FileName);
                fotoMhs.SizeMode = PictureBoxSizeMode.StretchImage;
            }
        }

        private void btnImportFormExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xlsx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });

                        dataGridView1.DataSource = result.Tables[0];
                    }
                }

                // Atur tombol sesuai kondisi setelah import
                btnImportToDatabase.Enabled = true;
                btnInsert.Enabled = false;
                btnUpdate.Enabled = false;
                btnDelete.Enabled = false;
                btnCari.Enabled = false;
                btnResetData.Enabled = false;
                btnTestInjection.Enabled = false;
            }
        }

        private void btnImportToDatabase_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = (DataTable)dataGridView1.DataSource;
                if (dt == null || dt.Rows.Count == 0)
                {
                    MessageBox.Show("Tidak ada data untuk diimport.");
                    return;
                }

                int sukses = 0;

                foreach (DataRow row in dt.Rows)
                {
                    string nim = row["NIM"].ToString().Trim();
                    string nama = row["Nama"].ToString().Trim();
                    string alamat = row["Alamat"].ToString().Trim();
                    string jk = row["JenisKelamin"].ToString().Trim();
                    string kodeProdi = row["KodeProdi"].ToString().Trim();

                    if (string.IsNullOrEmpty(nim) || string.IsNullOrEmpty(nama))
                        continue;

                    DateTime tglLahir;
                    if (!DateTime.TryParse(row["TanggalLahir"].ToString(), out tglLahir))
                        continue;

                    string fotoPath = row.Table.Columns.Contains("FotoPath")
                        ? row["FotoPath"].ToString().Trim()
                        : string.Empty;

                    // panggil helper
                    byte[] fotoBytes = ConvertImageFromPath(fotoPath);

                    dbLogic.InsertMhs(nim, nama, alamat, jk, tglLahir, kodeProdi, fotoBytes);
                    sukses++;
                }

                MessageBox.Show("Data mahasiswa berhasil ditambah: " + sukses);
                ClearForm();
                LoadData();
            }
            catch (SqlException ex)
            {
                MessageBox.Show("SqlException Error : " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("General Error : " + ex.Message);
            }
        }

        // helper method HARUS di dalam class Form1
        private byte[] ConvertImageFromPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return null;

            if (!File.Exists(path))
                return null;

            return File.ReadAllBytes(path);
        }

        private void btnCari_Click(object sender, EventArgs e)
        {
            string nim = txtNIM.Text;

            DataTable dt = dbLogic.GetMahasiswaByNIM(nim);

            if (dt.Rows.Count > 0)
            {
                txtNama.Text = dt.Rows[0]["Nama"].ToString();

                // kalau ada foto
                if (dt.Rows[0]["Foto"] != DBNull.Value)
                {
                    byte[] fotoBytes = (byte[])dt.Rows[0]["Foto"];
                    using (MemoryStream ms = new MemoryStream(fotoBytes))
                    {
                        fotoMhs.Image = Image.FromStream(ms);
                        fotoMhs.SizeMode = PictureBoxSizeMode.StretchImage;
                    }
                }
                else
                {
                    fotoMhs.Image = null;
                }
            }
            else
            {
                MessageBox.Show("Data tidak ditemukan!");
            }
        }
    }

}
