using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CRUDMahasiswaADO
{
    public partial class FormReport : Form
    {
        DAL dbLogic = new DAL();

        static string connectionString = "Data Source=ADVAN\\RIFALARDI; Initial Catalog=DBAkademikADO;User ID = sa; Password=PasswordSA";

        SqlConnection conn = new SqlConnection(connectionString);
        SqlDataAdapter da;
        DataTable dtMahasiswa;
        CrystalReport1 listMahasiswa = new CrystalReport1();
        string prodi { get; set; }
        DateTime tglmasuk { get; set; }

        public FormReport(string Prodi, DateTime TglMasuk)
        {
            InitializeComponent();

            prodi = Prodi;
            tglmasuk = TglMasuk;
            try
            {
                DataTable dtMahasiswa = dbLogic.getDataRekap(prodi, tglmasuk);
                listMahasiswa.SetDataSource(dtMahasiswa);
                crystalReportViewer1.ReportSource = listMahasiswa;
                crystalReportViewer1.Refresh();
            }
            catch (Exception ex)
            {
                //simpanLog(ex.Message);
                MessageBox.Show("Gagal load data: " + ex.Message);
            }
        }
    }
    }

