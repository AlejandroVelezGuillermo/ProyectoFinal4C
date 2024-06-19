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
using System.IO; // Necesario para manejo de archivos
using System.Data.OleDb; // Necesario para leer desde Excel

namespace Bases_Persona
{
    public partial class Datos : Form
    {
        static string ConexionString = @"server = localhost\SQLEXPRESS; database = Musica; integrated security = true";
        SqlConnection Conexion = new SqlConnection(ConexionString);
        SqlDataAdapter Adaptador;
        DataTable TablaDatos;

        public Datos()
        {
            InitializeComponent();
            string consulta = "select * from dbo.Spotify_Songs";
            Adaptador = new SqlDataAdapter(consulta, Conexion);
            TablaDatos = new DataTable();
            Conexion.Open();
            Adaptador.Fill(TablaDatos);
            dataGridView1.DataSource = TablaDatos;

            // Añadir los botones al formulario
            Button btnExportar = new Button { Text = "Exportar a Excel", Location = new Point(10, 300) };
            btnExportar.Click += BtnExportar_Click;
            this.Controls.Add(btnExportar);

            Button btnImportar = new Button { Text = "Importar desde Excel", Location = new Point(150, 300) };
            btnImportar.Click += BtnImportar_Click;
            this.Controls.Add(btnImportar);
        }

        private void BtnBuscar_Click(object sender, EventArgs e)
        {
            Busqueda1();
        }

        private void BtnRefrescar_Click(object sender, EventArgs e)
        {
            Recargar();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Busqueda2();
        }

        private void BtnExportar_Click(object sender, EventArgs e)
        {
            ExportarExcel();
        }

        private void BtnImportar_Click(object sender, EventArgs e)
        {
            ImportarDesdeExcel();
        }

        private void Busqueda1()
        {
            if (radioButton1.Checked == true)
            {
                string consulta = "select * from dbo.Spotify_Songs where track_name =" + "'" + textBox1.Text + "'" + "";
                Adaptador = new SqlDataAdapter(consulta, Conexion);
                TablaDatos = new DataTable();
                Adaptador.Fill(TablaDatos);
                dataGridView1.DataSource = TablaDatos;
            }
            else if (radioButton2.Checked == true)
            {
                string consulta = "select * from dbo.Spotify_Songs where artist_s_name =" + "'" + textBox1.Text + "'" + "";
                Adaptador = new SqlDataAdapter(consulta, Conexion);
                TablaDatos = new DataTable();
                Adaptador.Fill(TablaDatos);
                dataGridView1.DataSource = TablaDatos;
            }
            else if (radioButton3.Checked == true)
            {
                string consulta = "select * from dbo.Spotify_Songs where released_year =" + textBox1.Text + "";
                Adaptador = new SqlDataAdapter(consulta, Conexion);
                TablaDatos = new DataTable();
                Adaptador.Fill(TablaDatos);
                dataGridView1.DataSource = TablaDatos;

            }
            else
            {
                MessageBox.Show("Porfavor Seleccione Una Opbcion y llene el campo.");
            }
        }

        private void Busqueda2()
        {
            if (!string.IsNullOrWhiteSpace(textBox2.Text) && !string.IsNullOrWhiteSpace(textBox3.Text))
            {
                string consulta = "select * from dbo.Spotify_Songs where artist_s_name = '" + textBox2.Text + "' and released_year = " + textBox3.Text + " ORDER BY track_name ASC";
                Adaptador = new SqlDataAdapter(consulta, Conexion);
                TablaDatos = new DataTable();
                Adaptador.Fill(TablaDatos);
                dataGridView1.DataSource = TablaDatos;
            }
            else
            {
                MessageBox.Show("Por favor, asegúrate de llenar ambos campos.");
            }

        }

        private void Recargar()
        {
            string consulta = "select * from dbo.Spotify_Songs";
            Adaptador = new SqlDataAdapter(consulta, Conexion);
            TablaDatos = new DataTable();
            Adaptador.Fill(TablaDatos);
            dataGridView1.DataSource = TablaDatos;
            MessageBox.Show("Recargando Los Datos.");
        }

        private void ExportarExcel()
        {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    using (ExcelPackage pck = new ExcelPackage())
                    {
                        ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Sheet1");
                        for (int i = 0; i < dataGridView1.Columns.Count; i++)
                        {
                            ws.Cells[1, i + 1].Value = dataGridView1.Columns[i].HeaderText;
                        }
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            for (int j = 0; j < dataGridView1.Columns.Count; j++)
                            {
                                ws.Cells[i + 2, j + 1].Value = dataGridView1.Rows[i].Cells[j].Value?.ToString();
                            }
                        }
                        var bin = pck.GetAsByteArray();
                        File.WriteAllBytes(sfd.FileName, bin);
                    }
                    MessageBox.Show("Datos exportados exitosamente.");
                }
            }
        }

        private void ImportarDesdeExcel()
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string path = ofd.FileName;
                    string excelConnectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={path};Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                    using (OleDbConnection excelConnection = new OleDbConnection(excelConnectionString))
                    {
                        excelConnection.Open();
                        DataTable dtExcelSchema;
                        dtExcelSchema = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                        OleDbDataAdapter dataAdapter = new OleDbDataAdapter($"SELECT * FROM [{sheetName}]", excelConnection);
                        DataTable dt = new DataTable();
                        dataAdapter.Fill(dt);

                        // Insertar datos a la base de datos
                        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(Conexion))
                        {
                            bulkCopy.DestinationTableName = "dbo.Spotify_Songs";
                            bulkCopy.WriteToServer(dt);
                        }
                        MessageBox.Show("Datos importados exitosamente.");
                        Recargar(); // Recargar los datos para mostrar los cambios
                    }
                }
            }
        }
    }
}
