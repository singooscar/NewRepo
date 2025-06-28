using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

// Note: Ensure you have the necessary using directives for your project.

namespace KARDEX_FUNCIONAL_BDD
{
    public partial class Form1 : Form
    {
        string connectionString = ConfigurationManager.ConnectionStrings["Miconexion"].ConnectionString;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("SELECT IdProducto, NombreProducto FROM Productos", conn);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    comboBox1.Items.Add(new ComboBoxItem
                    {
                        Text = dr["NombreProducto"].ToString(),
                        Value = Convert.ToInt32(dr["IdProducto"])
                    });
                }
                comboBox2.Items.Add("Entrada");
                comboBox2.Items.Add("Salida");
            }

            dataGridView1.ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = false;
            dataGridView1.AllowUserToOrderColumns = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    MessageBox.Show("Conexión exitosa.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al conectar: " + ex.Message);
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selected = (ComboBoxItem)comboBox1.SelectedItem;
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                SqlDataAdapter da = new SqlDataAdapter(
                    "SELECT FechaMovimiento, TipoMovimiento, Cantidad, PrecioUnitario, Total, StockResultante FROM Movimientos WHERE IdProducto = @id ORDER BY FechaMovimiento",
                    conn
                );
                da.SelectCommand.Parameters.AddWithValue("@id", selected.Value);
                DataTable dt = new DataTable();
                da.Fill(dt);

                DataTable kardex = GenerarKardex(dt);
                dataGridView1.DataSource = kardex;

                if (kardex.Rows.Count > 0)
                {
                    int stockInicial = Convert.ToInt32(kardex.Rows[0]["Saldo"]);
                    int stockFinal = Convert.ToInt32(kardex.Rows[kardex.Rows.Count - 1]["Saldo"]);
                    label5.Text = "Stock actual: " + stockFinal;
                    label6.Text = "Stock inicial: " + stockInicial;
                }
                else
                {
                    label5.Text = "Stock actual: 0";
                    label6.Text = "Stock inicial: Sin movimientos";
                }

                // Mostrar el precio predeterminado
                SqlCommand precioCmd = new SqlCommand("SELECT PrecioUnitario FROM Productos WHERE IdProducto = @id", conn);
                precioCmd.Parameters.AddWithValue("@id", selected.Value);

                object precioResult = precioCmd.ExecuteScalar();
                if (precioResult != null)
                {
                    decimal precio = Convert.ToDecimal(precioResult);
                    label8.Text = "Precio: $" + precio.ToString("0.00");
                    textBox2.Text = precio.ToString("0.00"); // Opcional
                }
                else
                {
                    label8.Text = "Precio: N/A";
                    textBox2.Text = "";
                }
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem == null || comboBox2.SelectedItem == null)
            {
                MessageBox.Show("Por favor, selecciona un producto y el tipo de movimiento.");
                return;
            }

            int cantidad;
            if (!int.TryParse(textBox1.Text.Trim(), out cantidad) || cantidad <= 0)
            {
                MessageBox.Show("Cantidad inválida. Debe ser un número entero mayor que cero.");
                return;
            }

            decimal precio;
            if (!decimal.TryParse(textBox2.Text.Trim(), out precio) || precio <= 0)
            {
                MessageBox.Show("Precio inválido. Debe ser un número decimal mayor que cero.");
                return;
            }

            var producto = (ComboBoxItem)comboBox1.SelectedItem;
            string tipo = comboBox2.SelectedItem.ToString();

            int stockActual;
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("SELECT StockActual FROM Productos WHERE IdProducto = @id", conn);
                cmd.Parameters.AddWithValue("@id", producto.Value);

                object result = cmd.ExecuteScalar();
                if (result == null || !int.TryParse(result.ToString(), out stockActual))
                {
                    MessageBox.Show("No se pudo obtener el stock actual desde la base de datos.");
                    return;
                }
            }

            int nuevoStock = tipo == "Entrada" ? stockActual + cantidad : stockActual - cantidad;

            if (nuevoStock < 0)
            {
                MessageBox.Show("No hay suficiente stock para realizar la salida.");
                return;
            }

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlTransaction trx = conn.BeginTransaction();
                try
                {
                    SqlCommand insert = new SqlCommand(
                        "INSERT INTO Movimientos (IdProducto, FechaMovimiento, TipoMovimiento, Cantidad, PrecioUnitario, Total, StockResultante) " +
                        "VALUES (@id, @fecha, @tipo, @cant, @precio, @total, @saldo)", conn, trx);

                    insert.Parameters.AddWithValue("@id", producto.Value);
                    insert.Parameters.AddWithValue("@fecha", dateTimePicker1.Value);
                    insert.Parameters.AddWithValue("@tipo", tipo);
                    insert.Parameters.AddWithValue("@cant", cantidad);
                    insert.Parameters.AddWithValue("@precio", precio);
                    insert.Parameters.AddWithValue("@total", cantidad * precio);
                    insert.Parameters.AddWithValue("@saldo", nuevoStock);
                    insert.ExecuteNonQuery();

                    SqlCommand update = new SqlCommand("UPDATE Productos SET StockActual = @nuevo WHERE IdProducto = @id", conn, trx);
                    update.Parameters.AddWithValue("@nuevo", nuevoStock);
                    update.Parameters.AddWithValue("@id", producto.Value);
                    update.ExecuteNonQuery();

                    trx.Commit();
                    MessageBox.Show("Movimiento registrado correctamente.");
                    comboBox1_SelectedIndexChanged(null, null);
                }
                catch (Exception ex)
                {
                    trx.Rollback();
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        private DataTable GenerarKardex(DataTable movimientos)
        {
            DataTable kardex = new DataTable();
            kardex.Columns.Add("Fecha", typeof(DateTime));
            kardex.Columns.Add("Tipo", typeof(string));
            kardex.Columns.Add("Entrada", typeof(int));
            kardex.Columns.Add("Salida", typeof(int));
            kardex.Columns.Add("Precio Unitario", typeof(decimal));
            kardex.Columns.Add("Total", typeof(decimal));
            kardex.Columns.Add("Saldo", typeof(int));

            int saldo = 0;

            foreach (DataRow row in movimientos.Rows)
            {
                string tipo = row["TipoMovimiento"].ToString();
                int cantidad = Convert.ToInt32(row["Cantidad"]);
                decimal precio = Convert.ToDecimal(row["PrecioUnitario"]);
                decimal total = Convert.ToDecimal(row["Total"]);

                int entrada = tipo == "Entrada" ? cantidad : 0;
                int salida = tipo == "Salida" ? cantidad : 0;

                saldo += entrada - salida;

                kardex.Rows.Add(
                    Convert.ToDateTime(row["FechaMovimiento"]),
                    tipo,
                    entrada,
                    salida,
                    precio,
                    total,
                    saldo
                );
            }

            return kardex;
        }

        public class ComboBoxItem
        {
            public string Text { get; set; }
            public int Value { get; set; }  
            public override string ToString() => Text;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e) { }
        private void textBox1_TextChanged(object sender, EventArgs e) { }
        private void textBox2_TextChanged(object sender, EventArgs e) { }
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e) { }
        private void label4_Click(object sender, EventArgs e) { }
        private void label5_Click(object sender, EventArgs e) { }
        private void label6_Click(object sender, EventArgs e) { }
        private void label7_Click(object sender, EventArgs e) { }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e) { }
        private void label9_Click(object sender, EventArgs e) { }

        private void button3_Click(object sender, EventArgs e)
        {
            ExportarKardexAExcel();


        }

        private void ExportarKardexAExcel()
        {
            if (dataGridView1.DataSource == null)
            {
                MessageBox.Show("No hay datos para exportar.");
                return;
            }

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet hoja = workbook.Sheets[1];

            try
            {
                // Título
                hoja.Range["A1:G1"].Merge();
                hoja.Cells[1, 1] = "KARDEX";
                hoja.Cells[1, 1].Font.Bold = true;
                hoja.Cells[1, 1].Font.Size = 16;
                hoja.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // Información general
                hoja.Cells[2, 1] = "Artículo:";
                hoja.Cells[2, 2] = comboBox1.Text;

                hoja.Cells[3, 1] = "Método:";
                hoja.Cells[3, 2] = "Promedio ponderado";

                hoja.Cells[2, 5] = "Existencia mínima:";
                hoja.Cells[2, 6] = "60";

                hoja.Cells[3, 5] = "Existencia máxima:";
                hoja.Cells[3, 6] = "495";

                // Encabezados
                string[] headers = { "Fecha", "Tipo", "Entrada", "Salida", "Precio Unitario", "Total", "Saldo" };
                for (int i = 0; i < headers.Length; i++)
                {
                    hoja.Cells[5, i + 1] = headers[i];
                    hoja.Cells[5, i + 1].Font.Bold = true;
                    hoja.Cells[5, i + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                }

                // Llenar datos
                DataTable dt = (DataTable)dataGridView1.DataSource;
                int filaExcel = 6;

                foreach (DataRow fila in dt.Rows)
                {
                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        hoja.Cells[filaExcel, col + 1] = fila[col];
                    }
                    filaExcel++;
                }

                // Inventario final (última fila)
                hoja.Cells[filaExcel + 1, 1] = "Inventario Final:";
                hoja.Cells[filaExcel + 1, 2] = dt.Rows[dt.Rows.Count - 1]["Saldo"].ToString();
                hoja.Cells[filaExcel + 1, 2].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                hoja.Cells[filaExcel + 1, 1].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                hoja.Cells[filaExcel + 1, 1].Font.Bold = true;

                // Ajustar columnas
                hoja.Columns.AutoFit();

                // Mostrar Excel
                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al exportar: " + ex.Message);
                workbook.Close(false);
                excelApp.Quit();
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }
    }
}
