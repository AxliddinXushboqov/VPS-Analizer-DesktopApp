using ClosedXML.Excel;
using System.Diagnostics;
using System.Globalization;
using System.Net.Http.Json;
using System.Runtime.InteropServices;
using VPS_Analizer.Models;

namespace VPS_Analizer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        [DllImport("user32.dll")]
        public static extern void ReleaseCapture();
        [DllImport("user32.dll")]
        public static extern void SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        List<UserInfo>? accounts = null;

        private decimal? Balans = 0;
        private decimal? Equity = 0;
        private int AccountCount = 0;
        private bool sortDescending = true;

        private void panelTitleBar_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(this.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private async Task LoadAccounts(List<UserInfo> Accounts)
        {
            var result = await Task.Run(() =>
            {
                decimal totalBalans = 0;
                decimal totalEquity = 0;

                foreach (var acc in Accounts)
                {
                    totalBalans += decimal.Parse(acc.AccountBalance, CultureInfo.InvariantCulture);
                    totalEquity += decimal.Parse(acc.AccountEquity, CultureInfo.InvariantCulture);
                }

                return (Accounts, totalBalans, totalEquity);
            });

            UpdateUI(result);
        }

        void UpdateUI((List<UserInfo> Accounts, decimal totalBalans, decimal totalEquity) data)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => UpdateUI(data)));
                return;
            }

            flowLayoutPanel1.SuspendLayout();
            flowLayoutPanel1.Controls.Clear();

            foreach (var acc in data.Accounts)
            {
                var row = BuildAccountRow(acc);
                flowLayoutPanel1.Controls.Add(row);
            }

            label3.Text = $"{data.totalBalans} $";
            label4.Text = $"{data.totalEquity} $";
            label6.Text = $"{data.Accounts.Count} ta";

            flowLayoutPanel1.ResumeLayout();
        }

        private Panel BuildAccountRow(UserInfo acc)
        {
            Panel row = new Panel();
            row.Height = 50;
            row.Width = flowLayoutPanel1.Width - 25;
            row.BackColor = Color.FromArgb(45, 45, 45);
            row.Margin = new Padding(0, 2, 0, 2);

            Label lblId = new Label();
            lblId.Text = $"ID: {acc.VpsId}";
            lblId.ForeColor = Color.Gray;
            lblId.Font = new Font("Segoe UI", 8, FontStyle.Italic);
            lblId.AutoSize = true;
            lblId.Location = new Point(10, 5);
            row.Controls.Add(lblId);

            Label lblAcc = new Label();
            lblAcc.Text = acc.ClientLogin;
            lblAcc.ForeColor = Color.White;
            lblAcc.Font = new Font("Segoe UI", 12, FontStyle.Bold);
            lblAcc.AutoSize = true;
            lblAcc.Location = new Point(10, 20);

            lblAcc.Click += (s, e) =>
            {
                Clipboard.SetText(acc.ClientLogin);
                ToolTip tip = new ToolTip();
                tip.Show("Copied!", lblAcc, 0, -20, 1000);
            };
            row.Controls.Add(lblAcc);

            Panel invRect = new Panel();
            invRect.Size = new Size(12, 12);
            invRect.BackColor = acc.RobotStatus ? Color.LimeGreen : Color.Red;
            invRect.Location = new Point(130, 19);
            row.Controls.Add(invRect);

            Label lblInv = new Label();
            lblInv.Text = "Robot";
            lblInv.ForeColor = Color.White;
            lblInv.Font = new Font("Segoe UI", 8);
            lblInv.AutoSize = true;
            lblInv.Location = new Point(150, 17);
            row.Controls.Add(lblInv);

            Label lblLoad = new Label();
            lblLoad.Text = $"RAM {acc.ServerRam} %   CPU {acc.ServerCpu} %";
            lblLoad.ForeColor = Color.White;
            lblLoad.Font = new Font("Consolas", 9);
            lblLoad.AutoSize = true;
            lblLoad.Location = new Point(250, 17);
            row.Controls.Add(lblLoad);

            if (!string.IsNullOrEmpty(acc.AccountBalance))
            {
                Label lblBal = new Label();
                lblBal.Text = $"Balance {acc.AccountBalance} $   Equity {acc.AccountEquity} $";
                lblBal.ForeColor = Color.White;
                lblBal.Font = new Font("Consolas", 9);
                lblBal.AutoSize = true;
                lblBal.Location = new Point(480, 17);
                row.Controls.Add(lblBal);
            }

            Label lblExc = new Label();
            lblExc.Font = new Font("Consolas", 9);
            lblExc.AutoSize = true;
            lblExc.Location = new Point(960, 17);

            bool isTimeExpired = (DateTime.UtcNow - acc.LastCheckedTime).TotalSeconds > 50;
            string desc = acc.ProblemDescription?.Trim() ?? "";
            bool isNormal = desc == "Hammasi joyida!";
            bool isNew = desc == "Yangi Qo'shilgan!";

            if (isNew)
            {
                lblExc.Text = "Yangi Qo'shilgan!";
                lblExc.ForeColor = Color.LimeGreen;
            }
            else if (isTimeExpired)
            {
                lblExc.Text = "Aloqa ummuman yo'q!";
                lblExc.ForeColor = Color.OrangeRed;
            }
            else if (!string.IsNullOrWhiteSpace(desc) && !isNormal)
            {
                string shortText = desc.Length > 60 ? desc.Substring(0, 60) + "..." : desc;

                lblExc.Text = shortText;
                lblExc.ForeColor = Color.Red;

                ToolTip tip = new ToolTip();
                tip.SetToolTip(lblExc, desc);
            }
            else
            {
                lblExc.Text = desc;
                lblExc.ForeColor = Color.White;
            }

            row.Controls.Add(lblExc);

            Button btnView = new Button();
            btnView.Text = "Ko'rish";
            btnView.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            btnView.BackColor = Color.DodgerBlue;
            btnView.ForeColor = Color.White;
            btnView.FlatStyle = FlatStyle.Flat;
            btnView.FlatAppearance.BorderSize = 0;
            btnView.Size = new Size(70, 30);
            btnView.Location = new Point(row.Width - 85, 10);
            btnView.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            btnView.Click += (s, e) =>
            {
                string name = acc.VpsIp.Replace(':', '_');
                string rdpPath = Path.Combine(Application.StartupPath, "RDP", $"{name}.rdp");

                if (File.Exists(rdpPath))
                {
                    Clipboard.SetText(acc.VpsPassword);
                    Process.Start(new ProcessStartInfo { FileName = rdpPath, UseShellExecute = true });
                }
                else
                {
                    MessageBox.Show($"RDP fayl topilmadi:\n{rdpPath}");
                }
            };

            row.Controls.Add(btnView);

            return row;
        }

        private async void Form1_Load(object sender, EventArgs e)
        {
            button6.Enabled = false;

            _ = Task.Run(async () =>
            {
                while (true)
                {
                    // UI update - thread-safe
                    flowLayoutPanel1.SafeUpdate(panel =>
                    {
                        panel.SuspendLayout();
                        panel.Visible = false;
                        panel.Controls.Clear();
                    });

                    await LoadAllSources(); // og‘ir ish - background threadda

                    flowLayoutPanel1.SafeUpdate(panel =>
                    {
                        panel.Visible = true;
                        panel.ResumeLayout();
                    });

                    await Task.Delay(60000); // 1 daqiqa kutadi, ish tugagandan keyin
                }
            });
        }

        private void label10_Click(object sender, EventArgs e)
        {
            if (accounts == null || accounts.Count == 0)
                return;

            if (sortDescending)
            {
                accounts = accounts
                    .OrderByDescending(a => decimal.Parse(a.AccountBalance, CultureInfo.InvariantCulture) + decimal.Parse(a.AccountEquity, CultureInfo.InvariantCulture))
                    .ToList();
            }
            else
            {
                accounts = accounts
                    .OrderBy(a => decimal.Parse(a.AccountBalance, CultureInfo.InvariantCulture) + decimal.Parse(a.AccountEquity, CultureInfo.InvariantCulture))
                    .ToList();
            }

            sortDescending = !sortDescending;

            flowLayoutPanel1.Controls.Clear();
            LoadAccounts(accounts);
        }

        private void label9_Click(object sender, EventArgs e)
        {
            if (accounts == null || accounts.Count == 0)
                return;

            if (sortDescending)
            {
                accounts = accounts
                    .OrderByDescending(a => a.ServerRam + a.ServerCpu)
                    .ToList();
            }
            else
            {
                accounts = accounts
                    .OrderBy(a => a.ServerRam + a.ServerCpu)
                    .ToList();
            }

            sortDescending = !sortDescending;

            flowLayoutPanel1.Controls.Clear();
            LoadAccounts(accounts);
        }

        private void label7_Click(object sender, EventArgs e)
        {
            if (accounts == null || accounts.Count == 0)
                return;

            if (sortDescending)
            {
                accounts = accounts
                    .OrderByDescending(a => int.Parse(a.VpsId))
                    .ToList();
            }
            else
            {
                accounts = accounts
                    .OrderBy(a => int.Parse(a.VpsId))
                    .ToList();
            }

            sortDescending = !sortDescending;

            flowLayoutPanel1.Controls.Clear();
            LoadAccounts(accounts);
        }

        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
        }

        private async void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
            List<UserInfo> users = new List<UserInfo>();
            using HttpClient client = new HttpClient();
            string url = "https://vps-analizer-7a5f56f72765.herokuapp.com/api/User/AddVPS";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;

                try
                {
                    using (var workbook = new XLWorkbook(filePath))
                    {
                        var worksheet = workbook.Worksheet(1);
                        var rows = worksheet.RangeUsed().RowsUsed();

                        users.Clear();

                        foreach (var row in rows.Skip(1))
                        {
                            var id = row.Cell(1).GetValue<string>();
                            var login = row.Cell(2).GetValue<string>();
                            var vpsPassword = row.Cell(8).GetValue<string>();
                            var vpsIp = row.Cell(6).GetValue<string>();

                            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(login) || string.IsNullOrWhiteSpace(vpsPassword))
                                continue;

                            var user = new UserInfo
                            {
                                UserId = Guid.NewGuid(),
                                VpsId = id,
                                VpsIp = vpsIp,
                                VpsPassword = vpsPassword,
                                LastCheckedTime = DateTime.UtcNow,
                                ServerRam = "0",
                                ServerCpu = "0",
                                ClientLogin = login,
                                AccountBalance = "0",
                                AccountEquity = "0",
                                RobotStatus = false,
                                ProblemDescription = "Yangi Qo'shilgan!"
                            };

                            users.Add(user);
                        }
                    }

                    try
                    {
                        HttpResponseMessage response = await client.PostAsJsonAsync(url, users);

                        if (response.IsSuccessStatusCode)
                        {
                            ToolTip tip = new ToolTip();
                            Point mousePos = this.PointToClient(Cursor.Position);
                            tip.Show("Success!", this, mousePos.X + 10, mousePos.Y + 10, 1000);
                        }
                        else
                        {
                            string error = await response.Content.ReadAsStringAsync();
                            MessageBox.Show($"❌ Xatolik: {response.StatusCode}\n{error}");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"❌ Xatolik yuz berdi:\n{ex.Message}");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("❌ Xatolik: " + ex.Message);
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.flowLayoutPanel1.Controls.Clear();
            this.Form1_Load(sender, e);
        }

        private async Task LoadAllSources()
        {
            string url = "https://vps-analizer-7a5f56f72765.herokuapp.com/api/User/GetAllVPS";
            using HttpClient client = new HttpClient();
            List<UserInfo>? newAccounts = null;

            try
            {
                newAccounts = await client.GetFromJsonAsync<List<UserInfo>>(url);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Xatolik: {ex.Message}");
                return;
            }

            if (newAccounts == null)
                return;

            if (accounts != null && newAccounts.SequenceEqual(accounts))
                return;

            accounts = newAccounts;

            AccountCount = 0;
            Balans = 0;
            Equity = 0;

            foreach (var acc in accounts)
            {
                if ((DateTime.UtcNow - acc.LastCheckedTime).TotalSeconds > 50)
                {
                    acc.RobotStatus = false;
                }
            }

            accounts = accounts
                .OrderByDescending(a => (a.RobotStatus ? 0 : 1))
                .ToList();

            await LoadAccounts(accounts);
        }


        private async void button6_Click(object sender, EventArgs e)
        {
            string vpsId = textBox1.Text.Trim();

            if (string.IsNullOrEmpty(vpsId))
            {
                MessageBox.Show("Iltimos, VPS ID kiriting!");
                return;
            }

            string url = $"https://vps-analizer-7a5f56f72765.herokuapp.com/api/User/DeleteVPSById?VpsId={vpsId}";

            using (HttpClient client = new HttpClient())
            {
                try
                {
                    HttpResponseMessage response = await client.DeleteAsync(url);

                    if (response.IsSuccessStatusCode)
                    {
                        MessageBox.Show($"✅ VPS ID {vpsId} muvaffaqiyatli o‘chirildi!");
                    }
                    else
                    {
                        MessageBox.Show($"❌ Xatolik: {response.StatusCode}");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Xatolik: {ex.Message}");
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text))
                button6.Enabled = false;
            else
                button6.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
                this.WindowState = FormWindowState.Normal;
            else
                this.WindowState = FormWindowState.Maximized;
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            flowLayoutPanel1.Invalidate();
            flowLayoutPanel1.Refresh();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Normal || this.WindowState == FormWindowState.Maximized)
            {
                flowLayoutPanel1.AutoScrollPosition = new Point(0, 0);
                flowLayoutPanel1.PerformLayout();
            }

            flowLayoutPanel1.SuspendLayout();
            flowLayoutPanel1.ResumeLayout();
            flowLayoutPanel1.PerformLayout();

        }
    }

    public static class ControlExtensions
    {
        public static void SafeUpdate(this Control control, Action<Control> updateAction)
        {
            if (control.InvokeRequired)
                control.Invoke(new Action(() => updateAction(control)));
            else
                updateAction(control);
        }
    }

}