using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MimeKit;
using OfficeOpenXml;
using System.IO;
using MailKit.Net.Smtp;
using MailKit.Security;

namespace AutoMail
{
    public partial class Form1 : Form
    {
        private string userEmail;
        private string userPassword;

        public Form1()
        {
            InitializeComponent();
        }

        // Xử lý đăng nhập
        private void btnLogin_Click(object sender, EventArgs e)
        {
            userEmail = txtEmail.Text;
            userPassword = txtPassword.Text;

            if (string.IsNullOrEmpty(userEmail) || string.IsNullOrEmpty(userPassword))
            {
                MessageBox.Show("Vui lòng nhập email và mật khẩu.");
                return;
            }

            lblResult.Text = "Đăng nhập thành công!";
        }

        // Đọc dữ liệu từ file Excel
        private List<CreatorInfo> ReadExcelFile(string filePath)
        {
            List<CreatorInfo> creators = new List<CreatorInfo>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Kiểm tra xem workbook có ít nhất 1 sheet không
                if (package.Workbook.Worksheets.Count == 0)
                {
                    MessageBox.Show("Không có sheet nào trong workbook.");
                    return creators;
                }

                // In ra danh sách các sheet để kiểm tra
                for (int i = 0; i < package.Workbook.Worksheets.Count; i++)
                {
                    Console.WriteLine($"Sheet {i}: {package.Workbook.Worksheets[i].Name}");
                }

                // Chọn sheet đầu tiên
                var worksheet = package.Workbook.Worksheets[0];

                // Kiểm tra dimension (số hàng và cột) của worksheet
                if (worksheet.Dimension == null)
                {
                    MessageBox.Show("Sheet không có dữ liệu.");
                    return creators;
                }

                int rowCount = worksheet.Dimension.Rows;
                int columnCount = worksheet.Dimension.Columns;

                // Kiểm tra nếu có dữ liệu trong sheet
                if (rowCount < 2) // Không có dữ liệu (chỉ có tiêu đề)
                {
                    MessageBox.Show("Không có dữ liệu trong sheet.");
                    return creators;
                }

                // Đọc dữ liệu từ các dòng và cột
                for (int row = 2; row <= rowCount; row++) // Bắt đầu từ dòng 2 vì dòng 1 là tiêu đề
                {
                    var creator = new CreatorInfo
                    {
                        EmailAddress = worksheet.Cells[row, 1].Text,  // Email Address
                        CC = worksheet.Cells[row, 2].Text,            // CC
                        PIC = worksheet.Cells[row, 3].Text,           // PIC
                        UID = worksheet.Cells[row, 4].Text,           // UID
                        TiktokID = worksheet.Cells[row, 5].Text,      // Tiktok ID
                        Address = worksheet.Cells[row, 6].Text,       // Địa chỉ
                        Phone = worksheet.Cells[row, 7].Text,         // SDT
                        EmailCreator = worksheet.Cells[row, 8].Text,  // Email Creator
                        ContractStatus = worksheet.Cells[row, 9].Text, // HĐ
                        Time = worksheet.Cells[row, 10].Text,         // Time
                        LinkTiktok = worksheet.Cells[row, 11].Text,   // Link Tiktok
                        Melive = worksheet.Cells[row, 12].Text,       // Melive
                        Creator = worksheet.Cells[row, 13].Text,      // Creator
                        MergeStatus = worksheet.Cells[row, 14].Text   // Merge status
                    };

                    creators.Add(creator);
                }
            }

            return creators;
        }


        // Gửi email sử dụng MimeKit và MailKit
        private void SendEmailToCreators(List<CreatorInfo> creators)
        {
            try
            {
                using (var smtpClient = new SmtpClient())
                {
                    // Kết nối đến Gmail và sử dụng STARTTLS để bảo mật kết nối
                    smtpClient.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);

                    // Đăng nhập vào tài khoản Gmail
                    smtpClient.Authenticate(userEmail, userPassword);

                    foreach (var creator in creators)
                    {
                        var message = new MimeMessage();
                        message.From.Add(new MailboxAddress("Melive", userEmail)); // Chỉnh lại constructor đúng cách
                        message.To.Add(new MailboxAddress(creator.TiktokID, creator.EmailCreator)); // Đảm bảo truyền đúng tham số

                        var subject = $"[Melive MCN-Onboard] Hợp tác MCN Tiktok Shop với kênh - {creator.TiktokID}";
                        var body = $@"
Hi team,

Em gửi thông tin onboard creator/kênh: {creator.TiktokID}

Thông tin đánh giá sơ bộ về creator/kênh:
Creator phù hợp mảng: Mix Cat
Hình thức hợp tác mong muốn: Hợp tác Short video, livestream.
Vấn đề cần chú ý trong lúc làm việc: Hỗ trợ sản phẩm mẫu, voucher, quảng cáo.

Email: {creator.EmailCreator}
Địa chỉ: {creator.Address}
SĐT: {creator.Phone}

Thông tin hợp tác:
Handle name: {creator.TiktokID}
ID Tiktok: {creator.TiktokID}
Link Tiktok: https://www.tiktok.com/@{creator.TiktokID}
UID: {creator.UID}
Hình thức hợp tác: Partnership
Status: {creator.MergeStatus}
Thời gian link net: 01 năm, ngày bắt đầu link: {DateTime.Now.ToString("dd/MM/yyyy")}
Phần trăm chia sẻ: Melive, Creator
Thời hạn chia sẻ: chia sẻ toàn kênh, ko chia sẻ case by case

Best regards,
";

                        var textPart = new TextPart("plain")
                        {
                            Text = body
                        };

                        message.Subject = subject;
                        message.Body = textPart;

                        smtpClient.Send(message);
                    }

                    MessageBox.Show("Email đã được gửi thành công!");
                    smtpClient.Disconnect(true); // Ngắt kết nối sau khi gửi email
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi gửi email: " + ex.Message);
            }
        }

        // Chọn file Excel và gửi email
        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    var creators = ReadExcelFile(filePath);

                    // Bạn có thể kiểm tra và hiển thị kết quả đọc dữ liệu từ Excel
                    MessageBox.Show($"Đã đọc {creators.Count} creator từ file Excel.");
                }
            }
        }


        // Lớp chứa thông tin của creator
        public class CreatorInfo
        {
            public string EmailAddress { get; set; }
            public string CC { get; set; }
            public string PIC { get; set; }
            public string UID { get; set; }
            public string TiktokID { get; set; }
            public string Address { get; set; }
            public string Phone { get; set; }
            public string EmailCreator { get; set; }
            public string ContractStatus { get; set; }
            public string Time { get; set; }
            public string LinkTiktok { get; set; }
            public string Melive { get; set; }
            public string Creator { get; set; }
            public string MergeStatus { get; set; }
        }

    }
}
