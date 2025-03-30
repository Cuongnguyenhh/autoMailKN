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

        // Handle login
        private void btnLogin_Click(object sender, EventArgs e)
        {
            userEmail = txtEmail.Text;
            userPassword = txtPassword.Text;

            if (string.IsNullOrEmpty(userEmail) || string.IsNullOrEmpty(userPassword))
            {
                MessageBox.Show("Please enter email and password.");
                return;
            }

            try
            {
                using (var smtpClient = new MailKit.Net.Smtp.SmtpClient())
                {
                    // Connect to the Gmail SMTP server
                    smtpClient.Connect("smtp.gmail.com", 587, MailKit.Security.SecureSocketOptions.StartTls);

                    // Authenticate using the user's email account
                    smtpClient.Authenticate(userEmail, userPassword); // Use an app password if necessary

                    // If the connection is successful, display a success message
                    lblResult.Text = "Login successful!";
                    smtpClient.Disconnect(true);  // Disconnect after successful authentication
                }
            }
            catch (MailKit.Security.AuthenticationException ex)
            {
                // Login error due to incorrect information or blocked account
                MessageBox.Show("Login error: " + ex.Message);
            }
            catch (Exception ex)
            {
                // Catch any other errors
                MessageBox.Show("Login error: " + ex.Message);
            }
        }

        // Read data from an Excel file
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

                // Lấy sheet đầu tiên
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                // Kiểm tra nếu worksheet hợp lệ
                if (worksheet == null)
                {
                    MessageBox.Show("Không tìm thấy sheet hợp lệ.");
                    return creators;
                }

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
                        EmailCreator = worksheet.Cells[row, 1].Text,  // Email Creator
                        ContractStatus = worksheet.Cells[row, 8].Text, // HĐ
                        Time = worksheet.Cells[row, 9].Text,         // Time
                        LinkTiktok = worksheet.Cells[row, 10].Text,   // Link Tiktok
                        Melive = worksheet.Cells[row, 11].Text,       // Melive
                        Creator = worksheet.Cells[row, 12].Text,      // Creator
                        MergeStatus = worksheet.Cells[row, 13].Text   // Merge status
                    };

                    creators.Add(creator);
                }
            }

            return creators;
        }

        // Send email using MimeKit and MailKit
        private void SendEmailToCreators(List<CreatorInfo> creators)
        {
            try
            {
                using (var smtpClient = new SmtpClient())
                {
                    // Connect to Gmail and use STARTTLS for secure connection
                    smtpClient.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);

                    // Authenticate using the user's Gmail account
                    smtpClient.Authenticate(userEmail, userPassword);

                    foreach (var creator in creators)
                    {
                        var message = new MimeMessage();
                        message.From.Add(new MailboxAddress("Melive", userEmail)); // Correct constructor usage
                        message.To.Add(new MailboxAddress(creator.TiktokID, creator.EmailCreator)); // Ensure correct parameters

                        var subject = $"[Melive MCN-Onboard] MCN Tiktok Shop partnership with channel - {creator.TiktokID}";
                        var body = $@"
Hi team,

I'm sending the onboarding information for creator/channel: {creator.TiktokID}

Preliminary evaluation of the creator/channel:
----------------------------------------------------
- Creator fits in category:           Mix Cat
- Desired collaboration form:         Short video, livestream.
- Issues to consider during the work: Provide sample products, vouchers, advertising.

Contact Information:
----------------------------------------------------
- Email: {creator.EmailCreator}
- Address: {creator.Address}
- Phone: {creator.Phone}

Partnership information:
----------------------------------------------------
- Handle name: {creator.TiktokID}
- Tiktok ID: {creator.TiktokID}
- Tiktok Link: https://www.tiktok.com/@{creator.TiktokID}
- UID: {creator.UID}
- Collaboration type:                 Partnership
- Status:                            {creator.MergeStatus}
- Link duration:                     1 year
- Link start date:                   {DateTime.Now.ToString("dd/MM/yyyy")}
- Share percentage:                  Melive, Creator
- Share duration:                    Full channel share, not case by case

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

                    MessageBox.Show("Emails sent successfully!");
                    smtpClient.Disconnect(true); // Disconnect after sending emails
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error sending email: " + ex.Message);
            }
        }

        // Choose an Excel file and send email
        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    var creators = ReadExcelFile(filePath);

                    // You can check and display the result of reading data from the Excel file
                    MessageBox.Show($"Read {creators.Count} creators from the Excel file.");

                    SendEmailToCreators(creators);
                }
            }
        }

        // Class to store creator information
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
