using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Limilabs.Mail;
using Limilabs.Client.IMAP;
using System.Net.Mail;
namespace WindowsFormsApplication3
{
    public partial class Mail : Form
    {
        private Imap imap;
        private IMail imail;
        private DataTable table;
        private long idmail;
        public Mail()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            imap = new Imap();
            string m = textBox1.Text;
            int l = m.Length;
            int ll = textBox2.Text.Length;
            if (l > 8  )
            {
                if (m.Substring(m.Length - 9) == "gmail.com")
                {
                    imap.ConnectSSL("imap.gmail.com", 993);
                    try
                    {
                        table = new DataTable();
                        table.Columns.Add("IDMail", typeof(string));
                        table.Columns.Add("Subject", typeof(string));
                        table.Columns.Add("Date", typeof(string));
                        table.Columns.Add("From", typeof(string));
                        imap.Login(textBox1.Text, textBox2.Text);
                        MessageBox.Show("Почта открыта");
                    }
                    catch
                    {
                        MessageBox.Show("Не получился открыт почту");
                    }
                }


                if (m.Substring(m.Length - 7) == "mail.ru")
                {
                    imap.ConnectSSL("imap.mail.ru", 993);

                    try
                    {
                        table = new DataTable();
                        table.Columns.Add("IDMail", typeof(string));
                        table.Columns.Add("Subject", typeof(string));
                        table.Columns.Add("Date", typeof(string));
                        table.Columns.Add("From", typeof(string));
                        imap.Login(textBox1.Text, textBox2.Text);
                        MessageBox.Show("Почта открыта");
                    }
                    catch
                    {
                        MessageBox.Show("Не получился открыть почту");
                    }
                }





                if (m.Substring(m.Length - 9) == "yandex.ru")
                {
                    imap.ConnectSSL("imap.yandex.ru", 993);

                    try
                    {
                        table = new DataTable();
                        table.Columns.Add("IDMail", typeof(string));
                        table.Columns.Add("Subject", typeof(string));
                        table.Columns.Add("Date", typeof(string));
                        table.Columns.Add("From", typeof(string));
                        imap.Login(textBox1.Text, textBox2.Text);
                        MessageBox.Show("Почта открыта");
                    }
                    catch
                    {
                        MessageBox.Show("Не получился открыт почту");
                    }
                }
            }
            else  
            {
                MessageBox.Show("Неверный логин или пароль");
            }
            
           
        }
        private void button2_Click(object sender, EventArgs e)
        {

            imap.SelectInbox();
            List<long>  uids = imap.SearchFlag(Flag.All);
            
             int   i = uids.Count;
            int j = 0;
            foreach (long uid in uids)
            {
                if (j < i)
                {
                    try {
                        byte[] eml = imap.GetHeadersByUID(uid);
                        imail = new MailBuilder().CreateFromEml(eml);
                        TimeSpan t = dateTimePicker1.Value - imail.Date.Value;
                        TimeSpan t2 = dateTimePicker2.Value - imail.Date.Value;
                        if (t.Days <= 0 && t2.Days >= 0)
                        {
                            DataRow row = table.NewRow();
                            row["IDMail"] = uid.ToString();
                            row["Subject"] = imail.Subject;
                            row["Date"] = imail.Date.Value.ToString("dd/MM/yyyy");
                            row["From"] = imail.From.ToString();
                            table.Rows.Add(row);
                            table.AcceptChanges();
                            
                        }
                        else j--;
                    
                    }
                    catch
                    {
                        j++;
                    }
                    }
                else break;
            }
            dataGridView1.DataSource = table;
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
            {
                if (e.ColumnIndex == 0)
                {
                    string id = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    byte [] eml = imap.GetMessageByUID(long.Parse(id));
                    imail = new MailBuilder().CreateFromEml(eml);
                    richTextBox1.Text = imail.Text;
                    idmail = long.Parse(id);

                }
                else
                {
                    idmail = 0;
                }
            }
            else
            {
                idmail = 0;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try {
                if (idmail != 0)
                {
                    imap.DeleteMessageByUID(idmail);
                    foreach (DataRow row in table.Rows)
                    {
                        if (idmail == long.Parse(row["IDMail"].ToString())) 
                        {
                            table.Rows.Remove(row);
                            table.AcceptChanges();
                            break;
                        }
                    }
                    dataGridView1.Rows.Clear();
                    dataGridView1.DataSource = table;
                }
            }
            catch
            {
                MessageBox.Show("Удаленно");
            }
            }

        private void button5_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox5.Text = openFileDialog1.FileName.ToString();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            int lenthg1 = textBox1.Text.Length;
            int lenthg = textBox4.Text.Length;
            string m = textBox4.Text;
            if (lenthg1 > 9)
            {
               
                    if (m.Substring(m.Length - 7) == "mail.ru")
                    {
                        try
                    {
                        SmtpClient client = new SmtpClient("smtp.mail.ru", 587);
                        MailMessage message = new MailMessage();
                        message.From = new MailAddress(textBox1.Text);
                        message.To.Add(textBox4.Text);
                        message.Body = textBox7.Text;
                        message.Subject = textBox6.Text;
                        client.UseDefaultCredentials = false;
                        client.EnableSsl = true;
                        if (textBox5.Text != "")
                        {
                            message.Attachments.Add(new Attachment(textBox5.Text));

                        }
                        client.Credentials = new System.Net.NetworkCredential(textBox1.Text, textBox2.Text);
                        client.Send(message);
                        message = null;
                        MessageBox.Show("Сообщения отправленно");
                    }
                    catch
                    {
                        MessageBox.Show("Не получился отправить сообщения");
                    }
                }
               
                
                    if (m.Substring(m.Length - 9) == "gmail.com")
                    {
                        try
                        {
                            SmtpClient client = new SmtpClient("smtp.gmail.com", 587);
                            MailMessage message = new MailMessage();
                            message.From = new MailAddress(textBox1.Text);
                            message.To.Add(textBox4.Text);
                            message.Body = textBox7.Text;
                            message.Subject = textBox6.Text;
                            client.UseDefaultCredentials = false;
                            client.EnableSsl = true;
                            if (textBox5.Text != "")
                            {
                                message.Attachments.Add(new Attachment(textBox5.Text));

                            }
                            client.Credentials = new System.Net.NetworkCredential(textBox1.Text, textBox2.Text);
                            client.Send(message);
                            message = null;
                            MessageBox.Show("Сообщения отправленно");
                        }
                        catch
                        {
                            MessageBox.Show("Не получился отправить сообщения");
                        }
                    }
                
                
              
                    if (m.Substring(m.Length - 9) == "yandex.ru")
                    {
                        try
                        {
                            SmtpClient client = new SmtpClient("smtp.yandex.ru", 587);
                            MailMessage message = new MailMessage();
                            message.From = new MailAddress(textBox1.Text);
                            message.To.Add(textBox4.Text);
                            message.Body = textBox7.Text;
                            message.Subject = textBox6.Text;
                            client.UseDefaultCredentials = false;
                            client.EnableSsl = true;
                            if (textBox5.Text != "")
                            {
                                message.Attachments.Add(new Attachment(textBox5.Text));

                            }
                            client.Credentials = new System.Net.NetworkCredential(textBox1.Text, textBox2.Text);
                            client.Send(message);
                            message = null;
                            MessageBox.Show("Сообщения отправленно");
                        }
                        catch
                        {
                            MessageBox.Show("Не получился отправить сообщения");
                        }
                    }
                }
        
            else
            {
                MessageBox.Show("Неверный логин или пароль ");
            }
        }
    }
}