using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Mail_gönderme
{
    public partial class Form1 : Form
    {
        List<Ek> listeEk = new List<Ek>();
        int tempind;
        SmtpClient smtp = new SmtpClient();
        bool kontrol;
        public Form1()
        {
        InitializeComponent();
        }

        #region İnternet Kontrol
        public bool InternetKontrol()
        {
            try
            {
                System.Net.Sockets.TcpClient kontrol_client = new System.Net.Sockets.TcpClient("www.google.com.tr", 80);
                kontrol_client.Close();
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        public void Dogrulama()
        {
            bool kontrol = InternetKontrol(); 
                                               
            if (kontrol == true)
            {
                label1.Text = "İnternet Bağlantısı Açık...";
                label1.ForeColor = Color.Green;
                button2.Enabled = true;
            }
            else
            {
                label1.Text = "İnternet Bağlantısı Kapalı...";
                label1.ForeColor = Color.Red;
                button2.Enabled = false;
            }
        }

        #endregion

        #region Posta Gönderme
        private void button2_Click(object sender, EventArgs e)
        {

            SmtpClient sc = new SmtpClient();
            sc.Port = 587;
            sc.Host = "smtp.gmail.com";
            sc.EnableSsl = true;
            sc.Credentials = new NetworkCredential(textBox1.Text, textBox2.Text);
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress(textBox1.Text, "");
            foreach (var a in listBox1.Items)
                mail.To.Add(a.ToString());
            mail.Subject = textBox4.Text;
            mail.Body = textBox5.Text;
            mail.IsBodyHtml = true;

            if (listBox2.Items.Count > 0)//Ekleri Ekleme
               for (int j = 0; j < listBox2.Items.Count; j++)
                    for (int i = 0; i < listeEk.Count; i++)
                    {
                        if (listBox2.Items[j].ToString() == listeEk[i].dosyaAdi)
                        {
                            tempind = j;
                            mail.Attachments.Add(new Attachment(listeEk[i].dosyaYolu));
                        }
                    }
            try
            {
                sc.SendAsync(mail, (object)mail);
                kontrol = false;
            }
            catch (SmtpException ex)
            {
                MessageBox.Show("Hata : " + ex.Message);
            }
            sc.SendCompleted += Sc_SendCompleted;
            if(kontrol==false)
            {
                button1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
                button4.Enabled = false;
                button5.Enabled = false;
            }
            if(kontrol==true)
            {
                button1.Enabled = true;
                button2.Enabled = true;
                button3.Enabled = true;
                button4.Enabled = true;
                button5.Enabled = true;
            }
        }

        private void Sc_SendCompleted(object sender, AsyncCompletedEventArgs e)
        {
            MessageBox.Show("Mail Gönderildi","Bilgi");
            kontrol = true;
        }
        #endregion

        #region Gönderilecekleri Ekleme
        private void button1_Click(object sender, EventArgs e)
        {
            List<string> lst = textBox3.Text.Split(';').ToList();
            foreach(var a in lst)
            {
                listBox1.Items.Add(a.ToString());
            }
        }
        #endregion

        #region Ek Ekleme
        private void button4_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog file = new OpenFileDialog();
            //file.Filter = "Tüm dosyalar |*.xlsx| Excel Dosyası|*.xls";  
            file.CheckFileExists = false;
            file.Title = "Ek Dosyası Seçiniz..";
            
            if(file.ShowDialog() == DialogResult.OK)
            {
                Ek temp = new Ek();
                temp.dosyaAdi = file.SafeFileName;
                temp.dosyaYolu = file.FileName;
                listeEk.Add(temp);
            }
            foreach (var a in listeEk)
                listBox2.Items.Add(a.dosyaAdi);
        }
        #endregion

        #region Ek Silme
        private void button3_Click(object sender, EventArgs e)
        {
            if(listBox2.SelectedItems.Count>0)
            {
                string indexS = listBox2.SelectedItem.ToString();
                for (int a = 0; a < listBox2.Items.Count; a++)
                    if (listBox2.Items[a].ToString() == indexS)
                        tempind = a;
                listBox2.Items.RemoveAt(tempind);
                listeEk.RemoveAt(tempind);
            }
            
        }
        #endregion

        #region Kişi silme
        private void button5_Click(object sender, EventArgs e)
        {
            string indexS = listBox1.SelectedItem.ToString();
            for (int a = 0; a < listBox1.Items.Count; a++)
                if (listBox1.Items[a].ToString() == indexS)
                    tempind = a;
            listBox1.Items.RemoveAt(tempind);
        }
        #endregion
    }
}
