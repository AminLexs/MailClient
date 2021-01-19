using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Net.Sockets;
using System.Net.Security;
using System.Xml.Serialization;
using System.IO;
using System.Text.RegularExpressions;
using MailClient.Properties.client;
using MailClient.Properties.server;
using Joshi.Utils.Imap;
using OpenPop.Mime;
using OpenPop.Pop3;
using Message = OpenPop.Mime.Message;

namespace MailClient
{

    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            groupBox3.Hide();
            bNext.Hide();
            bPrev.Hide();
            gsettings.Hide();
            Directory.CreateDirectory("Files");
            Directory.CreateDirectory("Saves");
            saveFileDialog1.InitialDirectory = Environment.CurrentDirectory + "\\Saves";
            openFileDialog2.InitialDirectory = Environment.CurrentDirectory + "\\Saves";
            openFileDialog2.Filter= "Файл книги электронных почт (*.bxml)|*.bxml|Файл шаблонов (*.txml)|*.txml|Файл рассылки (*.mxml)|*.mxml|All files (*.*)|*.*";
        }
        int page,portpop=995,portsmtp=25;
        string serverpop = "", serversmtp = "";
        Thread myThread;
        List<MailingObject> listmailng = new List<MailingObject>();
        List<Template> listtemplate = new List<Template>();
        private void ShowMessages(int indexstart, int indexfinish)
        {
            messages.Items.Clear();
            dataGridView1.Rows.Clear();
            using (Pop3Client client = new Pop3Client())
            {
                if (String.IsNullOrEmpty(serverpop.Trim()))
                {
                    // Подключение к серверу
                    string servname = (tsender.Text.Split(new char[] { '@' }))[1];
                    try { client.Connect("pop." + servname, portpop, true); } catch (Exception exept) { MessageBox.Show("Проверьте порт и сервер POP3. " + exept.Message, "Неверные данные", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                }
                else try { client.Connect(serverpop, portpop, true); } catch (Exception exept) { MessageBox.Show("Проверьте порт и сервер POP3. " + exept.Message, "Неверные данные", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                // Аутентификация (проверка логина и пароля)
                client.Authenticate(tsender.Text, tpassword.Text, AuthenticationMethod.UsernameAndPassword);

                if (client.Connected)
                {

                    // Сообщения нумеруются от 1 до messageCount включительно
                    // Другим языком нумерация начинается с единицы
                    // Большинство серверов присваивают новым сообщениям наибольший номер (чем меньше номер тем старее сообщение)
                    // Т.к. цикл начинается с messageCount, то последние сообщения должны попасть в начало списк

                    progressBar1.Maximum = indexstart - indexfinish;
                    int row = 0;
                    for (int i = indexstart; i > indexfinish; i--)
                    {
                        progressBar1.Value = indexstart - i;
                        Message message = client.GetMessage(i);
                        string subject, date, from, body;
                        try { subject = message.Headers.Subject; } catch { subject = ""; } //заголовок
                        try { date = message.Headers.Date.ToString(); } catch { date = ""; } //Дата/Время
                        try { from = message.Headers.From.ToString(); } catch { from = ""; }//от кого
                        try { body = ""; } catch { body = ""; }

                        // ищем первую плейнтекст версию в сообщении
                        MessagePart mpPlain = message.FindFirstPlainTextVersion();

                        if (mpPlain != null)
                        {
                            Encoding enc = mpPlain.BodyEncoding;
                            body = enc.GetString(mpPlain.Body); //получаем текст сообщения
                        }

                        var att = message.FindAllAttachments();
                        dataGridView1.Rows.Add();
                        dataGridView1.Rows[row].Cells[2].Value = (new OneMessage(new string[] { subject, date, from, body, i.ToString() }, att));
                        dataGridView1.Rows[row].Cells[1].Value = false;
                        dataGridView1.Rows[row].Cells[0].Value = i;
                        row++;
                        messages.Items.Add(new OneMessage(new string[] { subject, date, from, body, i.ToString() }, att));
                    }
                    progressBar1.Value = 0;

                }
            }
        }
        private void Auth_Click(object sender, EventArgs e)
        {
            //System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = false;
            page = 0;
            try
            {
                tabs.Enabled = true;
                using (Pop3Client client = new Pop3Client())
                {
                    // Подключение к серверу

                    if (String.IsNullOrEmpty(serverpop.Trim()))
                    {
                        // Подключение к серверу
                        string servname = (tsender.Text.Split(new char[] { '@' }))[1];
                        try { client.Connect("pop." + servname, portpop, true); } catch (Exception exept) { MessageBox.Show("Проверьте порт и сервер POP3. " + exept.Message, "Неверные данные", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                    }
                    else try { client.Connect(serverpop, portpop, true); } catch (Exception exept) { MessageBox.Show("Проверьте порт и сервер POP3. " + exept.Message, "Неверные данные", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                    // Аутентификация (проверка логина и пароля)
                    try
                    {
                        client.Authenticate(tsender.Text, tpassword.Text, AuthenticationMethod.UsernameAndPassword);
                    }
                    catch (Exception exept)
                    {

                        tabs.Enabled = false;
                        MessageBox.Show("Проверьте логин и пароль. " + exept.Message, "Неверные данные", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    page++;
                    if (client.Connected)
                    {
                        int messageCount = client.GetMessageCount();
                        numericUpDown1.Maximum = messageCount / 100;
                        if (messageCount > 100)
                        {
                            bPrev.Hide();
                            bNext.Show();
                            ShowMessages(messageCount, messageCount - 100);
                        }
                        else ShowMessages(messageCount, 0);


                        groupBox1.Hide();
                        groupBox3.Show();
                        lemail.Text = tsender.Text;
                    }

                }
                gsettings.Hide();
            }
            catch { tabs.Enabled = false; }
            if (listmailng.Count != 0)
                if (DialogResult.No == MessageBox.Show("Сохранить предыдущие рассылки?", "Рассылка", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
                {
                    listmailng.Clear();
                }
            myThread = new Thread(new ThreadStart(Mailing));
            myThread.Start(); // запускаем поток

        }
        private void Mailing()
        {
            var server = new Server(25, new AuthorizedClient(tsender.Text, tpassword.Text));
            for (; ; )
            {
                DateTime datetime = DateTime.Now;
                foreach (var mailing in listmailng)
                {
                    if (mailing.Hours == datetime.Hour && mailing.Minutes == datetime.Minute)
                    {
                        foreach (var email in mailing.ListEmails) {
                            Client receiver = new Client(email);
                            var text = mailing.Template.Body; var subj = mailing.Template.Subject;
                            server.SendMessage(receiver, text, subj, mailing.Template.Attaches);
                        }
                    }
                }
                Thread.Sleep(60000);
            }
        }

        private void Bsend_Click(object sender, EventArgs e)
        {
            var server = new Server(portsmtp, new AuthorizedClient(tsender.Text, tpassword.Text));
            Client receiver = new Client(treceiver.Text);
            var text = messageBox.Text; var subj = tsubject.Text;

            server.SendMessage(receiver, text, subj, ListFiles.Items.Cast<object>().ToArray());
            messageBox.Clear();
            tsubject.Clear();
            ListFiles.Items.Clear();
        }
        private void Battach_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (!ListFiles.Items.Contains(openFileDialog1.FileName))
                    ListFiles.Items.Add(openFileDialog1.FileName);
            }

        }

        private void ListFiles_DoubleClick(object sender, EventArgs e)
        {
            if (ListFiles.SelectedIndex != -1)
                ListFiles.Items.RemoveAt(ListFiles.SelectedIndex);
        }



        private void Messages_DoubleClick(object sender, EventArgs e)
        {
            if (messages.SelectedIndex != -1)
            {
                OneMessage message = (OneMessage)messages.SelectedItem;
                richTextBox1.Text = "Тема: " + message.Subject + Environment.NewLine +
                    "Дата: " + message.Date + Environment.NewLine + "От кого: " + message.From + Environment.NewLine +
                    "Сообщение:" + message.Body;
                files.Items.Clear();
                foreach (var ado in message.Attaches)
                {
                    classforlistbox newitem = new classforlistbox(ado);
                    files.Items.Add(newitem);
                }
            }
        }

        private void Blogout_Click(object sender, EventArgs e)
        {
            groupBox3.Hide();
            groupBox1.Show();
            messages.Items.Clear();
            richTextBox1.Clear();
            treceiver.SelectedIndex = -1;
            tsubject.Clear();
            ListFiles.Items.Clear();
            messageBox.Clear();
            files.Items.Clear();
            dataGridView1.Rows.Clear();
            tabs.Enabled = false;
            myThread.Abort();
        }

        private void ShowPage()
        {
            using (Pop3Client client = new Pop3Client())
            {
                if (String.IsNullOrEmpty(serverpop.Trim()))
                {
                    // Подключение к серверу
                    string servname = (tsender.Text.Split(new char[] { '@' }))[1];
                    try { client.Connect("pop." + servname, portpop, true); } catch (Exception exept) { MessageBox.Show("Проверьте порт и сервер POP3. " + exept.Message, "Неверные данные", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                }
                else try { client.Connect(serverpop, portpop, true); } catch (Exception exept) { MessageBox.Show("Проверьте порт и сервер POP3. " + exept.Message, "Неверные данные", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                client.Authenticate(tsender.Text, tpassword.Text, AuthenticationMethod.UsernameAndPassword);
                if (client.Connected)
                {
                    int messageCount = client.GetMessageCount();
                    ShowMessages(messageCount - page * 100, (messageCount - page * 100 - 100) < 0 ? 0 : messageCount - page * 100 - 100);
                    if (numericUpDown1.Maximum <= page) bNext.Hide(); else bNext.Show();
                    if (page == 0) bPrev.Hide(); else bPrev.Show();
                }
            }
        }
        private void BNext_Click(object sender, EventArgs e)
        {
            page++;
            ShowPage();
        }

        private void BPrev_Click(object sender, EventArgs e)
        {
            page--;
            ShowPage();
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {

            if (checkBox1.Checked) { tpassword.PasswordChar = '\0'; } else { tpassword.PasswordChar = '*'; }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            page = (int)numericUpDown1.Value;
            ShowPage();
        }

        private void Files_DoubleClick(object sender, EventArgs e)
        {
            if (files.SelectedIndex != -1)
            {
               if(!Directory.Exists(Environment.CurrentDirectory+"\\Files")) Directory.CreateDirectory("Files");
                System.Diagnostics.Process.Start("explorer", "Files");
                MessagePart ado = ((classforlistbox)files.SelectedItem).Attach;
                ado.Save(new System.IO.FileInfo(System.IO.Path.Combine("Files", ado.FileName)));
            }
        }

        private void Bdelete_Click(object sender, EventArgs e)
        {
            List<int> arrayofindex = new List<int>();
            bool nofordelete = true;
            for (int i = 0; i < messages.Items.Count; i++)
            {
                if ((bool)dataGridView1.Rows[i].Cells[1].Value == true)
                {
                    nofordelete = false;
                    arrayofindex.Add((Int16.Parse(((OneMessage)messages.Items[i]).Index)));
                }
            }
            if (nofordelete) MessageBox.Show("Вы не выбрали ни одного сообщения для удаления.", "Не выбраны сообщения", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
            {
                string text = "Вы уверены, что хотите удалить сообщения с номерами: ";
                foreach (var index in arrayofindex) { text = text + index + ", "; }
                text = text + "Удаление невозможно отменить!";
                if (DialogResult.Yes == MessageBox.Show(text, "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                {
                    using (Pop3Client client = new Pop3Client())
                    {
                        if (String.IsNullOrEmpty(serverpop.Trim()))
                        {
                            // Подключение к серверу
                            string servname = (tsender.Text.Split(new char[] { '@' }))[1];
                            try { client.Connect("pop." + servname, portpop, true); } catch (Exception exept) { MessageBox.Show("Проверьте порт и сервер POP3. " + exept.Message, "Неверные данные", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                        }
                        else try { client.Connect(serverpop, portpop, true); } catch (Exception exept) { MessageBox.Show("Проверьте порт и сервер POP3. " + exept.Message, "Неверные данные", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                        client.Authenticate(tsender.Text, tpassword.Text, AuthenticationMethod.UsernameAndPassword);
                        if (client.Connected)
                        {
                            foreach (var index in arrayofindex)
                            {
                                client.DeleteMessage(index);
                            }
                        }
                        client.Disconnect();
                        MessageBox.Show("Сообщения удалены.", "Удаление", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        ShowPage();
                    }
                }

            }
        }

        private void DataGridView1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                int index = dataGridView1.CurrentCell.RowIndex;
                if (index >= 0 && index < messages.Items.Count)
                {
                    OneMessage message = (OneMessage)messages.Items[index];
                    richTextBox1.Text = "Тема: " + message.Subject + Environment.NewLine +
                        "Дата: " + message.Date + Environment.NewLine + "От кого: " + message.From + Environment.NewLine +
                        "Сообщение:" + message.Body;
                    files.Items.Clear();
                    foreach (var ado in message.Attaches)
                    {
                        classforlistbox newitem = new classforlistbox(ado);
                        files.Items.Add(newitem);
                    }
                }
            }
            catch { }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (myThread != null)
                myThread.Abort();

        }
        private bool isValid(string email)
        {
            string pattern = "[.\\-_a-z0-9]+@([a-z0-9][\\-a-z0-9]+\\.)+[a-z]{2,6}";
            Match isMatch = Regex.Match(email, pattern, RegexOptions.IgnoreCase);
            return isMatch.Success;
        }
        private void Baddemail_Click(object sender, EventArgs e)
        {
            if (isValid(emailformailing.Text))
            {
                if (!listemails.Items.Contains(emailformailing.Text))
                    listemails.Items.Add(emailformailing.Text);
                emailformailing.SelectedItem = -1;
            }
            else { MessageBox.Show("Введённый текст не является Email адресом.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        private void Bcreatetemplate_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(subjecttempl.Text) && !string.IsNullOrEmpty(bodytempl.Text) && !string.IsNullOrEmpty(NameTemplate.Text))
            {
                var files = filestempl.Items.Cast<object>().ToArray();
                Template newtempl = new Template(NameTemplate.Text, subjecttempl.Text, bodytempl.Text, files);
                listtemplate.Add(newtempl);
                listtempletes.Items.Add(newtempl);
                subjecttempl.Clear();bodytempl.Clear();filestempl.Items.Clear();NameTemplate.Clear();
            }
            else {
                string text = "Заполните поля: " + (string.IsNullOrEmpty(subjecttempl.Text) ? "Тему, " : "") +
                    (string.IsNullOrEmpty(bodytempl.Text) ? "Сообщение, " : "") + (string.IsNullOrEmpty(NameTemplate.Text) ? "Название шаблона." : "");
                MessageBox.Show(text,"Поля не заполнены",MessageBoxButtons.OK,MessageBoxIcon.Warning);
            }

        }

        private void Baddattachtempl_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (!filestempl.Items.Contains(openFileDialog1.FileName))
                    filestempl.Items.Add(openFileDialog1.FileName);
            }
        }

        private void Tabs_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = tabs.SelectedIndex;
            switch (index)
            {
                case 1:
                    templates.Items.Clear();
                    treceiver.Items.Clear();
                    foreach (var email in EmailBook.Items) { treceiver.Items.Add(email); }
                    foreach (var template in listtemplate)
                    {
                        templates.Items.Add(template);
                    }
                    break;
                case 2:
                    liststartedmailing.Items.Clear();
                    emailformailing.Items.Clear();
                    foreach (var email in EmailBook.Items) { emailformailing.Items.Add(email); }
                    foreach (var mailing in listmailng)
                    {
                        liststartedmailing.Items.Add(mailing);
                    }
                    mailingtempletes.Items.Clear();
                    foreach(var template in listtemplate)
                    {
                        mailingtempletes.Items.Add(template);
                    }
                    break;
                case 3:
                    listtempletes.Items.Clear();
                    foreach (var template in listtemplate)
                    {
                        listtempletes.Items.Add(template);
                    }
                    break;

            }
        }

        private void Bstartmailing_Click(object sender, EventArgs e)
        {
            var stringsemails = listemails.Items.Cast<String>().ToList();
            if (stringsemails.Count != 0 && !String.IsNullOrEmpty(NameMailing.Text) && mailingtempletes.SelectedItem != null)
            {
                MailingObject newmailing = new MailingObject(NameMailing.Text, stringsemails, (Template)mailingtempletes.SelectedItem, (int)hours.Value, (int)minutes.Value);
                liststartedmailing.Items.Add(newmailing);
                listmailng.Add(newmailing);
                listemails.Items.Clear();NameMailing.Clear();hours.Value = hours.Minimum;minutes.Value = minutes.Minimum;
            }
            else
            {
                string text = (stringsemails.Count == 0 ? "Добавьте электронные почты для рассылки. " : "") +
                    (mailingtempletes.SelectedItem == null ? "Выберите шаблон. " : "") + (String.IsNullOrEmpty(NameMailing.Text) ? "Введите имя рассылки." : "");
                MessageBox.Show(text, "Заполните все поля", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Templates_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (templates.SelectedItem != null)
            {
                Template templ =(Template) templates.SelectedItem;
                tsubject.Text = templ.Subject;
                messageBox.Text = templ.Body;
                ListFiles.Items.Clear();
                ListFiles.Items.AddRange(templ.Attaches);
            }
        }

        private void Bdeletemailing_Click(object sender, EventArgs e)
        {
            if (liststartedmailing.SelectedIndex != -1)
            {
                listmailng.RemoveAt(liststartedmailing.SelectedIndex);
                liststartedmailing.Items.RemoveAt(liststartedmailing.SelectedIndex);
            }
        }

        private void Bdeletetempl_Click(object sender, EventArgs e)
        {
            if (listtempletes.SelectedIndex != -1)
            {
                listtemplate.RemoveAt(listtempletes.SelectedIndex);
                listtempletes.Items.RemoveAt(listtempletes.SelectedIndex);
            }
        }

        private void BSavetempl_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(Environment.CurrentDirectory + "\\Saves")) Directory.CreateDirectory("Saves"); ;
            saveFileDialog1.DefaultExt = ".txml";
            if (DialogResult.OK == saveFileDialog1.ShowDialog())
            {
                try
                {
                    XmlSerializer formatter = new XmlSerializer(typeof(Template[]));
                    using (FileStream fs = new FileStream(saveFileDialog1.FileName, FileMode.Create))
                    {
                        formatter.Serialize(fs, listtemplate.ToArray());
                    }
                    MessageBox.Show("Сохранение прошло успешно.", "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch { MessageBox.Show("Произошла ошибка при сохранении.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            }
        }

        private void BLoadtempl_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(Environment.CurrentDirectory + "\\Saves")) Directory.CreateDirectory("Saves");
            if (DialogResult.OK == openFileDialog2.ShowDialog())
            {
                if (openFileDialog2.FileName.Contains(".txml"))
                {
                    try
                    {
                        XmlSerializer formatter = new XmlSerializer(typeof(Template[]));
                        listtemplate.Clear();
                        listtempletes.Items.Clear();
                        using (FileStream fs = new FileStream(openFileDialog2.FileName, FileMode.OpenOrCreate))
                        {
                            Template[] templates = (Template[])formatter.Deserialize(fs);
                            foreach (var template in templates)
                            {
                                listtemplate.Add(template);
                                listtempletes.Items.Add(template);
                            }

                        }
                        MessageBox.Show("Загрузка прошла успешна.", "Загрузка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch { MessageBox.Show("Произошла ошибка при загрузке. Возможно файл был повреждён.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                }
                else { MessageBox.Show("Попытка загрузки файла другого формата.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            }
        }

        private void Listtempletes_DoubleClick(object sender, EventArgs e)
        {
            if (listtempletes.SelectedItem!=null) {
                Template templ = (Template) listtempletes.SelectedItem;
                string text = "Название шаблона: " + (templ.Name != null ? templ.Name : "Пусто...") + System.Environment.NewLine +
                    "Тема: " + (templ.Subject != null ? templ.Subject : "Пусто...") + System.Environment.NewLine +
                    "Сообщение: " + (templ.Body != null ? templ.Body : "Пусто...") + System.Environment.NewLine+"Прикреплённые файлы: ";
                if (templ.Attaches.Length == 0) text = text + "Нет.";
                else
                {
                    string files="";
                    foreach (var attach in templ.Attaches) { files += attach+", "; }
                    text+=files;
                }
                MessageBox.Show(text,"Информация о шаблоне", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Liststartedmailing_DoubleClick(object sender, EventArgs e)
        {
            if (liststartedmailing.SelectedItem != null)
            {
                MailingObject mailing = (MailingObject)liststartedmailing.SelectedItem;
                string text = "Название рассылки: " + (mailing.Name != null ? mailing.Name : "Пусто...") + System.Environment.NewLine +
                    "Время рассылки: " + (mailing.Hours<10? "0"+mailing.Hours.ToString() : mailing.Hours.ToString()) + ":" + (mailing.Minutes < 10 ? "0" + mailing.Minutes.ToString() : mailing.Minutes.ToString())
                    + System.Environment.NewLine + "Выбранный шаблон: " + System.Environment.NewLine;
                if (mailing.Template != null)
                {
                    Template templ = mailing.Template;
                    text += "   Название шаблона: " + (templ.Name != null ? templ.Name : "Пусто...") + System.Environment.NewLine +
                    "   Тема: " + (templ.Subject != null ? templ.Subject : "Пусто...") + System.Environment.NewLine +
                    "   Сообщение: " + (templ.Body != null ? templ.Body : "Пусто...") + System.Environment.NewLine + "   Прикреплённые файлы: ";
                    if (templ.Attaches.Length == 0) text = text + "Нет.";
                    else
                    {
                        string files = "";
                        foreach (var attach in templ.Attaches) { files += attach + ", "; }
                        text += files + System.Environment.NewLine;
                    }
                }
                else text += System.Environment.NewLine;
                string emails = "";
                foreach (var email in mailing.ListEmails) { emails += email + ", "; }
                text +="Для кого рассылка: "+ emails;
                MessageBox.Show(text, "Информация о рассылке", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void BSaveMailing_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(Environment.CurrentDirectory + "\\Saves")) Directory.CreateDirectory("Saves"); ;
            saveFileDialog1.DefaultExt = ".mxml";
            if (DialogResult.OK == saveFileDialog1.ShowDialog())
            {
                try
                {
                    XmlSerializer formatter = new XmlSerializer(typeof(MailingObject[]));
                    using (FileStream fs = new FileStream(saveFileDialog1.FileName, FileMode.Create))
                    {
                        formatter.Serialize(fs, listmailng.ToArray());
                    }
                    MessageBox.Show("Сохранение прошло успешно.", "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch { MessageBox.Show("Произошла ошибка при сохранении.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            }
        }

        private void BLoadMailing_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(Environment.CurrentDirectory + "\\Saves")) Directory.CreateDirectory("Saves");
            if (DialogResult.OK == openFileDialog2.ShowDialog())
            {
                if (openFileDialog2.FileName.Contains(".mxml"))
                {
                    try
                    {
                        XmlSerializer formatter = new XmlSerializer(typeof(MailingObject[]));
                        listmailng.Clear();
                        liststartedmailing.Items.Clear();
                        using (FileStream fs = new FileStream(openFileDialog2.FileName, FileMode.OpenOrCreate))
                        {
                            MailingObject[] mailings = (MailingObject[])formatter.Deserialize(fs);
                            foreach (var mailing in mailings)
                            {
                                listmailng.Add(mailing);
                                liststartedmailing.Items.Add(mailing);
                            }

                        }
                        MessageBox.Show("Загрузка прошла успешна.", "Загрузка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch { MessageBox.Show("Произошла ошибка при загрузке. Возможно файл был повреждён.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                }
                else{ MessageBox.Show("Попытка загрузки файла другого формата.","Ошибка",MessageBoxButtons.OK,MessageBoxIcon.Warning); }
            }
        }

        private void Baddtobookemail_Click(object sender, EventArgs e)
        {
            if (isValid(emailtobook.Text))
            {
                if (!EmailBook.Items.Contains(emailtobook.Text))
                    EmailBook.Items.Add(emailtobook.Text);
                emailtobook.Clear();
            }
            else { MessageBox.Show("Введённый текст не является Email адресом.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        private void Filestempl_DoubleClick(object sender, EventArgs e)
        {
            if (filestempl.SelectedIndex != -1)
                filestempl.Items.RemoveAt(filestempl.SelectedIndex);
        }

        private void Bdeleteemail_Click(object sender, EventArgs e)
        {
            if (EmailBook.SelectedIndex != -1)
            {
                EmailBook.Items.RemoveAt(EmailBook.SelectedIndex);
            }
        }

        private void Bloadbook_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(Environment.CurrentDirectory + "\\Saves")) Directory.CreateDirectory("Saves");
            if (DialogResult.OK == openFileDialog2.ShowDialog())
            {
                if (openFileDialog2.FileName.Contains(".bxml"))
                {
                    try
                    {
                        XmlSerializer formatter = new XmlSerializer(typeof(object[]));
                        EmailBook.Items.Clear();
                        using (FileStream fs = new FileStream(openFileDialog2.FileName, FileMode.OpenOrCreate))
                        {
                            object[] emailsbook = (object[])formatter.Deserialize(fs);
                            EmailBook.Items.AddRange(emailsbook);
                        }
                        MessageBox.Show("Загрузка прошла успешна.", "Загрузка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch { MessageBox.Show("Произошла ошибка при загрузке. Возможно файл был повреждён.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                }
                else { MessageBox.Show("Попытка загрузки файла другого формата.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            }
        }

        private void Bsavebook_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(Environment.CurrentDirectory + "\\Saves")) Directory.CreateDirectory("Saves"); ;
            saveFileDialog1.DefaultExt = ".bxml";
            if (DialogResult.OK == saveFileDialog1.ShowDialog())
            {
                try
                {
                    XmlSerializer formatter = new XmlSerializer(typeof(object[]));
                    using (FileStream fs = new FileStream(saveFileDialog1.FileName, FileMode.Create))
                    {
                        formatter.Serialize(fs, EmailBook.Items);
                    }
                    MessageBox.Show("Сохранение прошло успешно.", "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch { MessageBox.Show("Произошла ошибка при сохранении.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            }
        }

        private void EmailBook_DoubleClick(object sender, EventArgs e)
        {
            if (EmailBook.SelectedIndex != -1)
                EmailBook.Items.RemoveAt(EmailBook.SelectedIndex);
        }

        private void B_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(pop3port.Text.Trim())) int.TryParse(pop3port.Text, out portpop); else portpop=995;
            if (!String.IsNullOrEmpty(smtpport.Text.Trim())) int.TryParse(smtpport.Text, out portsmtp); else portsmtp=25;
            if (!String.IsNullOrEmpty(pop3server.Text.Trim())) serverpop = pop3server.Text.Trim(); else serverpop="";
            if (!String.IsNullOrEmpty(smtpserver.Text.Trim())) serversmtp = smtpserver.Text.Trim(); else serversmtp="";
                     
        }

        private void Bsettings_Click(object sender, EventArgs e)
        {
            if (gsettings.Visible) gsettings.Hide(); else gsettings.Show();
        }
    }
    class classforlistbox
    {
        public MessagePart Attach { get; set; }
        public classforlistbox(MessagePart attach)
        {
            Attach = attach;
        }
        public override string ToString() { return Attach.FileName; }
    }
    class OneMessage
    {
        public string Subject { get; set; }
        public string Date { get; set; }
        public string From { get; set; }
        public string Body { get; set; }
        public string Index { get; set; }
        public List<MessagePart> Attaches { get; set; }

        public OneMessage(string[] messatr, List<MessagePart> attaches)
        {
            Subject = messatr[0];
            Date = messatr[1];
            From = messatr[2];
            Body = messatr[3];
            Index = messatr[4];
            Attaches = attaches;
        }
        public OneMessage(Message message)
        {
            Subject = message.Headers.Subject;
            Date = message.Headers.Date.ToString();
            From = message.Headers.From.ToString();

            MessagePart mpPlain = message.FindFirstPlainTextVersion();

            if (mpPlain != null)
            {
                Encoding enc = mpPlain.BodyEncoding;
                Body = enc.GetString(mpPlain.Body); //получаем текст сообщения
            }


        }
        public override string ToString() { return Index + ") " + Subject; }
    }
    public class MailingObject
    {
        public string Name { get; set; }
        public List<string> ListEmails{ get; set; }
        public Template Template { get; set; }
        public int Hours { get; set; }
        public int Minutes { get; set; }
        public MailingObject(string name,List<string>listemails,Template template,int hours,int minutes)
        {
            Name = name; ListEmails = listemails; Template = template; Hours = hours; Minutes = minutes;
        }
        public MailingObject() { }
        public override string ToString() { return Name; }
    }
    public class Template
    {
        public string Name { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public object[] Attaches { get; set; }

        public Template(string name, string subject, string body, object[] attaches)
        {
            Name = name; Subject = subject; Body = body; Attaches = attaches;
        }
        public Template() { }
        public override string ToString() { return Name; }
    }
}
