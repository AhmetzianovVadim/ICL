using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.NetworkInformation;
using System.Net;
using IWshRuntimeLibrary;
using System.IO;

namespace TestApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            programSettings(); //Метод, отвечающий за настройки программы
            createShortcut(); //Метод, создающий ярлык
            PCinfo(); //Метод выдающий информацию о компьютере пользователя
        }
        private void programSettings() 
        {
            this.Text = Properties.Settings.Default.FormName; //Устанавливаем имя формы, из файла конфигурации
            if (Properties.Settings.Default.SaveLocation) //Проверяем, хочет ли пользователь сохранять позицию окна после закрытия
            {
                this.Location = Properties.Settings.Default.Location; //Устанавливает координаты окна при запуске программы
                this.FormClosing += FormClosingEventHandler; //Привязываемся к обработчику закрытия формы
                StartPosition = FormStartPosition.Manual; //Указываем, что хотим использовать свои координаты окна
            } 
            else
                StartPosition = FormStartPosition.CenterScreen; //Стандартное положение окна в центре экрана
            if (Properties.Settings.Default.SaveAddress) //Проверяем хочет ли пользователь сохранять адрес сайта
            {
                this.textBox1.Text = Properties.Settings.Default.DefaultAddressOfSite; //Устанавливаем адрес сайта при запуске программы
                this.FormClosing += FormClosingEventHandler; ////Привязываемся к обработчику закрытия формы
            }
            else 
            {
                this.textBox1.Text = Properties.Settings.Default.DefaultAddressOfSite; //Устанавливаем адрес по умолчанию при запуске программы
            }
            try { this.Icon = new Icon(AppDomain.CurrentDomain.BaseDirectory + Properties.Settings.Default.Icon); } //Устанавливаем иконку формы
            catch { MessageBox.Show("Проверьте настройки иконки в файле конфигурации\nПрограмма всё ещё способна выполнять свои функции"); }
        }
        public void createShortcut() 
        {
            WshShell shell = new WshShell(); //Создаём объект WshShell
            string shortcutPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Ахметзянов_В_А_4310.lnk"; //указываем путь к ярылку и его название
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath); //Создаем объект ярлыка
            shortcut.IconLocation = AppDomain.CurrentDomain.BaseDirectory + @"\" + Properties.Settings.Default.Icon; //устанавливаем иконку ярлыка
            shortcut.Description = "Ярлык Ахметзянов Вадим Альбертович 4310"; //Добавляем описание ярлыка в всплывающей подсказке
            shortcut.TargetPath = AppDomain.CurrentDomain.BaseDirectory + @"\TestApp2.exe"; //Указываем путь к программе
            shortcut.Save(); //создаём ярлык
        }

        public void PCinfo() 
        {
            string Host = Dns.GetHostName(); // Получем имя ПК
            label2.Text = Host; //Отображаем его на форме
            IPAddress[] ips; // Создаём массив для IP 
            ips = Dns.GetHostAddresses(label2.Text); // Получаем IP с помощью имени ПК  
            foreach (IPAddress ip in ips)//Перебираем все результаты, в данном случае, один
            {
                label3.Text = ip.ToString(); //Выводим IP на форму
            }
        }

        public void ping() 
        {
            try
            {
                Ping pingSender = new Ping(); //Создаём объёкт класса Ping
                PingReply reply = pingSender.Send(textBox1.Text); //Отправляем запрос на проверку связи
                if (reply.Status == IPStatus.Success)  //Узнаём результат
                {
                    label1.ForeColor = Color.Green; //Перекрашиваем текст label1 в зелёный
                    label1.Text = "PING OK"; //Пишем, что всё в порядке
                    richTextBox1.ForeColor = Color.Green; //Перекрашиваем текст richTextBox1  в зелёный
                    // ↓↓↓ Выводим текст, с результатами пинга, новая информация не будет стирать предыдущую
                    richTextBox1.AppendText("Start log\r\n" + "Адрес:" + reply.Address.ToString() + "\r\nВремя: " + reply.RoundtripTime + "мс" + "\r\nРазмер пакета: " + reply.Buffer.Length + " байта" + "\r\nEnd Log\r\n\n"); 
                }
                else
                {
                    //Соединения нет, красим текст в красный, выдаём результат
                    richTextBox1.ForeColor = Color.Red;
                    richTextBox1.Text = "Соединения нет\n"; //Новая информация сотрёт старую
                    label1.ForeColor = Color.Red;
                    label1.Text = "PING FALSE"; 
                }
            }
            catch 
            {
                //Обрабатываем исключения, выдаём инструкции
                MessageBox.Show("Ошибка!\n1. Проверьте правильность ввода данных в поле \"Адрес сайта\"\n2. Проверьте подключение к Интернету");
            }    
        }

        private void button1_Click(object sender, EventArgs e) //Обработчик нажатия на кнопку
        {
            ping(); //Вызываем функцию пинга
        }

        private void FormClosingEventHandler(object sender, FormClosingEventArgs e) //Обработчик события закрытия формы
        {
            Properties.Settings.Default.Location = this.Location; //Присваиваем последние координаты
            Properties.Settings.Default.DefaultAddressOfSite = this.textBox1.Text; //Присваиваем последний введёный текст 
            Properties.Settings.Default.Save(); //Сохраняем изменения
        }
    }
}

