using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;



namespace em_2._0
{
    
    public partial class Form1 : Form
    {
        class ConfigData
        {
            public string path { get; set; }
            public string extension { get; set; }

            public string output_folder { get; set; }
        }
        public Form1()
        {
            InitializeComponent();
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        static List<string> SaveToPST(string[] filepaths, string outputPath)
        {
            List<string> processedFiles = new List<string>();

            Outlook.Application outlookApp = new Outlook.Application();
            Outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
            Outlook.Folder rootFolder = outlookNamespace.Session.DefaultStore.GetRootFolder() as Outlook.Folder;


            string filePath = @"C:\em\.config.json";
            string directory = "";
            if (File.Exists(filePath))
            {  // Olvassa be a JSON tartalmat a fájlból
                string jsonContent = File.ReadAllText(filePath);

                // Deszerializálja a JSON-t egy objektummá
                ConfigData configData = JsonConvert.DeserializeObject<ConfigData>(jsonContent);

                // Most már használhatod a configData objektumot, amely tartalmazza a JSON-ből kiolvasott adatokat

                directory = configData.output_folder;
            }
            else
            {
                MessageBox.Show("Ismeretlen Hiba");
            }

                // Ellenőrizzük, hogy a célmappa létezik-e, ha nem, létrehozzuk
                Outlook.Folder targetFolder;
            try
            {
                targetFolder = rootFolder.Folders[directory] as Outlook.Folder;
            }
            catch (System.Exception)
            {
                targetFolder = rootFolder.Folders.Add(directory, Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
            }

            // Fájlok feldolgozása és küldése
            for (int i = 0; i < filepaths.Length; i++)
            {
                // Az adatok hozzáadása a processedFiles listához
                processedFiles.Add(filepaths[i]);

                Outlook.MailItem mailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                mailItem.Subject = Path.GetFileName(filepaths[i]);
                mailItem.Body = "This is the attached file: " + Path.GetFileName(filepaths[i]);
                mailItem.Attachments.Add(filepaths[i]);
                mailItem.Save();
                mailItem.Move(targetFolder);
            }

            outlookNamespace.Logoff();

            // Visszaadjuk a feldolgozott fájlneveket
            return processedFiles;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = @"C:\em\.config.json";

            if (File.Exists(filePath))
            {
                // Olvassa be a JSON tartalmat a fájlból
                string jsonContent = File.ReadAllText(filePath);

                // Deszerializálja a JSON-t egy objektummá
                ConfigData configData = JsonConvert.DeserializeObject<ConfigData>(jsonContent);

                // Most már használhatod a configData objektumot, amely tartalmazza a JSON-ből kiolvasott adatokat
                MessageBox.Show($"A config file tartalma:\npath: {configData.path}\nKiterjesztés: {configData.extension}\nMentési mappa: {configData.output_folder}");
            }
            else
            {
                MessageBox.Show("A fájl nem létezik: " + filePath);
            }
            
        }

        private void Start_Click(object sender, EventArgs e)
        {
            string filePath = @"C:\em\.config.json";
            if (File.Exists(filePath))
            {
                // Olvassa be a JSON tartalmat a fájlból
                string jsonContent = File.ReadAllText(filePath);

                // Deszerializálja a JSON-t egy objektummá
                ConfigData configData = JsonConvert.DeserializeObject<ConfigData>(jsonContent);

                // Most már használhatod a configData objektumot, amely tartalmazza a JSON-ből kiolvasott adatokat

                string directoryPath = configData.path;
                string extension = configData.extension;
                

                string[] filepaths = Directory.GetFiles(directoryPath, "*" + extension, SearchOption.AllDirectories);

                // A feldolgozott fájlnevek megjelenítése az textBox1 szövegében
                foreach (string processedFile in SaveToPST(filepaths, "C:\\Output.pst"))
                {
                    textBox1.AppendText(processedFile + Environment.NewLine);
                }
            }
            else
            {
                MessageBox.Show("A fájl nem létezik: " + filePath);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();

            folderBrowserDialog1.Description = "Select a folder";
            folderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer; // Indítási mappa beállítása

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = @"C:\em\.config.json";

                if (File.Exists(filePath))
                {
                    // Olvassa be a JSON tartalmat a fájlból
                    string jsonContent = File.ReadAllText(filePath);

                    // Deszerializálja a JSON-t egy objektummá
                    ConfigData configData = JsonConvert.DeserializeObject<ConfigData>(jsonContent);

                    // Path módosítása

                    // Ha a felhasználó kiválasztott egy mappát, itt kezelheted le, pl. megteheted vele, amit akarsz
                    string newPath = folderBrowserDialog1.SelectedPath;
                    configData.path = newPath;

                    // A fájlba írás előtt JSON formátumra alakítjuk a configData objektumot
                    string updatedJsonContent = JsonConvert.SerializeObject(configData, Formatting.Indented);

                    // A fájlba írás
                    File.WriteAllText(filePath, updatedJsonContent);

                    MessageBox.Show($"Mappa sikeresen módosítva. Erre {newPath}");
                }
                else
                {
                    MessageBox.Show("A fájl nem létezik: " + filePath);
                }
                
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Létrehozunk egy szöveg beviteli mezőt
            TextBox inputTextBox = new TextBox();
            inputTextBox.Dock = DockStyle.Top;

            // Létrehozunk egy párbeszédpanelt, amely tartalmazza a szöveg beviteli mezőt
            Form inputForm = new Form();
            inputForm.Text = "kiterjesztés pl:.mp4";
            inputForm.Size = new System.Drawing.Size(180, 100);
            inputForm.Controls.Add(inputTextBox);

            // Létrehozunk egy gombot a párbeszédpanelen, hogy megerősíthessük a beírt szöveget
            Button confirmButton = new Button();
            confirmButton.Text = "OK";
            confirmButton.Dock = DockStyle.Bottom;
            confirmButton.Click += (confirmSender, confirmEventArgs) =>
            {
                string filePath = @"C:\em\.config.json";

                if (File.Exists(filePath))
                {
                    // Olvassa be a JSON tartalmat a fájlból
                    string jsonContent = File.ReadAllText(filePath);

                    // Deszerializálja a JSON-t egy objektummá
                    ConfigData configData = JsonConvert.DeserializeObject<ConfigData>(jsonContent);
                    string newExtension = inputTextBox.Text;
                    configData.extension = newExtension;

                    // A fájlba írás előtt JSON formátumra alakítjuk a configData objektumot
                    string updatedJsonContent = JsonConvert.SerializeObject(configData, Formatting.Indented);

                    // A fájlba írás
                    File.WriteAllText(filePath, updatedJsonContent);

                    MessageBox.Show("Kiterjesztés módosítva.");
                }
                else
                {
                    Console.WriteLine("A fájl nem létezik: " + filePath);
                }
                inputForm.Close(); // Bezárjuk a párbeszédpanelt
            };
            inputForm.Controls.Add(confirmButton);

            // Megjelenítjük a párbeszédpanelt modális módban, hogy ne lehessen más tevékenységet végezni
            inputForm.ShowDialog();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            // Létrehozunk egy szöveg beviteli mezőt
            TextBox inputTextBox = new TextBox();
            inputTextBox.Dock = DockStyle.Top;

            // Létrehozunk egy párbeszédpanelt, amely tartalmazza a szöveg beviteli mezőt
            Form inputForm = new Form();
            inputForm.Text = "kiterjesztés pl:.mp4";
            inputForm.Size = new System.Drawing.Size(180, 100);
            inputForm.Controls.Add(inputTextBox);

            // Létrehozunk egy gombot a párbeszédpanelen, hogy megerősíthessük a beírt szöveget
            Button confirmButton = new Button();
            confirmButton.Text = "OK";
            confirmButton.Dock = DockStyle.Bottom;
            confirmButton.Click += (confirmSender, confirmEventArgs) =>
            {
                string filePath = @"C:\em\.config.json";

                if (File.Exists(filePath))
                {
                    // Olvassa be a JSON tartalmat a fájlból
                    string jsonContent = File.ReadAllText(filePath);

                    // Deszerializálja a JSON-t egy objektummá
                    ConfigData configData = JsonConvert.DeserializeObject<ConfigData>(jsonContent);
                    string newExtension = inputTextBox.Text;
                    configData.output_folder = newExtension;

                    // A fájlba írás előtt JSON formátumra alakítjuk a configData objektumot
                    string updatedJsonContent = JsonConvert.SerializeObject(configData, Formatting.Indented);

                    // A fájlba írás
                    File.WriteAllText(filePath, updatedJsonContent);

                    MessageBox.Show("Mentési Mappa módosítva.");
                }
                else
                {
                    Console.WriteLine("A fájl nem létezik: " + filePath);
                }
                inputForm.Close(); // Bezárjuk a párbeszédpanelt
            };
            inputForm.Controls.Add(confirmButton);

            // Megjelenítjük a párbeszédpanelt modális módban, hogy ne lehessen más tevékenységet végezni
            inputForm.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string directoryPath = @"C:\em";
            string filePath = @"C:\em\.config.json";

            // Ellenőrizze, hogy létezik-e a könyvtár
            if (!Directory.Exists(directoryPath))
            {
                // Ha nem létezik, hozza létre
                Directory.CreateDirectory(directoryPath);
                MessageBox.Show("Status: Config mappa inicializálás.");
            }

            // Ellenőrizze, hogy létezik-e a fájl
            if (!File.Exists(filePath))
            {
                // Ha nem létezik, hozza létre
                using (FileStream fs = File.Create(filePath))
                    MessageBox.Show("Status: Config file configurálás.");
            }
            else
            {
                MessageBox.Show("Status: Üdv újra minden rendben van.\n\n");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
