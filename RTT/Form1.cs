using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace RTT
{
    public partial class Form1 : Form
    {
        List<String> filesNames = new List<string>();

        public int OpenWordFileAndReplace(String fileName, String name)
        {
			String currFileName = fileName.Replace(".docx", "_" + name + ".docx");
			Word.Application app = new Word.Application();
            try
            {
                File.Copy(fileName, currFileName);

                Object fileNameObj = @currFileName;
                Object missing = Type.Missing;
                app.Documents.Open(ref fileNameObj);
                Word.Find find = app.Selection.Find;

				FindAndReplace(find, missing, "!POLE1!", textBox1.Text);
				FindAndReplace(find, missing, "!POLE2!", textBox2.Text);
				FindAndReplace(find, missing, "!POLE3!", textBox3.Text);
				FindAndReplace(find, missing, "!POLE4!", textBox4.Text);
				FindAndReplace(find, missing, "!POLE5!", textBox5.Text);
				FindAndReplace(find, missing, "!POLE6!", textBox6.Text);
				FindAndReplace(find, missing, "!POLE7!", textBox7.Text);
				FindAndReplace(find, missing, "!POLE8!", textBox8.Text);
				FindAndReplace(find, missing, "!POLE9!", textBox9.Text);
				FindAndReplace(find, missing, "!POLE10!", textBox10.Text);
				FindAndReplace(find, missing, "!POLE11!", textBox11.Text);
				FindAndReplace(find, missing, "!POLE12!", textBox12.Text);
				FindAndReplace(find, missing, "!POLE13!", textBox13.Text);
				FindAndReplace(find, missing, "!POLE14!", textBox14.Text);
				FindAndReplace(find, missing, "!POLE15!", textBox15.Text);
				FindAndReplace(find, missing, "!POLE16!", textBox16.Text);
				FindAndReplace(find, missing, "!POLE17!", textBox17.Text);
				FindAndReplace(find, missing, "!POLE18!", textBox18.Text);
				FindAndReplace(find, missing, "!POLE19!", textBox19.Text);
				FindAndReplace(find, missing, "!POLE20!", textBox20.Text);
				FindAndReplace(find, missing, "!POLE21!", textBox21.Text);
				FindAndReplace(find, missing, "!POLE22!", textBox22.Text);
				FindAndReplace(find, missing, "!POLE23!", textBox21.Text);
				FindAndReplace(find, missing, "!POLE24!", textBox24.Text);
				FindAndReplace(find, missing, "!POLE25!", textBox25.Text);
				FindAndReplace(find, missing, "!POLE26!", textBox26.Text);
				FindAndReplace(find, missing, "!POLE27!", textBox27.Text);
				FindAndReplace(find, missing, "!POLE28!", textBox28.Text);
				FindAndReplace(find, missing, "!POLE29!", textBox29.Text);
				FindAndReplace(find, missing, "!POLE30!", textBox30.Text);

                app.ActiveDocument.Save();
                app.ActiveDocument.Close();
                app.Quit();
				return 0;
            }
            catch (Exception exep)
            {
				app.Quit();
                MessageBox.Show("Ошибка. " + exep.Message);
				try
				{
					File.Delete(currFileName);
				}
				catch (Exception e) { }
				return -1;
            }

        }

        public void FindAndReplace(Word.Find find, Object missing, String findText, String replacebleText)
        {
			String[] partsOfRepText = replacebleText.Split(' ');
			for (int i = 0; i < partsOfRepText.Length; i++) {
				if (i != partsOfRepText.Length - 1)
					partsOfRepText[i] += " " + findText;

				find.Text = findText;
				find.Replacement.Text = partsOfRepText[i];
				Object wrap = Word.WdFindWrap.wdFindContinue;
				Object replace = Word.WdReplace.wdReplaceAll;
				find.Execute(FindText: Type.Missing,
					MatchCase: false,
					MatchWholeWord: false,
					MatchWildcards: false,
					MatchSoundsLike: missing,
					MatchAllWordForms: false,
					Forward: true,
					Wrap: wrap,
					Format: false,
					ReplaceWith: missing, Replace: replace);

		}
        }

        public Form1()
        {
            InitializeComponent();
			this.MouseWheel += new MouseEventHandler(this_MouseWheel);
        }

		void this_MouseWheel(object sender, MouseEventArgs e)
		{
			if (e.Delta > 0)
			{
				if (panel1.VerticalScroll.Value - 50 < 0) {
					panel1.VerticalScroll.Value = 0;
				} 
				else
					panel1.VerticalScroll.Value -= 50;
			}
			else if (e.Delta != 0)
				panel1.VerticalScroll.Value += 50;
		}

        private void Form1_Load(object sender, EventArgs e)
        { }

        private void ExploreFile_Click(object sender, EventArgs e)
        {
            openFileDialog.ShowDialog();
            for (int i = 0; i < openFileDialog.SafeFileNames.Length; i++)
                filesNames.Add(Path.GetDirectoryName(openFileDialog.FileName) + "//" + openFileDialog.SafeFileNames[i]);

            fullFileName.Text = "Выбрано файлов: " + openFileDialog.SafeFileNames.Length;
        }

        private void Replace_Click(object sender, EventArgs e)
        {
			if (codeName.Text != "" && filesNames.Count > 0)
			{
				progressBar1.Value = 0;
				progressBar1.Maximum = filesNames.Count;
				progressBar1.Visible = true;

				String newFilesNames = "";

				foreach (String fileName in filesNames)
				{
					progressBar1.Value++;
					if (OpenWordFileAndReplace(fileName, codeName.Text) == -1)
					{
						progressBar1.Visible = false;
						MessageBox.Show("В ходе выполнения программы возникли ошибки.");
						return;
					}
					newFilesNames += (fileName.Replace(".docx", "_" + codeName.Text + ".docx ") + Environment.NewLine);
				}

				MessageBox.Show("Замены выполнены и сохранены в файлы/файл: " + newFilesNames + "Have a nice day.", "Замены успешно выполнены");
				progressBar1.Visible = false;
			}
			else
				MessageBox.Show("Введите постфикс и выберите необходимые файлы.");
        }
    }
}
