using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace PDFBinder
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            UpdateUI();
        }

        public void AddInputFile(string file)
        {
            switch (Combiner.TestSourceFile(file))
            {
                case Combiner.SourceTestResult.Unreadable:
                    MessageBox.Show(string.Format("文件不能以PDF格式打开:\n\n{0}", file), "Illegal file type", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                case Combiner.SourceTestResult.Protected:
                    MessageBox.Show(string.Format("PDF文档不能复制:\n\n{0}", file), "Permission denied", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    break;
                case Combiner.SourceTestResult.Ok:
                    inputListBox.Items.Add(file);
                    break;
            }
        }

        public void UpdateUI()
        {
            if (inputListBox.Items.Count < 2)
            {
                completeButton.Enabled = false;
                helpLabel.Text = "拖放PDF文件到上面框内或者使用工具栏的“添加文件”按钮";
            }
            else
            {
                completeButton.Enabled = true;
                helpLabel.Text = "选择好文件后执行“合并”操作";
            }

            if (inputListBox.SelectedIndex < 0)
            {
                removeButton.Enabled = moveUpButton.Enabled = moveDownButton.Enabled = false;
            }
            else
            {
                removeButton.Enabled = true;
                moveUpButton.Enabled = (inputListBox.SelectedIndex > 0);
                moveDownButton.Enabled = (inputListBox.SelectedIndex < inputListBox.Items.Count - 1);
            }
        }

        private void inputListBox_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop, false) ? DragDropEffects.All : DragDropEffects.None;
        }

        private void inputListBox_DragDrop(object sender, DragEventArgs e)
        {
            var fileNames = (string[]) e.Data.GetData(DataFormats.FileDrop);
            Array.Sort(fileNames);

            foreach (var file in fileNames)
            {
                AddInputFile(file);
            }

            UpdateUI();
        }

        private void combineButton_Click(object sender, EventArgs e)
        {
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                using (var combiner = new Combiner(saveFileDialog.FileName))
                {
                    progressBar.Visible = true;
                    this.Enabled = false;

                    for (int i = 0; i < inputListBox.Items.Count; i++)
                    {
                        combiner.AddFile((string)inputListBox.Items[i]);
                        progressBar.Value = (int)(((i + 1) / (double)inputListBox.Items.Count) * 100);
                    }


                    this.Enabled = true;
                    progressBar.Visible = false;
                }

                System.Diagnostics.Process.Start(saveFileDialog.FileName);
            }
        }

        private void addFileButton_Click(object sender, EventArgs e)
        {
            if (addFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (string file in addFileDialog.FileNames)
                {
                    AddInputFile(file);
                }

                UpdateUI();
            }
        }

        private void inputListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateUI();
        }

        private void removeButton_Click(object sender, EventArgs e)
        {
            inputListBox.Items.Remove(inputListBox.SelectedItem);
        }

        private void moveItemButton_Click(object sender, EventArgs e)
        {
            object dataItem = inputListBox.SelectedItem;
            int index = inputListBox.SelectedIndex;

            if (sender == moveUpButton)
                index--;
            if (sender == moveDownButton)
                index++;

            inputListBox.Items.Remove(dataItem);
            inputListBox.Items.Insert(index, dataItem);

            inputListBox.SelectedIndex = index;
        }
    }
}