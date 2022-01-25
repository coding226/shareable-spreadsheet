using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Simulator;

namespace SpreadsheetApp
{
    public partial class Form1 : Form
    {

        private Panel buttonPanel = new Panel();
        private DataGridView dataGridView = new DataGridView();
        private Button loadButton = new Button();
        private Button saveButton = new Button();
        private Button searchButton = new Button();
        private Button swapButton = new Button();
        private int selectedRow = 0;

        SharableSpreadsheet sheet = new SharableSpreadsheet();
        private int rowIndexFromMouseDown;
        private int rowIndexOfItemUnderMouseToDrop;
        private Rectangle dragBoxFromMouseDown;
        public Form1()
        {
            this.Load += new EventHandler(Form1_Load);
        }

        private void Form1_Load(System.Object sender, System.EventArgs e)
        {
            SetupLayout();
        }

        private void SetupLayout()
        {
            this.Size = new Size(600, 500);

            loadButton.Text = "Load";
            loadButton.Location = new Point(10, 10);
            loadButton.Click += new EventHandler(loadButton_Click);

            saveButton.Text = "Save";
            saveButton.Location = new Point(100, 10);
            saveButton.Click += new EventHandler(saveButton_Click);

            searchButton.Text = "Search";
            searchButton.Location = new Point(200, 10);
            searchButton.Click += new EventHandler(searchButton_Click);

            swapButton.Text = "Swap Rows";
            swapButton.Location = new Point(300, 10);
            swapButton.Click += new EventHandler(swapRows_Click);

            buttonPanel.Controls.Add(loadButton);
            buttonPanel.Controls.Add(saveButton);
            buttonPanel.Controls.Add(searchButton);
            buttonPanel.Controls.Add(swapButton);
            buttonPanel.Height = 50;
            buttonPanel.Dock = DockStyle.Bottom;

            this.Controls.Add(this.buttonPanel);
        }

        private void loadButton_Click(object sender, EventArgs e)
        {
            sheet.load("D:\\TamilVanan\\Temp\\" + "User [65] .csv");
            String[] rows = sheet.loadData;
            String[] firstRowCols = rows[0].Split(",");

            SetupDataGridView(rows.Length, firstRowCols.Length);

            foreach (String row in rows)
            {
                String[] rowData = row.Split(",");
                int index = dataGridView.Rows.Add(rowData);
                dataGridView.Rows[index].HeaderCell.Value = "" + (index + 1);
            }
        }

        private void addNewRowButton_Click(object sender, EventArgs e)
        {
            this.dataGridView.Rows.Add();
        }
        private void saveButton_Click(object sender, EventArgs e)
        {
            sheet.save("D:\\TamilVanan\\Temp\\" + "Out.csv");
        }

        private void swapRows_Click(object sender, EventArgs e)
        {
            String row1 = "";
            String row2 = "";
            if (InputBox2("Swap Rows", "Row1:", "Row2:", ref row1, ref row2) == DialogResult.OK)
            {
                int row1Index = Int16.Parse(row1);
                int row2Index = Int16.Parse(row2);
                bool result = this.sheet.exchangeRows(row1Index, row2Index);
                int rowCount = dataGridView.Rows.Count;
                if (result)
                {
                    int cols = dataGridView.Rows[row1Index - 1].Cells.Count;
                    DataGridViewRow oldRow1 = dataGridView.Rows[row1Index - 1];
                    DataGridViewRow oldRow2 = dataGridView.Rows[row2Index - 1];

                    DataGridViewRow newRow1 = new DataGridViewRow();
                    newRow1.CreateCells(dataGridView);

                    DataGridViewRow newRow2 = new DataGridViewRow();
                    newRow2.CreateCells(dataGridView);

                    for (int i = 0; i < cols; i++)
                    {
                        newRow1.Cells[i].Value = oldRow2.Cells[i].Value;
                        newRow2.Cells[i].Value = oldRow1.Cells[i].Value;
                    }
                    dataGridView.Rows.Remove(oldRow1);
                    dataGridView.Rows.Insert(row1Index - 1, newRow1);

                    dataGridView.Rows.Remove(oldRow2);
                    dataGridView.Rows.Insert(row2Index - 1, newRow2);
                } else
                {
                    string box_title = "Swap Rows";
                    string box_msg = "Swap Failed";
                    MessageBox.Show(box_msg, box_title);
                }
                for (int i = 0; i < rowCount; i++)
                {
                    dataGridView.Rows[i].HeaderCell.Value = "" + (i + 1);
                }
            }
        }

        private void searchButton_Click(object sender, EventArgs e)
        {
            String value = "Search";
            if (InputBox("Search", "Search For:", ref value) == DialogResult.OK)
            {
                int row = 0;
                int col = 0;

                bool result = this.sheet.searchString(value, ref row, ref col);
                if (result) {
                    dataGridView.ClearSelection();
                    DataGridViewRow currentrow = dataGridView.Rows[row];
                    currentrow.Cells[col].Selected = true;

                    string box_title = "Result";
                    string box_msg = "Found in Row: " + (row + 1) + " Column: " + (col + 1);

                    MessageBox.Show(box_msg, box_title);
                } else
                {
                    string box_title = "Result";
                    string box_msg = "Not Found";
                    MessageBox.Show(box_msg, box_title);
                }

            }
        }

        private void dataGridView_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            String value = (String)dataGridView[e.ColumnIndex, e.RowIndex].Value;
            sheet.setCell(e.RowIndex, e.ColumnIndex, value);
        }

        private void dataGridView_CellClicked(object sender, DataGridViewCellMouseEventArgs e)
        {
            int row = e.RowIndex;
            int col = e.ColumnIndex;
            if (col == -1 && row > -1)
            {
                dataGridView.ClearSelection();
                DataGridViewRow currentrow = dataGridView.Rows[e.RowIndex];
                foreach (DataGridViewCell cell in currentrow.Cells) 
                {
                    cell.Selected = true;
                }
                    

                if (e.Button == MouseButtons.Right)
                {
                    ContextMenuStrip cms = new ContextMenuStrip();
                    selectedRow = row + 1;
                    cms.Items.Add("Add Row", null, AddRow);
                    int y = row * 20;
                    cms.Show(dataGridView, new Point(50, y));

                }
            } else if (col > -1 && row == -1)
            {
                dataGridView.ClearSelection();
                foreach (DataGridViewRow selRow in dataGridView.Rows)
                    selRow.Cells[e.ColumnIndex].Selected = true;
            }

        }

        private void AddRow(object sender, EventArgs e)
        {
            this.sheet.addRow1(selectedRow);

            int rows = this.sheet.getSheet().GetLength(0);
            int cols = this.sheet.getSheet().GetLength(1);

            DataGridViewRow row = new DataGridViewRow();
            row.CreateCells(dataGridView);

            for(int i = 0; i < cols; i++)
            {
                row.Cells[i].Value = this.sheet.getSheet()[selectedRow - 1, i];
            }
            dataGridView.Rows.Insert(selectedRow - 1, row);

            for (int i = 0; i < rows; i++)
            {
                dataGridView.Rows[i].HeaderCell.Value = "" + (i + 1);
            }
        }

        private void dataGridView1_MouseMove(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                // If the mouse moves outside the rectangle, start the drag.
                if (dragBoxFromMouseDown != Rectangle.Empty &&
                    !dragBoxFromMouseDown.Contains(e.X, e.Y))
                {

                    // Proceed with the drag and drop, passing in the list item.                    
                    DragDropEffects dropEffect = dataGridView.DoDragDrop(
                    dataGridView.Rows[rowIndexFromMouseDown],
                    DragDropEffects.Copy);
                }
            }
        }

        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            // Get the index of the item the mouse is below.
            rowIndexFromMouseDown = dataGridView.HitTest(e.X, e.Y).RowIndex;
            if (rowIndexFromMouseDown != -1)
            {
                // Remember the point where the mouse down occurred. 
                // The DragSize indicates the size that the mouse can move 
                // before a drag event should be started.                
                Size dragSize = SystemInformation.DragSize;

                // Create a rectangle using the DragSize, with the mouse position being
                // at the center of the rectangle.
                dragBoxFromMouseDown = new Rectangle(new Point(e.X - (dragSize.Width / 2),
                                                               e.Y - (dragSize.Height / 2)),
                                    dragSize);
            }
            else
                // Reset the rectangle if the mouse is not over an item in the ListBox.
                dragBoxFromMouseDown = Rectangle.Empty;
        }

        private void dataGridView1_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            // The mouse locations are relative to the screen, so they must be 
            // converted to client coordinates.
            Point clientPoint = dataGridView.PointToClient(new Point(e.X, e.Y));

            // Get the row index of the item the mouse is below. 
            rowIndexOfItemUnderMouseToDrop =
                dataGridView.HitTest(clientPoint.X, clientPoint.Y).RowIndex;

            // If the drag operation was a move then remove and insert the row.
            if (e.Effect == DragDropEffects.Move)
            {
                DataGridViewRow rowToMove = e.Data.GetData(
                    typeof(DataGridViewRow)) as DataGridViewRow;
                dataGridView.Rows.RemoveAt(rowIndexFromMouseDown);
                dataGridView.Rows.Insert(rowIndexOfItemUnderMouseToDrop, rowToMove);

            }
        }

        private void SetupDataGridView(int rows, int cols)
        {
            this.Controls.Add(dataGridView);
            dataGridView.AutoGenerateColumns = false;
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            dataGridView.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView.ColumnHeadersDefaultCellStyle.Font = new Font(dataGridView.Font, FontStyle.Bold);

            dataGridView.Name = "Table";
            dataGridView.Location = new Point(8, 8);
            dataGridView.Size = new Size(500, 400);
            dataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            dataGridView.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dataGridView.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridView.RowHeadersVisible = true;
            dataGridView.RowHeadersWidth = 50;
            dataGridView.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dataGridView.MultiSelect = true;
            dataGridView.Dock = DockStyle.Top;
            dataGridView.GridColor = Color.Black;
            dataGridView.ScrollBars = ScrollBars.Both;
            dataGridView.AllowDrop = true;

            dataGridView.CellValidated += new DataGridViewCellEventHandler(dataGridView_CellValidated);
            dataGridView.CellMouseClick += new DataGridViewCellMouseEventHandler(dataGridView_CellClicked);

            DataGridViewColumn[] columns = new DataGridViewColumn[cols];
            for (int i = 0; i < columns.Length; ++i)
            {
                DataGridViewColumn column = new DataGridViewTextBoxColumn();
                column.CellTemplate = new DataGridViewTextBoxCell();
                column.FillWeight = 1;
                column.Name = "Col " + (i + 1);
                column.Frozen = false;
                columns[i] = column;
            }

            dataGridView.Columns.AddRange(columns);

            dataGridView.CellFormatting += new
                DataGridViewCellFormattingEventHandler(
                dataGridView_CellFormatting);
        }

        private void dataGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e != null)
            {
                if (this.dataGridView.Columns[e.ColumnIndex].Name == "Release Date")
                {
                    if (e.Value != null)
                    {
                        try
                        {
                            e.Value = DateTime.Parse(e.Value.ToString())
                                .ToLongDateString();
                            e.FormattingApplied = true;
                        }
                        catch (FormatException)
                        {
                            Console.WriteLine("{0} is not a valid date.", e.Value.ToString());
                        }
                    }
                }
            }
        }

        static DialogResult InputBox(string title, string promptText, ref string value)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }

        static DialogResult InputBox2(string title, string promptText, String promptText2, ref string value1, ref string value2)
        {
            Form form = new Form();
            Label label = new Label();
            Label label2 = new Label();
            TextBox textBox = new TextBox();
            TextBox textBox2 = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            label2.Text = promptText2;
            textBox.Text = value1;
            textBox2.Text = value2;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 200, 13);
            textBox.SetBounds(12, 36, 200, 20);
            label2.SetBounds(129, 20, 200, 13);
            textBox2.SetBounds(132, 36, 200, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            label2.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            textBox2.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, label2, textBox2, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            value1 = textBox.Text;
            value2 = textBox2.Text;
            return dialogResult;
        }
    }

}
