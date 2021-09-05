using System;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using NationalInstruments.NI4882;

namespace Utilities._4194Plot {
    /// <summary>
    /// Summary description for Form1.
    /// </summary>
    public class MainForm : System.Windows.Forms.Form {
        private System.Windows.Forms.Button captureButton;
        private System.Windows.Forms.Button fileButton;
        private Device device;
        private System.Windows.Forms.NumericUpDown gpibAddressNumericUpDown;
        private System.Windows.Forms.Label gpibAddressLabel;

        const int ofsToUnit = 5;    // offset into prefix array below to "" (unit)
        string[] prefix4194 = { "f", "p", "n", "μ", "m", "", "k", "M", "G" };
        string labelX = "", labelY1 = "", labelY2 = "";
        ushort pointCount;

        // Raw 4194 COPY input and array of lines
        string rawInput;
        string[] dataIn;
        char[] arr;
        string s;
        
        // Buffers for numerical data from 4194 table file, large enough to hold
        // the maximum number of points.
        const int defaultArraySize = 401;
        double[] dataX = new double[defaultArraySize];
        double[] dataY1 = new double[defaultArraySize];
        double[] dataY2 = new double[defaultArraySize];

        double priMarkerFreq = 0.0, secMarkerFreq = 0.0;
        double priMarkerAVal = 0.0, secMarkerAVal = 0.0;
        double priMarkerBVal = 0.0, secMarkerBVal = 0.0;

        double multX = 1.0, multY1 = 1.0, multY2 = 1.0;

        StreamWriter swp, swd;

        string temp;
        ushort lineCount;
        private TextBox textBoxPlotFile;
        int i;
        private Label boardNumberLabel;
        private NumericUpDown boardNumericUpDown;
        private Label label1;
        private TextBox textBoxDataFile;
        private Label label2;
        private Button dataButton;

        // ==================================================================================================
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;

        public MainForm() {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();
        }

        // ==================================================================================================
        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing) {
            if (disposing) {
                if (device != null) {
                    device.Dispose();
                }
                if (components != null) {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.captureButton = new System.Windows.Forms.Button();
            this.fileButton = new System.Windows.Forms.Button();
            this.gpibAddressNumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.gpibAddressLabel = new System.Windows.Forms.Label();
            this.textBoxPlotFile = new System.Windows.Forms.TextBox();
            this.boardNumberLabel = new System.Windows.Forms.Label();
            this.boardNumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxDataFile = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dataButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.gpibAddressNumericUpDown)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.boardNumericUpDown)).BeginInit();
            this.SuspendLayout();
            // 
            // captureButton
            // 
            this.captureButton.Location = new System.Drawing.Point(155, 156);
            this.captureButton.Name = "captureButton";
            this.captureButton.Size = new System.Drawing.Size(85, 26);
            this.captureButton.TabIndex = 5;
            this.captureButton.Text = "&Capture";
            this.captureButton.Click += new System.EventHandler(this.captureButton_Click);
            // 
            // fileButton
            // 
            this.fileButton.Image = ((System.Drawing.Image)(resources.GetObject("fileButton.Image")));
            this.fileButton.Location = new System.Drawing.Point(350, 63);
            this.fileButton.Name = "fileButton";
            this.fileButton.Size = new System.Drawing.Size(33, 23);
            this.fileButton.TabIndex = 3;
            this.fileButton.Click += new System.EventHandler(this.fileButton_Click);
            // 
            // gpibAddressNumericUpDown
            // 
            this.gpibAddressNumericUpDown.Location = new System.Drawing.Point(326, 11);
            this.gpibAddressNumericUpDown.Maximum = new decimal(new int[] {
            127,
            0,
            0,
            0});
            this.gpibAddressNumericUpDown.Name = "gpibAddressNumericUpDown";
            this.gpibAddressNumericUpDown.Size = new System.Drawing.Size(40, 20);
            this.gpibAddressNumericUpDown.TabIndex = 2;
            this.gpibAddressNumericUpDown.Value = new decimal(new int[] {
            16,
            0,
            0,
            0});
            this.gpibAddressNumericUpDown.ValueChanged += new System.EventHandler(this.gpibAddressNumericUpDown_ValueChanged);
            // 
            // gpibAddressLabel
            // 
            this.gpibAddressLabel.Location = new System.Drawing.Point(230, 12);
            this.gpibAddressLabel.Name = "gpibAddressLabel";
            this.gpibAddressLabel.Size = new System.Drawing.Size(90, 18);
            this.gpibAddressLabel.TabIndex = 15;
            this.gpibAddressLabel.Text = "GPIB Address:";
            // 
            // textBoxPlotFile
            // 
            this.textBoxPlotFile.AllowDrop = true;
            this.textBoxPlotFile.Location = new System.Drawing.Point(12, 65);
            this.textBoxPlotFile.Name = "textBoxPlotFile";
            this.textBoxPlotFile.Size = new System.Drawing.Size(332, 20);
            this.textBoxPlotFile.TabIndex = 4;
            // 
            // boardNumberLabel
            // 
            this.boardNumberLabel.Location = new System.Drawing.Point(12, 12);
            this.boardNumberLabel.Name = "boardNumberLabel";
            this.boardNumberLabel.Size = new System.Drawing.Size(90, 18);
            this.boardNumberLabel.TabIndex = 18;
            this.boardNumberLabel.Text = "Board Number:";
            // 
            // boardNumericUpDown
            // 
            this.boardNumericUpDown.Location = new System.Drawing.Point(108, 11);
            this.boardNumericUpDown.Maximum = new decimal(new int[] {
            32,
            0,
            0,
            0});
            this.boardNumericUpDown.Name = "boardNumericUpDown";
            this.boardNumericUpDown.Size = new System.Drawing.Size(40, 20);
            this.boardNumericUpDown.TabIndex = 1;
            this.boardNumericUpDown.ValueChanged += new System.EventHandler(this.boardNumericUpDown_ValueChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 46);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Plot File:";
            // 
            // textBoxDataFile
            // 
            this.textBoxDataFile.Location = new System.Drawing.Point(12, 119);
            this.textBoxDataFile.Name = "textBoxDataFile";
            this.textBoxDataFile.Size = new System.Drawing.Size(332, 20);
            this.textBoxDataFile.TabIndex = 19;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 100);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(52, 13);
            this.label2.TabIndex = 20;
            this.label2.Text = "Data File:";
            // 
            // dataButton
            // 
            this.dataButton.Image = ((System.Drawing.Image)(resources.GetObject("dataButton.Image")));
            this.dataButton.Location = new System.Drawing.Point(350, 116);
            this.dataButton.Name = "dataButton";
            this.dataButton.Size = new System.Drawing.Size(32, 23);
            this.dataButton.TabIndex = 21;
            this.dataButton.UseVisualStyleBackColor = true;
            this.dataButton.Click += new System.EventHandler(this.dataButton_Click);
            // 
            // MainForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(400, 194);
            this.Controls.Add(this.dataButton);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBoxDataFile);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.boardNumberLabel);
            this.Controls.Add(this.boardNumericUpDown);
            this.Controls.Add(this.textBoxPlotFile);
            this.Controls.Add(this.gpibAddressLabel);
            this.Controls.Add(this.gpibAddressNumericUpDown);
            this.Controls.Add(this.fileButton);
            this.Controls.Add(this.captureButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(184, 100);
            this.Name = "MainForm";
            this.Text = "4194 Plot";
            this.Load += new System.EventHandler(this.MainForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.gpibAddressNumericUpDown)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.boardNumericUpDown)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main() {
            Application.Run(new MainForm());
        }

        private void MainForm_Load(object sender, System.EventArgs e) {
            disableControls(false);
            // Retrieve file names from settings
            textBoxPlotFile.Text = Properties.Settings.Default.FilePath;
            textBoxDataFile.Text = Properties.Settings.Default.DataPath;
            // Retrieve board and GPIB addresses from settings
            boardNumericUpDown.Value = Properties.Settings.Default.BoardNum;
            gpibAddressNumericUpDown.Value = Properties.Settings.Default.GpibAddress;
        }

        // ==================================================================================================
        // TRUE to disable all controls while capture is in operation
        //
        private void disableControls(bool isSessionOpen) {
            gpibAddressNumericUpDown.Enabled = !isSessionOpen;
            boardNumericUpDown.Enabled = !isSessionOpen;
            captureButton.Enabled = !isSessionOpen;
            dataButton.Enabled = !isSessionOpen;
            fileButton.Enabled = !isSessionOpen;
            textBoxPlotFile.Enabled = !isSessionOpen;
        }

        // ==================================================================================================
        // Utility to read a double value from the 4194
        //
        private double getDouble(string message) {
            device.Write(message);
            return Convert.ToDouble(device.ReadString());
        }

        // ==================================================================================================
        // Limit an int between max and min values
        //
        private int Limit(int val, int min, int max) {
            if (val > max)
                val = max;
            if (val < min)
                val = min;
            return val;
        }

        // ==================================================================================================
        // Return an exponent string to concatenate to the mantissa if a multiplier
        // char is provided
        //
        private string GetMult(char mult) {
            switch (mult) {
                case 'f':
                    return ("e-15");
                case 'p':
                    return ("e-12");
                case 'n':
                    return ("e-9");
                case 'u':
                    return ("e-6");
                case 'm':
                    return ("e-3");
                case 'K':
                    return ("e+3");
                case 'M':
                    return ("e+6");
                default:
                    return ("");
            }
        }

        // ==================================================================================================
        // Get labels
        // The 4194 label string is in a fixed format:
        //      X  starts at char offset  3, extends 18 chars
        //      Y1 starts at char offset 21, extends 16 chars
        //      Y2 starts at char offset 37, extends 15 chars
        // This routine extracts and formats individual label strings from the 4194 label string
        // suitable for axis labels.
        //
        private void getLabels(string s) {
            // Get X label, trim whitespace
            labelX = s.Substring(3, 18);
            labelX = labelX.Trim();
            // Remove any spaces inside unit brackets
            if (labelX.Contains("[")) {
                i = labelX.IndexOf('[');
                if (labelX[i + 1] == ' ')
                    labelX = labelX.Substring(0, i + 1) + labelX.Substring(i + 2);
            }
            if (labelX.Contains("]")) {
                i = labelX.IndexOf(']');
                if (labelX[i - 1] == ' ')
                    labelX = labelX.Substring(0, i - 1) + labelX.Substring(i);
            }

            // Get Y1 label, trim whitespace, replace "ohm"
            labelY1 = s.Substring(21, 16);
            labelY1 = labelY1.Trim();
            labelY1 = labelY1.Replace("ohm", "Ω");
            // Underscores are interpreted by Veusz as "switch to subscript"
            labelY1 = labelY1.Replace("_", "");
            // Remove any spaces inside unit brackets
            if (labelY1.Contains("[")) {
                i = labelY1.IndexOf('[');
                if (labelY1[i + 1] == ' ')
                    labelY1 = labelY1.Substring(0, i + 1) + labelY1.Substring(i + 2);
            }
            if (labelY1.Contains("]")) {
                i = labelY1.IndexOf(']');
                if (labelY1[i - 1] == ' ')
                    labelY1 = labelY1.Substring(0, i - 1) + labelY1.Substring(i);
            }

            // Get Y2 label, trim whitespace, replace "ohm"
            labelY2 = s.Substring(37, 15);
            labelY2 = labelY2.Trim();
            labelY2 = labelY2.Replace("ohm", "Ω");
            // Underscores are interpreted by Veusz as "switch to subscript"
            labelY2 = labelY2.Replace("_", "");
            // Remove any spaces inside unit brackets
            if (labelY2.Contains("[")) {
                i = labelY2.IndexOf('[');
                if (labelY2[i + 1] == ' ')
                    labelY2 = labelY2.Substring(0, i + 1) + labelY2.Substring(i + 2);
            }
            if (labelY2.Contains("]")) {
                i = labelY2.IndexOf(']');
                if (labelY2[i - 1] == ' ')
                    labelY2 = labelY2.Substring(0, i - 1) + labelY2.Substring(i);
            }
        }

        // ==================================================================================================
        // Scale data and markers to engineering units, insert multiplier prefix into axis
        // label strings if req'd
        //
        private void scaleData() {
            int ip;
            int exp;
            double max;

            // Adjust X display unit prefix
            // Get largest magnitude value in the data
            max = Math.Max(Math.Abs(dataX.Max()), Math.Abs(dataX.Min()));
            // Log and truncate toward negative inf
            exp = (int)(Math.Floor(Math.Log10(max)));
            // Get index which groups integers into correct groups of three
            ip = exp >= 0 ? (exp / 3) + ofsToUnit : ((exp - 2) / 3) + ofsToUnit;
            ip = Limit(ip, 0, prefix4194.Length - 1);
            i = labelX.IndexOf('[') + 1;
            labelX = labelX.Insert(i, prefix4194[ip]);
            // Calculate required data multiplier
            exp = 3 * (ofsToUnit - ip);
            multX = Math.Pow(10.0, (double)exp);

            // Adjust Y1 display unit prefix
            // Get largest magnitude value in the data
            max = Math.Max(Math.Abs(dataY1.Max()), Math.Abs(dataY1.Min()));
            // Log and truncate toward negative inf
            exp = (int)(Math.Floor(Math.Log10(max)));
            // Get index which groups integers into correct groups of three
            ip = exp >= 0 ? (exp / 3) + ofsToUnit : ((exp - 2) / 3) + ofsToUnit;
            ip = Limit(ip, 0, prefix4194.Length - 1);
            i = labelY1.IndexOf('[') + 1;
            labelY1 = labelY1.Insert(i, prefix4194[ip]);
            labelY1 = labelY1.Replace(" [ ]", "");
            // Calculate required data multiplier
            exp = 3 * (ofsToUnit - ip);
            multY1 = Math.Pow(10.0, (double)exp);

            // Adjust Y2 display unit prefix
            if (labelY2.Contains("PHASE"))
            {
                // Only allow mdeg or deg for phase
                ip = Limit(ip, ofsToUnit - 1, ofsToUnit);
            }
            else if (labelY2.Contains("Q "))
            {
                // Q is only in units
                ip = ofsToUnit;
            }
            else
            {
                // Get largest magnitude value in the data
                max = Math.Max(Math.Abs(dataY2.Max()), Math.Abs(dataY2.Min()));
                // Log and truncate toward negative inf
                exp = (int)(Math.Floor(Math.Log10(max)));
                // Get index which groups integers into correct groups of three
                ip = exp >= 0 ? (exp / 3) + ofsToUnit : ((exp - 2) / 3) + ofsToUnit;
                ip = Limit(ip, 0, prefix4194.Length - 1);
            }
            i = labelY2.IndexOf('[') + 1;
            labelY2 = labelY2.Insert(i, prefix4194[ip]);
            labelY2 = labelY2.Replace(" [ ]", "");
            // Calculate required data multiplier
            exp = 3 * (ofsToUnit - ip);
            multY2 = Math.Pow(10.0, (double)exp);

            // Scale the numbers
            for (int j = 0; j < pointCount; j++) {
                dataX[j] = dataX[j] * multX;
                dataY1[j] = dataY1[j] * multY1;
                dataY2[j] = dataY2[j] * multY2;
            }

            // Scale markers
            priMarkerFreq *= multX;
            secMarkerFreq *= multX;
            priMarkerAVal *= multY1;
            secMarkerAVal *= multY1;
            priMarkerBVal *= multY2;
            secMarkerBVal *= multY2;
        }

        // ==================================================================================================
        // Capture button initiates comm with 4194 and writes output to a text
        // file suitable for importing to Veusz plotting software.
        //
        private void captureButton_Click(object sender, System.EventArgs e) {

            try {
                // First check for valid output file
                swp = new StreamWriter(textBoxPlotFile.Text);
                swd = new StreamWriter(textBoxDataFile.Text);
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
                return;
            }

            Cursor.Current = Cursors.WaitCursor;
                
            // Intialize the device
            // TBD: catch error?  (This is difficult.  Creating a Board and running Board.FindListener() does 
            // does not always work.  4194 does not respond to IDN? message anyway, so it doesn't
            // seem possible to verify that a 4194 is actually at the selected address
            // anyway, even if a listener is found.)
            device = new Device(Convert.ToInt16(boardNumericUpDown.Value),
                                Convert.ToByte(gpibAddressNumericUpDown.Value));
            disableControls(true);

            // Get primary marker data
            priMarkerFreq = getDouble("MKR?\n");
            priMarkerAVal = getDouble("MKRA?\n");
            priMarkerBVal = getDouble("MKRB?\n");

            // Get secondary marker data
            secMarkerFreq = getDouble("SMKR?\n");
            secMarkerAVal = getDouble("SMKRA?\n");
            secMarkerBVal = getDouble("SMKRB?\n");

            // Set copy mode 2 and read raw data
            device.Write("CPYM2\n");
            device.Write("COPY\n");
            rawInput = device.ReadString(50000);    // up to N chars
                
            // Convert to string array of lines
            dataIn = rawInput.Split(new Char[] { '\r' });

            // Read dataIn for labels and to convert data to double arrays
            lineCount = 0;
            pointCount = 0;

            foreach (string raw in dataIn) {
                // Clean the string of non-printable ASCII chars
                arr = raw.Where(c => ((c >= 0x20) && (c <= 0x7E))).ToArray();
                s = new string(arr);

                lineCount++;
                if (lineCount == 3) {
                    // Get X, Y1, and Y2 labels
                    getLabels(s);
                }
                else {

                    if (s.Length < 40) 
                        continue;

                    // Typ line of data & col index:
                    //           111111111122222222223333333333444444444455
                    // 0123456789012345678901234567890123456789012345678901
                    //   1           100.000      2.12511  G    -26.6899  m

                    // Get X, trim, eliminate embedded spaces, convert to double
                    temp = s.Substring(3, 18);
                    temp = temp.Trim();
                    temp = temp.Replace(" ", string.Empty);
                    dataX[pointCount] = Convert.ToDouble(temp);

                    // Get Y1 data point & multiplier, trim, and convert to double
                    temp = s.Substring(21, 13);
                    temp = temp.Trim();
                    temp = string.Concat(temp, GetMult(s[36]));
                    dataY1[pointCount] = Convert.ToDouble(temp);

                    // Get Y2 data point & multiplier, trim, and convert to double
                    temp = s.Substring(37, 12);
                    temp = temp.Trim();
                    temp = string.Concat(temp, GetMult(s[51]));
                    dataY2[pointCount] = Convert.ToDouble(temp);

                    pointCount++;
                }
            }
            // We now have arrays of doubles from original 4194 data file in 
            // dataX, dataY1, and dataY2 arrays.  Resize to actual data size.
            Array.Resize(ref dataX, pointCount);
            Array.Resize(ref dataY1, pointCount);
            Array.Resize(ref dataY2, pointCount);

            // Write data file: one header line followed by unscaled CSV data
            swd.WriteLine("{0},{1},{2}", labelX, labelY1, labelY2);
            for (i = 0; i < pointCount; i++) {
                swd.WriteLine("{0},{1},{2}", dataX[i], dataY1[i], dataY2[i]);
            }
            swd.Flush();
            swd.Dispose();

            // Determine engineering unit scaling for all axes and adjust data & markers
            scaleData();

            // We now have axis label strings and scaled numbers.  Write Veusz descriptor
            // lines followed by space-separated data.
            // Primary marker:
            swp.WriteLine("descriptor pmkr pmkr_y1 pmkr_y2");
            swp.WriteLine("{0} {1} {2}", priMarkerFreq, priMarkerAVal, priMarkerBVal);
            // Secondary marker:
            swp.WriteLine("descriptor smkr smkr_y1 smkr_y2");
            swp.WriteLine("{0} {1} {2}", secMarkerFreq, secMarkerAVal, secMarkerBVal);
            // Axi labels (use 'u' prefix to quoted strings for unicode chars):
            swp.WriteLine("descriptor label_x(text) label_y1(text) label_y2(text)");
            swp.WriteLine("u\"{0}\" u\"{1}\" u\"{2}\"", labelX, labelY1, labelY2);
            // Data:
            swp.WriteLine("descriptor data_x data_y1 data_y2");
            for (i = 0; i < pointCount; i++) {
                swp.WriteLine("{0} {1} {2}", dataX[i], dataY1[i], dataY2[i]);
            }
            swp.Flush();
            swp.Dispose();
                
            // Close
            device.GoToLocal();
            device.Dispose();
            disableControls(false);
            Cursor.Current = Cursors.Default;
        }

        private void fileButton_Click(object sender, EventArgs e) {
            // Create OpenFileDialog 
            OpenFileDialog dlg = new OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".txt";
            dlg.Filter = "Text Files (*.txt)|*.txt";
            dlg.AddExtension = true;
            dlg.CheckFileExists = false;
            dlg.DefaultExt = ".txt";
            dlg.ReadOnlyChecked = false;

            // Display OpenFileDialog by calling ShowDialog method 
            DialogResult result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox 
            if (result == DialogResult.OK) {
                // Update file name and save in permanent settings
                textBoxPlotFile.Text = dlg.FileName;
                Properties.Settings.Default.FilePath = textBoxPlotFile.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void boardNumericUpDown_ValueChanged(object sender, EventArgs e) {
            Properties.Settings.Default.BoardNum = boardNumericUpDown.Value;
            Properties.Settings.Default.Save();
        }

        private void gpibAddressNumericUpDown_ValueChanged(object sender, EventArgs e) {
            Properties.Settings.Default.GpibAddress = gpibAddressNumericUpDown.Value;
            Properties.Settings.Default.Save();
        }

        private void dataButton_Click(object sender, EventArgs e) {
            // Create OpenFileDialog 
            OpenFileDialog dlg = new OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".csv";
            dlg.Filter = "CSV Files (*.csv)|*.csv";
            dlg.AddExtension = true;
            dlg.CheckFileExists = false;
            dlg.DefaultExt = ".csv";
            dlg.ReadOnlyChecked = false;

            // Display OpenFileDialog by calling ShowDialog method 
            DialogResult result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox 
            if (result == DialogResult.OK) {
                // Update file name and save in permanent settings
                textBoxDataFile.Text = dlg.FileName;
                Properties.Settings.Default.DataPath = textBoxDataFile.Text;
                Properties.Settings.Default.Save();
            }
        }
    }
}
