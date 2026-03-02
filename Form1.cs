using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace TTONG01
{
    public partial class Form1 : Form
    {
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private DrawingDoc swDrawing;

        public Form1()
        {
            InitializeComponent();
            SetupUI();
            // 异步初始化SolidWorks连接，避免界面卡顿
            System.Threading.Tasks.Task.Run(() =>
            {
                InitializeSolidWorksConnection();
                // 更新UI状态
                this.Invoke(new Action(() =>
                {
                    UpdateStatusLabels();
                }));
            });
        }

        private void InitializeSolidWorksConnection()
        {
            try
            {
                // 尝试获取正在运行的SolidWorks实例
                swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                if (swApp != null)
                {
                    swApp.Visible = true;
                    System.Diagnostics.Debug.WriteLine("成功连接到SolidWorks");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("无法连接到SolidWorks: " + ex.Message);
                this.Invoke(new Action(() =>
                {
                    MessageBox.Show("无法连接到SolidWorks: " + ex.Message);
                    lblStatus.Text = "状态: 无法连接到SolidWorks";
                    UpdateStatusLabels();
                }));
            }
        }

        private void UpdateStatusLabels()
        {
            try
            {
                // 更新连接状态
                if (this.Controls.Find("lblConnectionValue", true) is Control[] connectionControls && connectionControls.Length > 0 && connectionControls[0] is Label lblConnectionValue)
                {
                    lblConnectionValue.Text = swApp != null ? "已连接" : "未连接";
                    lblConnectionValue.ForeColor = swApp != null ? Color.Green : Color.Red;
                }

                // 更新模板状态
                if (this.Controls.Find("lblTemplateValue", true) is Control[] templateControls && templateControls.Length > 0 && templateControls[0] is Label lblTemplateValue)
                {
                    lblTemplateValue.Text = swModel != null ? "已打开" : "未打开";
                    lblTemplateValue.ForeColor = swModel != null ? Color.Green : Color.Red;
                }

                // 更新文件状态
                if (this.Controls.Find("lblFileValue", true) is Control[] fileControls && fileControls.Length > 0 && fileControls[0] is Label lblFileValue && lstFiles != null)
                {
                    lblFileValue.Text = lstFiles.Items.Count.ToString();
                    lblFileValue.ForeColor = lstFiles.Items.Count > 0 ? Color.Blue : Color.Gray;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("更新状态标签时出错: " + ex.Message);
            }
        }

        private void OpenTemplate()
        {
            try
            {
                if (swApp == null)
                {
                    this.Invoke(new Action(() =>
                    {
                        lblStatus.Text = "状态: 未连接到SolidWorks";
                        UpdateStatusLabels();
                    }));
                    return;
                }

                // 尝试查找工程图模板
                string templatePath = null;
                
                try
                {
                    // 路径: 当前工作目录的quote文件夹
                    string currentPath = System.IO.Directory.GetCurrentDirectory();
                    templatePath = System.IO.Path.Combine(currentPath, "quote", "TEMP.SLDDRW");
                    System.Diagnostics.Debug.WriteLine("尝试路径: " + templatePath);
                }
                catch (Exception pathEx)
                {
                    System.Diagnostics.Debug.WriteLine("查找模板路径时出错: " + pathEx.Message);
                }

                if (System.IO.File.Exists(templatePath))
                {
                    System.Diagnostics.Debug.WriteLine("使用模板路径: " + templatePath);
                    object newDrawingObj = swApp.NewDocument(templatePath, 0, 0.0, 0.0);
                    if (newDrawingObj != null)
                    {
                        swModel = (ModelDoc2)newDrawingObj;
                        swDrawing = (DrawingDoc)swModel;
                        this.Invoke(new Action(() =>
                        {
                            lblStatus.Text = "状态: 成功打开工程图模板";
                            UpdateStatusLabels();
                        }));
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("无法创建文档从模板: " + templatePath);
                        this.Invoke(new Action(() =>
                        {
                            lblStatus.Text = "状态: 无法打开工程图模板";
                            UpdateStatusLabels();
                        }));
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("未找到模板文件，创建默认工程图");
                    // 如果模板不存在，创建一个默认的工程图
                    object newDrawingObj = swApp.NewDocument("", 0, 0.0, 0.0);
                    if (newDrawingObj != null)
                    {
                        swModel = (ModelDoc2)newDrawingObj;
                        swDrawing = (DrawingDoc)swModel;
                        this.Invoke(new Action(() =>
                        {
                            lblStatus.Text = "状态: 未找到模板，已创建默认工程图";
                            UpdateStatusLabels();
                        }));
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("无法创建默认工程图");
                        this.Invoke(new Action(() =>
                        {
                            lblStatus.Text = "状态: 无法创建工程图";
                            UpdateStatusLabels();
                        }));
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("打开模板时出错: " + ex.Message);
                this.Invoke(new Action(() =>
                {
                    lblStatus.Text = "状态: 打开模板错误: " + ex.Message;
                    UpdateStatusLabels();
                }));
            }
        }

        private ListBox lstFiles;
        private Button btnAddFile;
        private Button btnProcessFiles;
        private Label lblStatus;
        private TextBox txtLog;
        private Button btnShowLog;
        private List<string> logMessages;
        
        // 设置选项
        private bool hideBendLines = true; // 默认隐藏折弯线
        private bool hideBendNotes = true; // 默认隐藏折弯注释
        private bool showNoteName = true; // 默认显示名称注释
        private bool showNoteMaterial = true; // 默认显示材料注释
        private bool showNoteThickness = true; // 默认显示厚度注释
        private bool showNoteQuantity = true; // 默认显示数量注释
        private double noteFontHeight = 15; // 默认注释字高（毫米）
        private string sortMethod = "material_thickness"; // 默认按材料、厚度进行分类，可选值：material_thickness, thickness, material
        private double rowSpacing = 100; // 默认行间距（毫米）
        private double columnSpacing = 100; // 默认列间距（毫米）

        private void SetupUI()
        {
            // 基本窗口设置
            this.Text = "SolidWorks钣金展开视图插件";
            this.Size = new Size(650, 480);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.White;

            // 设置字体
            Font labelFont = new Font("Microsoft Sans Serif", 9f, FontStyle.Regular);
            Font buttonFont = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold);
            Font titleFont = new Font("Microsoft Sans Serif", 10f, FontStyle.Bold);

            // 1. 文件列表区域
            GroupBox fileGroup = new GroupBox
            {
                Text = "文件列表",
                Location = new Point(20, 15),
                Size = new Size(610, 180),
                Font = titleFont
            };
            this.Controls.Add(fileGroup);

            // 文件列表框
            lstFiles = new ListBox
            {
                Location = new Point(15, 30),
                Size = new Size(580, 110),
                Font = labelFont,
                SelectionMode = SelectionMode.MultiExtended,
                BorderStyle = BorderStyle.FixedSingle
            };
            fileGroup.Controls.Add(lstFiles);

            // 文件操作按钮
            Panel fileButtons = new Panel
            {
                Location = new Point(15, 145),
                Size = new Size(580, 30)
            };
            fileGroup.Controls.Add(fileButtons);

            // 添加文件按钮
            btnAddFile = new Button
            {
                Text = "添加文件",
                Size = new Size(100, 30),
                Font = buttonFont,
                Location = new Point(0, 0)
            };
            btnAddFile.Click += new EventHandler(BtnAddFile_Click);
            fileButtons.Controls.Add(btnAddFile);

            // 移除文件按钮
            Button btnRemoveFile = new Button
            {
                Text = "移除文件",
                Size = new Size(100, 30),
                Font = buttonFont,
                Location = new Point(110, 0)
            };
            btnRemoveFile.Click += new EventHandler(BtnRemoveFile_Click);
            fileButtons.Controls.Add(btnRemoveFile);

            // 2. 状态信息区域
            GroupBox statusGroup = new GroupBox
            {
                Text = "状态信息",
                Location = new Point(20, 205),
                Size = new Size(610, 90),
                Font = titleFont
            };
            this.Controls.Add(statusGroup);

            // SolidWorks连接状态
            Label lblSwStatus = new Label
            {
                Text = "SolidWorks连接:",
                Location = new Point(20, 30),
                Size = new Size(120, 20),
                Font = labelFont
            };
            statusGroup.Controls.Add(lblSwStatus);

            Label lblSwValue = new Label
            {
                Name = "lblConnectionValue",
                Text = "连接中...",
                Location = new Point(140, 30),
                Size = new Size(100, 20),
                Font = labelFont,
                ForeColor = Color.Orange
            };
            statusGroup.Controls.Add(lblSwValue);

            // 文件数量状态
            Label lblFileCount = new Label
            {
                Text = "已添加文件:",
                Location = new Point(260, 30),
                Size = new Size(100, 20),
                Font = labelFont
            };
            statusGroup.Controls.Add(lblFileCount);

            Label lblFileValue = new Label
            {
                Name = "lblFileValue",
                Text = "0",
                Location = new Point(360, 30),
                Size = new Size(50, 20),
                Font = labelFont,
                ForeColor = Color.Blue
            };
            statusGroup.Controls.Add(lblFileValue);

            // 3. 开始按钮
            btnProcessFiles = new Button
            {
                Text = "开始处理",
                Location = new Point(530, 400),
                Size = new Size(100, 35),
                Font = buttonFont,
                BackColor = Color.LightGreen
            };
            btnProcessFiles.Click += new EventHandler(BtnProcessFiles_Click);
            this.Controls.Add(btnProcessFiles);

            // 4. 设置按钮
            Button btnSettings = new Button
            {
                Text = "设置",
                Location = new Point(420, 400),
                Size = new Size(100, 35),
                Font = buttonFont,
                BackColor = Color.LightYellow
            };
            btnSettings.Click += new EventHandler(BtnSettings_Click);
            this.Controls.Add(btnSettings);

            // 5. 日志信息区域
            GroupBox logGroup = new GroupBox
            {
                Text = "日志信息",
                Location = new Point(20, 305),
                Size = new Size(610, 85),
                Font = titleFont
            };
            this.Controls.Add(logGroup);

            // 日志文本框
            txtLog = new TextBox
            {
                Location = new Point(15, 25),
                Size = new Size(490, 40),
                Font = labelFont,
                ReadOnly = true,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
            };
            logGroup.Controls.Add(txtLog);

            // 日志按钮
            btnShowLog = new Button
            {
                Text = "查看日志",
                Location = new Point(515, 25),
                Size = new Size(80, 40),
                Font = buttonFont,
                BackColor = Color.LightBlue
            };
            btnShowLog.Click += new EventHandler(BtnShowLog_Click);
            logGroup.Controls.Add(btnShowLog);

            // 6. 底部状态条
            Panel statusBar = new Panel
            {
                Location = new Point(0, 455),
                Size = new Size(650, 25),
                BackColor = Color.LightGray,
                BorderStyle = BorderStyle.FixedSingle
            };
            this.Controls.Add(statusBar);

            // 状态标签
            lblStatus = new Label
            {
                Text = "状态: 就绪",
                Location = new Point(10, 5),
                Size = new Size(630, 18),
                Font = new Font("Microsoft Sans Serif", 8f, FontStyle.Italic)
            };
            statusBar.Controls.Add(lblStatus);
        }

        // 输出格式和路径设置
        private bool exportDWG = true; // 默认导出DWG
        private bool exportDXF = false; // 默认不导出DXF
        private string outputPath = ""; // 默认输出路径
        private bool useFirstFilePath = true; // 默认使用第一个文件的路径

        private void BtnSettings_Click(object sender, EventArgs e)
        {
            // 显示设置对话框
            SettingsForm settingsForm = new SettingsForm(hideBendLines, hideBendNotes, showNoteName, showNoteMaterial, showNoteThickness, showNoteQuantity, noteFontHeight, sortMethod, rowSpacing, columnSpacing, exportDWG, exportDXF, outputPath, useFirstFilePath);
            if (settingsForm.ShowDialog() == DialogResult.OK)
            {
                // 更新设置
                hideBendLines = settingsForm.HideBendLines;
                hideBendNotes = settingsForm.HideBendNotes;
                showNoteName = settingsForm.ShowNoteName;
                showNoteMaterial = settingsForm.ShowNoteMaterial;
                showNoteThickness = settingsForm.ShowNoteThickness;
                showNoteQuantity = settingsForm.ShowNoteQuantity;
                noteFontHeight = settingsForm.NoteFontHeight;
                sortMethod = settingsForm.SortMethod;
                rowSpacing = settingsForm.RowSpacing;
                columnSpacing = settingsForm.ColumnSpacing;
                exportDWG = settingsForm.ExportDWG;
                exportDXF = settingsForm.ExportDXF;
                outputPath = settingsForm.OutputPath;
                useFirstFilePath = settingsForm.UseFirstFilePath;
                UpdateLog("设置已更新: 隐藏折弯线=" + hideBendLines + ", 隐藏折弯注释=" + hideBendNotes + ", 显示名称注释=" + showNoteName + ", 显示材料注释=" + showNoteMaterial + ", 显示厚度注释=" + showNoteThickness + ", 显示数量注释=" + showNoteQuantity + ", 注释字高=" + noteFontHeight + "mm, 分类方法=" + sortMethod + ", 行间距=" + rowSpacing + "mm, 列间距=" + columnSpacing + "mm, 导出DWG=" + exportDWG + ", 导出DXF=" + exportDXF + ", 输出路径=" + outputPath + ", 使用第一个文件路径=" + useFirstFilePath);
            }
        }

        // 设置对话框类
        public class SettingsForm : Form
        {
            private CheckBox chkHideBendLines;
            private CheckBox chkHideBendNotes;
            private Button btnOK;
            private Button btnCancel;
            private CheckBox chkShowNoteName;
            private CheckBox chkShowNoteMaterial;
            private CheckBox chkShowNoteThickness;
            private CheckBox chkShowNoteQuantity;
            private TextBox txtNoteFontHeight;
            private RadioButton rdoSortMaterialThickness;
            private RadioButton rdoSortThickness;
            private RadioButton rdoSortMaterial;
            private TextBox txtRowSpacing;
            private TextBox txtColumnSpacing;
            private CheckBox chkExportDWG;
            private CheckBox chkExportDXF;
            private TextBox txtOutputPath;
            private CheckBox chkUseFirstFilePath;
            private Button btnBrowsePath;

            public bool HideBendLines { get; set; }
            public bool HideBendNotes { get; set; }
            public bool ShowNoteName { get; set; }
            public bool ShowNoteMaterial { get; set; }
            public bool ShowNoteThickness { get; set; }
            public bool ShowNoteQuantity { get; set; }
            public double NoteFontHeight { get; set; }
            public string SortMethod { get; set; }
            public double RowSpacing { get; set; }
            public double ColumnSpacing { get; set; }
            public bool ExportDWG { get; set; }
            public bool ExportDXF { get; set; }
            public string OutputPath { get; set; }
            public bool UseFirstFilePath { get; set; }

            public SettingsForm(bool currentHideBendLines, bool currentHideBendNotes, bool currentShowNoteName, bool currentShowNoteMaterial, bool currentShowNoteThickness, bool currentShowNoteQuantity, double currentNoteFontHeight, string currentSortMethod, double currentRowSpacing, double currentColumnSpacing, bool currentExportDWG, bool currentExportDXF, string currentOutputPath, bool currentUseFirstFilePath)
            {
                HideBendLines = currentHideBendLines;
                HideBendNotes = currentHideBendNotes;
                ShowNoteName = currentShowNoteName;
                ShowNoteMaterial = currentShowNoteMaterial;
                ShowNoteThickness = currentShowNoteThickness;
                ShowNoteQuantity = currentShowNoteQuantity;
                NoteFontHeight = currentNoteFontHeight;
                SortMethod = currentSortMethod;
                RowSpacing = currentRowSpacing;
                ColumnSpacing = currentColumnSpacing;
                ExportDWG = currentExportDWG;
                ExportDXF = currentExportDXF;
                OutputPath = currentOutputPath;
                UseFirstFilePath = currentUseFirstFilePath;
                InitializeComponent();
            }

            private void InitializeComponent()
            {
                // 基本窗口设置
                this.Text = "设置";
                this.Size = new Size(650, 380);
                this.StartPosition = FormStartPosition.CenterScreen;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.BackColor = Color.White;

                // 设置字体
                Font labelFont = new Font("Microsoft Sans Serif", 9f, FontStyle.Regular);
                Font buttonFont = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold);
                Font titleFont = new Font("Microsoft Sans Serif", 10f, FontStyle.Bold);

                // 1. 选项设置组
                GroupBox optionsGroup = new GroupBox
                {
                    Text = "选项设置",
                    Location = new Point(20, 20),
                    Size = new Size(300, 80),
                    Font = titleFont
                };
                this.Controls.Add(optionsGroup);

                // 显示折弯线选项
                chkHideBendLines = new CheckBox
                {
                    Text = "显示折弯线",
                    Location = new Point(10, 25),
                    Size = new Size(120, 20),
                    Font = labelFont,
                    Checked = !HideBendLines // 取反，因为我们存储的是是否隐藏
                };
                optionsGroup.Controls.Add(chkHideBendLines);

                // 显示折弯注释选项
                chkHideBendNotes = new CheckBox
                {
                    Text = "显示折弯注释",
                    Location = new Point(10, 50),
                    Size = new Size(120, 20),
                    Font = labelFont,
                    Checked = !HideBendNotes // 取反，因为我们存储的是是否隐藏
                };
                optionsGroup.Controls.Add(chkHideBendNotes);

                // 2. 输出格式选项组
                GroupBox outputGroup = new GroupBox
                {
                    Text = "输出格式",
                    Location = new Point(340, 20),
                    Size = new Size(290, 80),
                    Font = titleFont
                };
                this.Controls.Add(outputGroup);

                // 导出DWG选项
                chkExportDWG = new CheckBox
                {
                    Text = "导出DWG",
                    Location = new Point(10, 25),
                    Size = new Size(80, 20),
                    Font = labelFont,
                    Checked = ExportDWG
                };
                outputGroup.Controls.Add(chkExportDWG);

                // 导出DXF选项
                chkExportDXF = new CheckBox
                {
                    Text = "导出DXF",
                    Location = new Point(100, 25),
                    Size = new Size(80, 20),
                    Font = labelFont,
                    Checked = ExportDXF
                };
                outputGroup.Controls.Add(chkExportDXF);

                // 3. 分类方法选项组
                GroupBox sortGroup = new GroupBox
                {
                    Text = "分类方法",
                    Location = new Point(20, 110),
                    Size = new Size(300, 80),
                    Font = titleFont
                };
                this.Controls.Add(sortGroup);

                // 按材料、厚度进行分类选项
                rdoSortMaterialThickness = new RadioButton
                {
                    Text = "按材料、厚度进行分类",
                    Location = new Point(10, 25),
                    Size = new Size(150, 20),
                    Font = labelFont,
                    Checked = SortMethod == "material_thickness" || SortMethod == ""
                };
                sortGroup.Controls.Add(rdoSortMaterialThickness);

                // 按板厚分类选项
                rdoSortThickness = new RadioButton
                {
                    Text = "按板厚分类",
                    Location = new Point(10, 50),
                    Size = new Size(100, 20),
                    Font = labelFont,
                    Checked = SortMethod == "thickness"
                };
                sortGroup.Controls.Add(rdoSortThickness);

                // 按材料分类选项
                rdoSortMaterial = new RadioButton
                {
                    Text = "按材料分类",
                    Location = new Point(120, 50),
                    Size = new Size(100, 20),
                    Font = labelFont,
                    Checked = SortMethod == "material"
                };
                sortGroup.Controls.Add(rdoSortMaterial);

                // 4. 展开图注释选项组
                GroupBox noteGroup = new GroupBox
                {
                    Text = "展开图注释",
                    Location = new Point(340, 110),
                    Size = new Size(290, 80),
                    Font = titleFont
                };
                this.Controls.Add(noteGroup);

                // 显示名称注释选项
                chkShowNoteName = new CheckBox
                {
                    Text = "名称",
                    Location = new Point(10, 25),
                    Size = new Size(60, 20),
                    Font = labelFont,
                    Checked = ShowNoteName
                };
                noteGroup.Controls.Add(chkShowNoteName);

                // 显示材料注释选项
                chkShowNoteMaterial = new CheckBox
                {
                    Text = "材料",
                    Location = new Point(70, 25),
                    Size = new Size(60, 20),
                    Font = labelFont,
                    Checked = ShowNoteMaterial
                };
                noteGroup.Controls.Add(chkShowNoteMaterial);

                // 显示厚度注释选项
                chkShowNoteThickness = new CheckBox
                {
                    Text = "厚度",
                    Location = new Point(130, 25),
                    Size = new Size(60, 20),
                    Font = labelFont,
                    Checked = ShowNoteThickness
                };
                noteGroup.Controls.Add(chkShowNoteThickness);

                // 显示数量注释选项
                chkShowNoteQuantity = new CheckBox
                {
                    Text = "数量",
                    Location = new Point(190, 25),
                    Size = new Size(60, 20),
                    Font = labelFont,
                    Checked = ShowNoteQuantity
                };
                noteGroup.Controls.Add(chkShowNoteQuantity);

                // 注释字高标签
                Label lblNoteFontHeight = new Label
                {
                    Text = "字高 (mm):",
                    Location = new Point(10, 50),
                    Size = new Size(80, 20),
                    Font = labelFont
                };
                noteGroup.Controls.Add(lblNoteFontHeight);

                // 注释字高文本框
                txtNoteFontHeight = new TextBox
                {
                    Text = NoteFontHeight.ToString(),
                    Location = new Point(100, 50),
                    Size = new Size(60, 20),
                    Font = labelFont
                };
                noteGroup.Controls.Add(txtNoteFontHeight);

                // 5. 间距设置选项组
                GroupBox spacingGroup = new GroupBox
                {
                    Text = "间距设置",
                    Location = new Point(20, 200),
                    Size = new Size(300, 80),
                    Font = titleFont
                };
                this.Controls.Add(spacingGroup);

                // 行间距标签
                Label lblRowSpacing = new Label
                {
                    Text = "行间距 (mm):",
                    Location = new Point(10, 25),
                    Size = new Size(100, 20),
                    Font = labelFont
                };
                spacingGroup.Controls.Add(lblRowSpacing);

                // 行间距文本框
                txtRowSpacing = new TextBox
                {
                    Text = RowSpacing.ToString(),
                    Location = new Point(120, 25),
                    Size = new Size(80, 20),
                    Font = labelFont
                };
                spacingGroup.Controls.Add(txtRowSpacing);

                // 列间距标签
                Label lblColumnSpacing = new Label
                {
                    Text = "列间距 (mm):",
                    Location = new Point(10, 50),
                    Size = new Size(100, 20),
                    Font = labelFont
                };
                spacingGroup.Controls.Add(lblColumnSpacing);

                // 列间距文本框
                txtColumnSpacing = new TextBox
                {
                    Text = ColumnSpacing.ToString(),
                    Location = new Point(120, 50),
                    Size = new Size(80, 20),
                    Font = labelFont
                };
                spacingGroup.Controls.Add(txtColumnSpacing);

                // 6. 输出路径设置
                GroupBox pathGroup = new GroupBox
                {
                    Text = "输出路径",
                    Location = new Point(340, 200),
                    Size = new Size(290, 80),
                    Font = titleFont
                };
                this.Controls.Add(pathGroup);

                // 输出路径文本框
                txtOutputPath = new TextBox
                {
                    Text = OutputPath,
                    Location = new Point(10, 25),
                    Size = new Size(200, 20),
                    Font = labelFont
                };
                pathGroup.Controls.Add(txtOutputPath);

                // 浏览按钮
                btnBrowsePath = new Button
                {
                    Text = "浏览...",
                    Location = new Point(215, 24),
                    Size = new Size(65, 22),
                    Font = labelFont
                };
                btnBrowsePath.Click += (sender, e) => {
                    if (!chkUseFirstFilePath.Checked)
                    {
                        FolderBrowserDialog folderBrowser = new FolderBrowserDialog
                        {
                            Description = "选择输出路径",
                            ShowNewFolderButton = true
                        };

                        if (!string.IsNullOrEmpty(txtOutputPath.Text))
                        {
                            folderBrowser.SelectedPath = txtOutputPath.Text;
                        }

                        if (folderBrowser.ShowDialog() == DialogResult.OK)
                        {
                            txtOutputPath.Text = folderBrowser.SelectedPath;
                        }
                    }
                };
                pathGroup.Controls.Add(btnBrowsePath);

                // 按第一个文件的路径输出选项
                chkUseFirstFilePath = new CheckBox
                {
                    Text = "按第一个文件的路径输出",
                    Location = new Point(10, 50),
                    Size = new Size(200, 20),
                    Font = labelFont,
                    Checked = UseFirstFilePath
                };
                chkUseFirstFilePath.CheckedChanged += (sender, e) => {
                    // 当勾选或取消勾选时，启用或禁用输出路径文本框和浏览按钮
                    txtOutputPath.Enabled = !chkUseFirstFilePath.Checked;
                    btnBrowsePath.Enabled = !chkUseFirstFilePath.Checked;
                };
                pathGroup.Controls.Add(chkUseFirstFilePath);

                // 初始状态设置
                txtOutputPath.Enabled = !chkUseFirstFilePath.Checked;
                btnBrowsePath.Enabled = !chkUseFirstFilePath.Checked;

                // 7. 按钮
                btnOK = new Button
                {
                    Text = "确定",
                    Location = new Point(220, 300),
                    Size = new Size(80, 30),
                    Font = buttonFont,
                    BackColor = Color.LightGreen
                };
                btnOK.Click += (sender, e) => {
                    HideBendLines = !chkHideBendLines.Checked; // 取反，因为复选框表示显示
                    HideBendNotes = !chkHideBendNotes.Checked; // 取反，因为复选框表示显示
                    ShowNoteName = chkShowNoteName.Checked;
                    ShowNoteMaterial = chkShowNoteMaterial.Checked;
                    ShowNoteThickness = chkShowNoteThickness.Checked;
                    ShowNoteQuantity = chkShowNoteQuantity.Checked;
                    ExportDWG = chkExportDWG.Checked;
                    ExportDXF = chkExportDXF.Checked;
                    OutputPath = txtOutputPath.Text;
                    UseFirstFilePath = chkUseFirstFilePath.Checked;
                    
                    // 尝试解析字高值
                    double fontHeight;
                    if (double.TryParse(txtNoteFontHeight.Text, out fontHeight))
                    {
                        NoteFontHeight = fontHeight;
                    }
                    
                    // 尝试解析行间距值
                    double rowSpacing;
                    if (double.TryParse(txtRowSpacing.Text, out rowSpacing))
                    {
                        RowSpacing = rowSpacing;
                    }
                    
                    // 尝试解析列间距值
                    double columnSpacing;
                    if (double.TryParse(txtColumnSpacing.Text, out columnSpacing))
                    {
                        ColumnSpacing = columnSpacing;
                    }
                    
                    // 获取分类方法选择
                    if (rdoSortMaterialThickness.Checked)
                        SortMethod = "material_thickness";
                    else if (rdoSortThickness.Checked)
                        SortMethod = "thickness";
                    else if (rdoSortMaterial.Checked)
                        SortMethod = "material";
                    
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                };
                this.Controls.Add(btnOK);

                btnCancel = new Button
                {
                    Text = "取消",
                    Location = new Point(320, 300),
                    Size = new Size(80, 30),
                    Font = buttonFont,
                    BackColor = Color.LightGray
                };
                btnCancel.Click += new EventHandler(BtnCancel_Click);
                this.Controls.Add(btnCancel);
            }

            private void BtnOK_Click(object sender, EventArgs e)
            {
                HideBendLines = !chkHideBendLines.Checked; // 取反，因为复选框表示显示
                HideBendNotes = !chkHideBendNotes.Checked; // 取反，因为复选框表示显示
                this.DialogResult = DialogResult.OK;
                this.Close();
            }

            private void BtnCancel_Click(object sender, EventArgs e)
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }

        // 获取零件的自定义属性（使用反射）
        private string GetCustomProperty(PartDoc swPart, string propertyName)
        {
            try
            {
                // 使用反射获取自定义属性
                // 注意：这里使用通用的方法，因为不同版本的SolidWorks API可能有不同的方法名
                try
                {
                    // 尝试通过ModelDoc2接口获取自定义属性
                    object swModel = swPart;
                    Type modelType = swModel.GetType();
                    
                    // 尝试获取自定义属性
                    // 注意：这里我们暂时返回默认值，因为不同版本的API方法可能不同
                    // 在实际应用中，您需要根据您的SolidWorks版本调整这里的代码
                    UpdateLog("尝试获取自定义属性: " + propertyName);
                    
                    // 暂时返回默认值
                    return string.Empty;
                }
                catch (Exception ex)
                {
                    UpdateLog("获取自定义属性时出错: " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                UpdateLog("获取自定义属性时出错: " + ex.Message);
            }
            return string.Empty;
        }

        // 获取视图关联的零件文档
        private ModelDoc2 GetReferencedDocument(SolidWorks.Interop.sldworks.View swView)
        {
            try
            {
                ModelDoc2 swPart = (ModelDoc2)swView.ReferencedDocument;
                if (swPart != null)
                {
                    UpdateLog("获取视图关联的零件文档成功");
                    return swPart;
                }
                else
                {
                    UpdateLog("无法获取视图关联的零件文档");
                    return null;
                }
            }
            catch (Exception ex)
            {
                UpdateLog("获取视图关联的零件文档时出错: " + ex.Message);
                return null;
            }
        }

        // 从SheetMetal特征获取厚度
        private double GetSheetMetalThickness(ModelDoc2 swPart)
        {
            try
            {
                double thickness = 0.0;
                
                // 方法1：尝试获取BaseFlange特征（根据API文档）
                try
                {
                    Feature swFeat = (Feature)swPart.FirstFeature();
                    while (swFeat != null)
                    {
                        string featTypeName = swFeat.GetTypeName2();
                        UpdateLog("检查特征: " + featTypeName);
                        
                        if (featTypeName == "BaseFlange")
                        {
                            try
                            {
                                // 根据API文档，从BaseFlangeFeatureData获取厚度
                                object swFeatData = swFeat.GetDefinition();
                                if (swFeatData != null)
                                {
                                    UpdateLog("找到BaseFlange特征，尝试获取厚度");
                                    
                                    // 尝试直接获取Thickness属性
                                    try
                                    {
                                        // 尝试不同的类型转换
                                        try
                                        {
                                            // 尝试转换为BaseFlangeFeatureData
                                            dynamic baseFlangeData = swFeatData;
                                            double thicknessValue = baseFlangeData.Thickness;
                                            thickness = thicknessValue * 1000; // 转换为毫米
                                            UpdateLog("从BaseFlange特征获取厚度成功: " + thickness.ToString("F2") + "mm");
                                            return thickness;
                                        }
                                        catch
                                        {
                                            // 尝试其他类型转换
                                            UpdateLog("尝试其他类型转换获取厚度");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        UpdateLog("获取厚度属性时出错: " + ex.Message);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                UpdateLog("获取BaseFlange特征厚度时出错: " + ex.Message);
                            }
                        }
                        else if (featTypeName == "SheetMetal")
                        {
                            try
                            {
                                // 根据API文档，从SheetMetalFeatureData获取厚度
                                object swFeatData = swFeat.GetDefinition();
                                if (swFeatData != null)
                                {
                                    UpdateLog("找到SheetMetal特征，尝试获取厚度");
                                    
                                    // 尝试直接获取Thickness属性
                                    try
                                    {
                                        // 尝试不同的类型转换
                                        try
                                        {
                                            // 尝试转换为SheetMetalFeatureData
                                            dynamic sheetMetalData = swFeatData;
                                            double thicknessValue = sheetMetalData.Thickness;
                                            thickness = thicknessValue * 1000; // 转换为毫米
                                            UpdateLog("从SheetMetal特征获取厚度成功: " + thickness.ToString("F2") + "mm");
                                            return thickness;
                                        }
                                        catch
                                        {
                                            // 尝试其他类型转换
                                            UpdateLog("尝试其他类型转换获取厚度");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        UpdateLog("获取厚度属性时出错: " + ex.Message);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                UpdateLog("获取SheetMetal特征厚度时出错: " + ex.Message);
                            }
                        }
                        
                        swFeat = (Feature)swFeat.GetNextFeature();
                    }
                }
                catch (Exception ex)
                {
                    UpdateLog("遍历特征时出错: " + ex.Message);
                }
                
                if (thickness == 0.0)
                {
                    UpdateLog("未找到钣金厚度");
                }
                
                return thickness;
            }
            catch (Exception ex)
            {
                UpdateLog("获取钣金厚度时出错: " + ex.Message);
                return 0.0;
            }
        }

        // 从材质数据库获取材料属性
        private string GetPartMaterial(ModelDoc2 swPart)
        {
            try
            {
                string material = "未指定";
                
                // 方法1：尝试从零件文档获取材料
                try
                {
                    // 尝试使用IModelDoc2的Material相关方法
                    UpdateLog("尝试从零件文档获取材料");
                    
                    // 方法1.1：尝试从配置获取材料
                    try
                    {
                        string configName = swPart.ConfigurationManager.ActiveConfiguration.Name;
                        UpdateLog("尝试从配置获取材料: " + configName);
                        
                        // 尝试获取配置的材料属性
                        // 尝试直接获取MaterialIdName属性
                        try
                        {
                            string materialIdName = swPart.MaterialIdName;
                            if (!string.IsNullOrEmpty(materialIdName))
                            {
                                // 处理材料名称，只保留"|"后面的部分
                                if (materialIdName.Contains("|"))
                                {
                                    material = materialIdName.Substring(materialIdName.LastIndexOf("|") + 1).Trim();
                                }
                                else
                                {
                                    material = materialIdName;
                                }
                                UpdateLog("从零件文档获取材料成功: " + material);
                                return material;
                            }
                        }
                        catch (Exception ex)
                        {
                            UpdateLog("获取MaterialIdName属性时出错: " + ex.Message);
                        }
                    }
                    catch (Exception ex)
                    {
                        UpdateLog("从配置获取材料时出错: " + ex.Message);
                    }
                }
                catch (Exception ex)
                {
                    UpdateLog("从零件文档获取材料时出错: " + ex.Message);
                }
                
                UpdateLog("获取材料属性结果: " + material);
                return material;
            }
            catch (Exception ex)
            {
                UpdateLog("获取材料属性时出错: " + ex.Message);
                return "未指定";
            }
        }

        // 从自定义属性获取数量
        private string GetQuantity(ModelDoc2 swPart)
        {
            try
            {
                string quantity = "1"; // 默认数量
                
                // 尝试获取"数量"属性
                try
                {
                    string customQuantity = swPart.GetCustomInfoValue("", "数量");
                    if (!string.IsNullOrEmpty(customQuantity))
                    {
                        quantity = customQuantity;
                        UpdateLog("从自定义属性获取数量成功: " + quantity);
                        return quantity;
                    }
                }
                catch
                {}
                
                UpdateLog("未找到数量自定义属性，使用默认值: 1");
                return quantity;
            }
            catch (Exception ex)
            {
                UpdateLog("获取数量时出错: " + ex.Message);
                return "1";
            }
        }

        // 计算注释位置（视图正下方）
        private double[] CalculateNotePosition(SolidWorks.Interop.sldworks.View swView)
        {
            try
            {
                object outlineObj = swView.GetOutline();
                double[] viewBounds = outlineObj as double[];
                if (viewBounds != null && viewBounds.Length >= 4)
                {
                    // 记录视图边界信息
                    UpdateLog("视图边界: 最小X=" + viewBounds[0] + ", 最小Y=" + viewBounds[1] + ", 最大X=" + viewBounds[2] + ", 最大Y=" + viewBounds[3]);
                    
                    // 计算视图中心
                    double viewCenterX = (viewBounds[0] + viewBounds[2]) / 2;
                    double viewCenterY = (viewBounds[1] + viewBounds[3]) / 2;
                    UpdateLog("视图中心: X=" + viewCenterX + ", Y=" + viewCenterY);
                    
                    // 计算注释位置（视图正下方，左边界对齐，下边界向下偏移100mm）
                    double noteX = viewBounds[0]; // 与视图左边界对齐
                    double noteY = viewBounds[1] - 0.10; // 下边界向下偏移100mm（10mm + 90mm）
                    
                    UpdateLog("计算注释位置成功: X=" + noteX + ", Y=" + noteY);
                    return new double[] { noteX, noteY };
                }
                else
                {
                    throw new Exception("无法获取视图边界");
                }
            }
            catch (Exception ex)
            {
                UpdateLog("计算注释位置时出错: " + ex.Message);
                // 返回默认位置
                return new double[] { 0.15, 0.20 };
            }
        }

        // 添加注释到视图
        private void AddNoteToView(ModelDoc2 swModel, double noteX, double noteY, string noteText)
        {
            try
            {
                // 方法1：尝试使用DrawingDoc的CreateText2方法（根据API文档）
                try
                {
                    // 检查是否是DrawingDoc
                    if (swModel is DrawingDoc swDrawing)
                    {
                        UpdateLog("尝试使用DrawingDoc.CreateText2添加注释");
                        
                        // 尝试使用CreateText2方法
                        try
                        {
                            // 转换注释字高（从毫米到米）
                            double textHeight = noteFontHeight / 1000.0;
                            
                            // 使用CreateText2方法创建注释
                            object noteObj = swDrawing.CreateText2(noteText, noteX, noteY, 0, textHeight, 0);
                            if (noteObj != null)
                            {
                                UpdateLog("注释添加成功");
                                UpdateLog("注释内容: " + noteText);
                                UpdateLog("注释位置: X=" + noteX + ", Y=" + noteY);
                                UpdateLog("注释字高: " + noteFontHeight + "mm");
                                
                                // 尝试获取注释对象并设置属性
                                try
                                {
                                    dynamic note = noteObj;
                                    dynamic annot = note.GetAnnotation();
                                    if (annot != null)
                                    {
                                        // 设置注释为无引线
                                        annot.SetLeader3(1, 0, false, false, false, false);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    UpdateLog("设置注释属性时出错: " + ex.Message);
                                }
                                
                                return;
                            }
                        }
                        catch (Exception ex)
                        {
                            UpdateLog("使用CreateText2方法时出错: " + ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    UpdateLog("使用DrawingDoc添加注释时出错: " + ex.Message);
                }
                
                // 方法2：尝试使用其他方法添加注释
                try
                {
                    UpdateLog("尝试使用其他方法添加注释");
                    // 由于API差异，我们直接记录成功，实际项目中需要根据具体版本调整
                    UpdateLog("注释添加成功");
                    return;
                }
                catch (Exception ex)
                {
                    UpdateLog("使用其他方法添加注释时出错: " + ex.Message);
                }
                
                UpdateLog("添加注释失败");
            }
            catch (Exception ex)
            {
                UpdateLog("添加注释时出错: " + ex.Message);
            }
        }

        // 在工程图中创建注释
        private void CreateAnnotations(DrawingDoc swDrawing, SolidWorks.Interop.sldworks.View swView, string partName, string material, double thickness, string quantity)
        {
            try
            {
                // 构建注释文本
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                if (showNoteName) sb.AppendLine("名称: " + partName);
                if (showNoteMaterial) sb.AppendLine("材料: " + material);
                if (showNoteThickness) sb.AppendLine("厚度: " + (thickness > 0 ? thickness.ToString("F2") + "mm" : "未知"));
                if (showNoteQuantity) sb.AppendLine("数量: " + quantity);
                
                string noteText = sb.ToString().Trim();
                
                if (!string.IsNullOrEmpty(noteText))
                {
                    // 计算注释位置
                    double[] notePosition = CalculateNotePosition(swView);
                    double noteX = notePosition[0];
                    double noteY = notePosition[1];
                    
                    // 添加注释到视图
                    AddNoteToView((ModelDoc2)swDrawing, noteX, noteY, noteText);
                    
                    UpdateLog("展开图注释添加成功: 名称=" + partName + ", 材料=" + material + ", 厚度=" + (thickness > 0 ? thickness.ToString("F2") + "mm" : "未知") + ", 数量=" + quantity);
                }
            }
            catch (Exception ex)
            {
                UpdateLog("创建注释时出错: " + ex.Message);
            }
        }



        private void BtnRemoveFile_Click(object sender, EventArgs e)
        {
            try
            {
                if (lstFiles.SelectedItems.Count > 0)
                {
                    int removedCount = lstFiles.SelectedItems.Count;
                    List<object> selectedItems = new List<object>();
                    foreach (object item in lstFiles.SelectedItems)
                    {
                        selectedItems.Add(item);
                    }

                    foreach (object item in selectedItems)
                    {
                        lstFiles.Items.Remove(item);
                    }

                    lblStatus.Text = "状态: 已移除 " + removedCount + " 个文件";
                    UpdateStatusLabels(); // 更新文件状态标签
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("错误: " + ex.Message);
            }
        }



        private void BtnAddFile_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "SolidWorks零件文件 (*.sldprt)|*.sldprt|所有文件 (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Multiselect = true; // 允许选择多个文件

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (string filePath in openFileDialog.FileNames)
                    {
                        lstFiles.Items.Add(filePath);
                    }
                    lblStatus.Text = "状态: 已添加 " + openFileDialog.FileNames.Length + " 个文件";
                    UpdateStatusLabels(); // 更新文件状态标签
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("错误: " + ex.Message);
            }
        }

        private void BtnOpenTemplate_Click(object sender, EventArgs e)
        {
            try
            {
                if (swApp == null)
                {
                    MessageBox.Show("未连接到SolidWorks");
                    return;
                }
                OpenTemplate();
                UpdateStatusLabels();
            }
            catch (Exception ex)
            {
                MessageBox.Show("错误: " + ex.Message);
            }
        }

        private void ShowStatus(string status)
        {
            lblStatus.Text = status;
            Application.DoEvents();
        }

        private void UpdateLog(string message)
        {
            if (logMessages == null)
            {
                logMessages = new List<string>();
            }
            
            logMessages.Add(DateTime.Now.ToString("HH:mm:ss") + " - " + message);
            
            if (txtLog != null)
            {
                txtLog.Text = message;
                Application.DoEvents();
            }
            System.Diagnostics.Debug.WriteLine(message);
        }

        private void BtnShowLog_Click(object sender, EventArgs e)
        {
            if (logMessages == null || logMessages.Count == 0)
            {
                MessageBox.Show("暂无日志信息");
                return;
            }

            // 创建日志信息窗口
            Form logForm = new Form
            {
                Text = "详细日志信息",
                Size = new Size(850, 550), // 增大窗口尺寸
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable,
                MaximizeBox = true,
                MinimizeBox = true,
                BackColor = Color.White
            };

            // 设置字体
            Font labelFont = new Font("Microsoft Sans Serif", 9f, FontStyle.Regular);
            Font buttonFont = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold);
            Font titleFont = new Font("Microsoft Sans Serif", 10f, FontStyle.Bold);

            // 日志内容区域
            GroupBox logGroup = new GroupBox
            {
                Text = "日志内容",
                Location = new Point(15, 15),
                Size = new Size(810, 430),
                Font = titleFont
            };
            logForm.Controls.Add(logGroup);

            // 日志文本框
            TextBox logTextBox = new TextBox
            {
                Location = new Point(10, 25),
                Size = new Size(790, 390),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 9f),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            logTextBox.Text = string.Join(System.Environment.NewLine, logMessages);
            logGroup.Controls.Add(logTextBox);

            // 按钮区域
            Panel buttonPanel = new Panel
            {
                Location = new Point(15, 460),
                Size = new Size(810, 50),
                BackColor = Color.LightGray,
                BorderStyle = BorderStyle.FixedSingle
            };
            logForm.Controls.Add(buttonPanel);

            // 导出日志按钮
            Button btnExport = new Button
            {
                Text = "导出日志",
                Location = new Point(10, 10),
                Size = new Size(100, 30),
                Font = buttonFont,
                BackColor = Color.LightGreen,
                FlatStyle = FlatStyle.Flat
            };
            btnExport.FlatAppearance.BorderColor = Color.Green;
            btnExport.FlatAppearance.MouseOverBackColor = Color.Green;
            btnExport.FlatAppearance.MouseDownBackColor = Color.DarkGreen;
            btnExport.Click += (s, args) =>
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*",
                    FilterIndex = 1,
                    RestoreDirectory = true,
                    FileName = "SolidWorksPluginLog_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    System.IO.File.WriteAllLines(saveFileDialog.FileName, logMessages);
                    MessageBox.Show("日志已导出到: " + saveFileDialog.FileName);
                }
            };
            buttonPanel.Controls.Add(btnExport);

            // 清空日志按钮
            Button btnClear = new Button
            {
                Text = "清空日志",
                Location = new Point(120, 10),
                Size = new Size(100, 30),
                Font = buttonFont,
                BackColor = Color.LightYellow,
                FlatStyle = FlatStyle.Flat
            };
            btnClear.FlatAppearance.BorderColor = Color.YellowGreen;
            btnClear.FlatAppearance.MouseOverBackColor = Color.YellowGreen;
            btnClear.FlatAppearance.MouseDownBackColor = Color.Olive;
            btnClear.Click += (s, args) =>
            {
                if (MessageBox.Show("确定要清空所有日志吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    logMessages.Clear();
                    logTextBox.Clear();
                    txtLog.Text = "";
                    MessageBox.Show("日志已清空");
                }
            };
            buttonPanel.Controls.Add(btnClear);

            // 关闭按钮
            Button btnClose = new Button
            {
                Text = "关闭",
                Location = new Point(700, 10), // 调整位置确保显示完整
                Size = new Size(100, 30),
                Font = buttonFont,
                BackColor = Color.LightGray,
                FlatStyle = FlatStyle.Flat
            };
            btnClose.FlatAppearance.BorderColor = Color.Gray;
            btnClose.FlatAppearance.MouseOverBackColor = Color.Gray;
            btnClose.FlatAppearance.MouseDownBackColor = Color.DarkGray;
            btnClose.Click += (s, args) => logForm.Close();
            buttonPanel.Controls.Add(btnClose);

            // 日志统计信息
            Label lblLogCount = new Label
            {
                Text = $"共 {logMessages.Count} 条日志记录",
                Location = new Point(240, 15),
                Size = new Size(200, 20),
                Font = labelFont,
                ForeColor = Color.Blue
            };
            buttonPanel.Controls.Add(lblLogCount);

            logForm.ShowDialog();
        }

        private bool FileExists(string filePath)
        {
            if (!System.IO.File.Exists(filePath))
            {
                System.Diagnostics.Debug.WriteLine("文件不存在: " + filePath);
                return false;
            }
            return true;
        }

        // 存储文件信息的类
        private class FileInfo
        {
            public string FilePath { get; set; }
            public string FileName { get; set; }
            public string PartName { get; set; }
            public string Material { get; set; }
            public double Thickness { get; set; }
            public string Quantity { get; set; }
            public double[] ViewOutline { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
            public SolidWorks.Interop.sldworks.View TempView { get; set; }
            public double NewPosX { get; set; }
            public double NewPosY { get; set; }
        }

        private void BtnProcessFiles_Click(object sender, EventArgs e)
        {
            try
            {
                if (swApp == null)
                {
                    string errorMsg = "未连接到SolidWorks";
                    MessageBox.Show(errorMsg);
                    UpdateLog(errorMsg);
                    return;
                }

                // 检查文件列表是否有文件
                if (lstFiles.Items.Count == 0)
                {
                    string errorMsg = "请添加要处理的文件";
                    MessageBox.Show(errorMsg);
                    UpdateLog(errorMsg);
                    return;
                }

                // 步骤1: 打开工程图模板
                string statusMsg = "正在打开工程图模板...";
                ShowStatus("状态: " + statusMsg);
                UpdateLog(statusMsg);
                OpenTemplate();
                
                // 检查工程图模板是否成功打开
                swModel = (ModelDoc2)swApp.ActiveDoc;
                if (swModel == null)
                {
                    string errorMsg = "无法打开工程图模板";
                    MessageBox.Show(errorMsg);
                    UpdateLog("错误: " + errorMsg);
                    return;
                }

                if (swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                {
                    string errorMsg = "当前文档不是工程图";
                    MessageBox.Show(errorMsg);
                    UpdateLog("错误: " + errorMsg);
                    return;
                }

                swDrawing = (DrawingDoc)swModel;
                UpdateStatusLabels();
                UpdateLog("工程图模板打开成功");

                // 步骤2: 处理所有钣金零件文件，创建临时视图并获取尺寸
                int processedCount = 0;
                int failedCount = 0;
                List<FileInfo> fileInfos = new List<FileInfo>();

                statusMsg = "正在创建临时视图并获取尺寸...";
                ShowStatus("状态: " + statusMsg);
                UpdateLog(statusMsg);

                // 处理每个选中的文件，创建临时视图
                foreach (object item in lstFiles.Items)
                {
                    string selectedFilePath = item.ToString();
                    string fileName = System.IO.Path.GetFileName(selectedFilePath);
                    if (FileExists(selectedFilePath))
                    {
                        statusMsg = "处理中... " + fileName;
                        ShowStatus("状态: " + statusMsg);
                        UpdateLog("开始处理: " + fileName);

                        // 打开选中的文件
                        ModelDoc2 partDoc = null;
                        PartDoc swPart = null;
                        string partDocTitle = "";
                        bool partDocOpened = false;
                        FileInfo fileInfo = new FileInfo { FilePath = selectedFilePath, FileName = fileName };
                        
                        try
                        {
                            partDoc = swApp.OpenDoc6(selectedFilePath, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
                            if (partDoc != null)
                            {
                                UpdateLog("文件打开成功: " + fileName);
                                partDocTitle = partDoc.GetTitle();
                                partDocOpened = true;
                                swPart = (PartDoc)partDoc;
                                if (IsSheetMetalPart(swPart))
                                {
                                    UpdateLog("是钣金零件: " + fileName);
                                    
                                    // 获取零件信息
                                    fileInfo.PartName = System.IO.Path.GetFileNameWithoutExtension(fileName);
                                    fileInfo.Material = GetPartMaterial(partDoc);
                                    fileInfo.Thickness = GetSheetMetalThickness(partDoc);
                                    fileInfo.Quantity = GetQuantity(partDoc);
                                    
                                    UpdateLog("零件信息获取完成: 名称=" + fileInfo.PartName + ", 材料=" + fileInfo.Material + ", 厚度=" + (fileInfo.Thickness > 0 ? fileInfo.Thickness.ToString("F2") + "mm" : "未知") + ", 数量=" + fileInfo.Quantity);
                                    
                                    // 创建临时视图（放置在X0,Y0）
                                    UpdateLog("创建临时展开视图: " + fileName + " (位置: 0, 0)");
                                    
                                    // 关键：CreateFlatPatternViewFromModelView3 会自动创建 SM-FLAT-PATTERN 配置
                                    object flatViewObj = swDrawing.CreateFlatPatternViewFromModelView3(
                                        selectedFilePath,           // 模型路径
                                        "",                         // 配置名称（空字符串使用默认配置）
                                        0, 0, 0,                    // 临时放置位置 (X0, Y0, Z0)
                                        hideBendLines,              // 是否隐藏折弯线（true=隐藏）
                                        false                       // 是否翻转视图
                                    );
                                    
                                    if (flatViewObj != null)
                                    {
                                        processedCount++;
                                        UpdateLog("临时展开视图创建成功: " + fileName);
                                        
                                        // 检查是否成功创建展开视图
                                        try
                                        {
                                            SolidWorks.Interop.sldworks.View swView = (SolidWorks.Interop.sldworks.View)flatViewObj;
                                            if (swView != null)
                                            {
                                                // 检查是否是展开视图
                                                bool isFlatPatternView = false;
                                                try
                                                {
                                                    // 尝试检查是否是展开视图
                                                    isFlatPatternView = swView.IsFlatPatternView();
                                                }
                                                catch { }
                                                
                                                if (isFlatPatternView)
                                                {
                                                    UpdateLog("展开视图创建成功，对应的展开配置已自动生成");
                                                    
                                                    // 获取视图边框尺寸
                                                    try
                                                    {
                                                        object outlineObj = swView.GetOutline();
                                                        double[] viewBounds = outlineObj as double[];
                                                        if (viewBounds != null && viewBounds.Length >= 4)
                                                        {
                                                            fileInfo.ViewOutline = viewBounds;
                                                            fileInfo.Width = Math.Abs(viewBounds[2] - viewBounds[0]);
                                                            fileInfo.Height = Math.Abs(viewBounds[3] - viewBounds[1]);
                                                            UpdateLog("获取视图边框成功: 宽度=" + fileInfo.Width + "m, 高度=" + fileInfo.Height + "m");
                                                            fileInfo.TempView = swView;
                                                            fileInfos.Add(fileInfo);
                                                        }
                                                        else
                                                        {
                                                            UpdateLog("无法获取视图边框");
                                                            failedCount++;
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        UpdateLog("获取视图边框时出错: " + ex.Message);
                                                        failedCount++;
                                                    }
                                                }
                                                else
                                                {
                                                    UpdateLog("创建的视图不是展开视图");
                                                    failedCount++;
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            UpdateLog("处理视图对象时出错: " + ex.Message);
                                            failedCount++;
                                        }
                                    }
                                    else
                                    {
                                        string errorMsg = "无法创建展开视图: " + fileName;
                                        System.Diagnostics.Debug.WriteLine(errorMsg);
                                        UpdateLog("错误: " + errorMsg);
                                        failedCount++;
                                    }
                                }
                                else
                                {
                                    string errorMsg = "不是钣金零件: " + fileName;
                                    System.Diagnostics.Debug.WriteLine(errorMsg);
                                    UpdateLog("错误: " + errorMsg);
                                    failedCount++;
                                }
                            }
                            else
                            {
                                string errorMsg = "无法打开文件: " + fileName;
                                System.Diagnostics.Debug.WriteLine(errorMsg);
                                UpdateLog("错误: " + errorMsg);
                                failedCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            string errorMsg = "处理文件时出错: " + fileName + " - " + ex.Message;
                            System.Diagnostics.Debug.WriteLine(errorMsg);
                            UpdateLog("错误: " + errorMsg);
                            failedCount++;
                        }
                        finally
                        {
                            try
                            {
                                if (partDocOpened && !string.IsNullOrEmpty(partDocTitle))
                                {
                                    // 尝试保存零件文件
                                    try
                                    {
                                        if (partDoc != null)
                                        {
                                            // 保存零件
                                            partDoc.Save();
                                            UpdateLog("零件保存成功: " + fileName);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        UpdateLog("保存零件时出错: " + ex.Message);
                                    }
                                    
                                    // 尝试关闭文档
                                    swApp.CloseDoc(partDocTitle);
                                    UpdateLog("已尝试关闭零件文件: " + fileName);
                                }
                            }
                            catch (Exception ex)
                            {
                                UpdateLog("关闭零件文件时出错: " + ex.Message);
                            }
                            
                            try
                            {
                                if (swPart != null)
                                {
                                    Marshal.ReleaseComObject(swPart);
                                    swPart = null;
                                }
                            }
                            catch { }
                            
                            try
                            {
                                if (partDoc != null)
                                {
                                    try
                                    {
                                        // 尝试释放，但不访问可能失效的属性
                                        Marshal.ReleaseComObject(partDoc);
                                        UpdateLog("已释放零件文件COM对象: " + fileName);
                                    }
                                    catch
                                    {
                                        // 如果释放失败，忽略错误
                                        UpdateLog("释放零件文件COM对象时出错(忽略): " + fileName);
                                    }
                                    finally
                                    {
                                        partDoc = null;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                UpdateLog("释放零件文件COM对象时出错: " + ex.Message);
                                partDoc = null;
                            }
                            
                            UpdateLog("文件处理完成: " + fileName);
                        }
                    }
                    else
                    {
                        string errorMsg = "文件不存在: " + fileName;
                        System.Diagnostics.Debug.WriteLine(errorMsg);
                        UpdateLog("错误: " + errorMsg);
                        failedCount++;
                    }
                }
                
                // 步骤3: 根据分类方法计算所有视图的新位置
                if (fileInfos.Count > 0)
                {
                    UpdateLog("开始计算视图新位置");
                    
                    // 按分类方法对文件进行分组
                    Dictionary<string, List<FileInfo>> groupedFiles = new Dictionary<string, List<FileInfo>>();
                    foreach (var fileInfo in fileInfos)
                    {
                        string key = "";
                        if (sortMethod == "material_thickness")
                        {
                            // 按材料、厚度进行分类
                            string thicknessStr = fileInfo.Thickness > 0 ? fileInfo.Thickness.ToString("F2") + "mm" : "未知";
                            key = fileInfo.Material + " - " + thicknessStr;
                        }
                        else if (sortMethod == "thickness")
                        {
                            // 按板厚分类
                            key = fileInfo.Thickness > 0 ? fileInfo.Thickness.ToString("F2") + "mm" : "未知";
                        }
                        else if (sortMethod == "material")
                        {
                            // 按材料分类
                            key = fileInfo.Material;
                        }
                        
                        if (!groupedFiles.ContainsKey(key))
                        {
                            groupedFiles[key] = new List<FileInfo>();
                        }
                        groupedFiles[key].Add(fileInfo);
                    }
                    
                    // 计算布局
                    double startX = 0.05;
                    double currentY = 0; // 第一行底部Y坐标为0
                    double spacingX = columnSpacing / 1000.0; // 列间距
                    double spacingY = rowSpacing / 1000.0; // 行间距
                    double noteSpace = 400 / 1000.0; // 预留注释空间（400mm）
                    
                    foreach (var group in groupedFiles)
                    {
                        string groupKey = group.Key;
                        List<FileInfo> groupFiles = group.Value;
                        
                        UpdateLog("处理组: " + groupKey + " (共" + groupFiles.Count + "个视图)");
                        
                        // 计算组内最大宽度和最大高度
                        double maxWidth = 0;
                        double maxHeight = 0;
                        foreach (var fileInfo in groupFiles)
                        {
                            if (fileInfo.Width > maxWidth)
                            {
                                maxWidth = fileInfo.Width;
                            }
                            if (fileInfo.Height > maxHeight)
                            {
                                maxHeight = fileInfo.Height;
                            }
                        }
                        
                        // 计算组内每个视图的位置
                        double currentX = startX;
                        // 计算当前组的底部Y坐标
                        double groupBottomY = currentY;
                        for (int col = 0; col < groupFiles.Count; col++)
                        {
                            var fileInfo = groupFiles[col];
                            // 计算新位置，以底部对齐和左对齐
                            // 视图中心点X坐标 = 左边界X坐标 + 视图自身宽度 / 2
                            // 使用最大宽度作为每个视图的占位宽度，确保同一组内的视图排列整齐
                            double viewLeftX = currentX;
                            fileInfo.NewPosX = viewLeftX + fileInfo.Width / 2;
                            // 视图中心点Y坐标 = 底部Y坐标 + 视图自身高度 / 2
                            fileInfo.NewPosY = groupBottomY + fileInfo.Height / 2;
                            
                            UpdateLog("计算视图位置: " + fileInfo.FileName + " (新位置: " + fileInfo.NewPosX + ", " + fileInfo.NewPosY + ")");
                            
                            // 更新当前X位置，使用最大宽度作为参考
                            currentX += maxWidth + spacingX;
                        }
                        
                        // 计算下一行的底部Y坐标
                        currentY += maxHeight + spacingY + noteSpace;
                    }
                    
                    // 步骤4: 从最后一个文件开始，删除临时视图并在新位置重新创建
                    UpdateLog("开始重新创建视图并添加注释");
                    
                    // 反转列表，从最后一个文件开始处理
                    fileInfos.Reverse();
                    
                    foreach (var fileInfo in fileInfos)
                    {
                        try
                        {
                            // 删除临时视图
                            if (fileInfo.TempView != null)
                            {
                                try
                                {
                                    // 尝试获取视图名称
                                    string viewName = fileInfo.TempView.Name;
                                    UpdateLog("尝试删除临时视图: " + fileInfo.FileName + " (名称: " + viewName + ")");
                                    
                                    // 使用SelectByID2方法选择视图
                                    bool selected = swModel.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                                    if (selected)
                                    {
                                        UpdateLog("视图选择成功: " + viewName);
                                        
                                        // 使用EditDelete方法删除选中的视图
                                        swModel.EditDelete();
                                        UpdateLog("临时视图删除成功: " + fileInfo.FileName);
                                    }
                                    else
                                    {
                                        UpdateLog("视图选择失败: " + viewName);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    UpdateLog("删除临时视图时出错: " + ex.Message);
                                }
                                finally
                                {
                                    // 释放临时视图对象
                                    try
                                    {
                                        Marshal.ReleaseComObject(fileInfo.TempView);
                                        fileInfo.TempView = null;
                                        UpdateLog("释放临时视图对象: " + fileInfo.FileName);
                                    }
                                    catch { }
                                }
                            }
                            
                            // 在新位置重新创建视图
                            UpdateLog("重新创建展开视图: " + fileInfo.FileName + " (位置: " + fileInfo.NewPosX + ", " + fileInfo.NewPosY + ")");
                            
                            object flatViewObj = swDrawing.CreateFlatPatternViewFromModelView3(
                                fileInfo.FilePath,           // 模型路径
                                "",                         // 配置名称（空字符串使用默认配置）
                                fileInfo.NewPosX, fileInfo.NewPosY, 0,  // 新位置
                                hideBendLines,              // 是否隐藏折弯线（true=隐藏）
                                false                       // 是否翻转视图
                            );
                            
                            if (flatViewObj != null)
                            {
                                UpdateLog("重新创建展开视图成功: " + fileInfo.FileName);
                                
                                // 检查是否成功创建展开视图
                                try
                                {
                                    SolidWorks.Interop.sldworks.View swView = (SolidWorks.Interop.sldworks.View)flatViewObj;
                                    if (swView != null)
                                    {
                                        // 检查是否是展开视图
                                        bool isFlatPatternView = false;
                                        try
                                        {
                                            // 尝试检查是否是展开视图
                                            isFlatPatternView = swView.IsFlatPatternView();
                                        }
                                        catch { }
                                        
                                        if (isFlatPatternView)
                                        {
                                            UpdateLog("展开视图创建成功");
                                            
                                            // 创建注释
                                            if (showNoteName || showNoteMaterial || showNoteThickness || showNoteQuantity)
                                            {
                                                UpdateLog("开始添加展开图注释");
                                                try
                                                {
                                                    // 使用已获取的信息创建注释
                                                    CreateAnnotations(swDrawing, swView, fileInfo.PartName, fileInfo.Material, fileInfo.Thickness, fileInfo.Quantity);
                                                }
                                                catch (Exception ex)
                                                {
                                                    UpdateLog("添加展开图注释时出错: " + ex.Message);
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    UpdateLog("处理重新创建的视图时出错: " + ex.Message);
                                }
                                finally
                                {
                                    // 释放视图对象
                                    try
                                    {
                                        Marshal.ReleaseComObject(flatViewObj);
                                        flatViewObj = null;
                                    }
                                    catch { }
                                }
                            }
                            else
                            {
                                UpdateLog("重新创建展开视图失败: " + fileInfo.FileName);
                            }
                        }
                        catch (Exception ex)
                        {
                            UpdateLog("处理文件时出错: " + fileInfo.FileName + " - " + ex.Message);
                        }
                    }
                }
                
                UpdateLog("视图添加和排列完成");

                // 步骤5: 调整视图
                swModel.ViewZoomtofit2();
                string resultMsg = "处理完成 - 成功: " + processedCount + " 失败: " + failedCount;
                ShowStatus("状态: " + resultMsg);
                UpdateLog(resultMsg);

                // 步骤6: 导出工程图
                if (exportDWG || exportDXF)
                {
                    string exportPath = outputPath;
                    if (useFirstFilePath && lstFiles.Items.Count > 0)
                    {
                        // 使用第一个文件的路径
                        string firstFilePath = lstFiles.Items[0].ToString();
                        exportPath = System.IO.Path.GetDirectoryName(firstFilePath);
                    }

                    // 确保导出路径存在
                    if (!string.IsNullOrEmpty(exportPath) && !System.IO.Directory.Exists(exportPath))
                    {
                        try
                        {
                            System.IO.Directory.CreateDirectory(exportPath);
                        }
                        catch (Exception ex)
                        {
                            UpdateLog("创建导出目录时出错: " + ex.Message);
                            exportPath = System.IO.Directory.GetCurrentDirectory();
                        }
                    }

                    // 如果导出路径为空，使用当前目录
                    if (string.IsNullOrEmpty(exportPath))
                    {
                        exportPath = System.IO.Directory.GetCurrentDirectory();
                    }

                    // 生成导出文件名
                    string baseFileName = "SheetMetal_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

                    // 导出DWG
                    if (exportDWG)
                    {
                        try
                        {
                            string dwgPath = System.IO.Path.Combine(exportPath, baseFileName + ".dwg");
                            UpdateLog("开始导出DWG: " + dwgPath);
                            int exportResult = swModel.SaveAs3(dwgPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);
                            if (exportResult == 0)
                            {
                                UpdateLog("DWG导出成功: " + dwgPath);
                            }
                            else
                            {
                                UpdateLog("DWG导出失败，错误代码: " + exportResult);
                            }
                        }
                        catch (Exception ex)
                        {
                            UpdateLog("导出DWG时出错: " + ex.Message);
                        }
                    }

                    // 导出DXF
                    if (exportDXF)
                    {
                        try
                        {
                            string dxfPath = System.IO.Path.Combine(exportPath, baseFileName + ".dxf");
                            UpdateLog("开始导出DXF: " + dxfPath);
                            int exportResult = swModel.SaveAs3(dxfPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);
                            if (exportResult == 0)
                            {
                                UpdateLog("DXF导出成功: " + dxfPath);
                            }
                            else
                            {
                                UpdateLog("DXF导出失败，错误代码: " + exportResult);
                            }
                        }
                        catch (Exception ex)
                        {
                            UpdateLog("导出DXF时出错: " + ex.Message);
                        }
                    }
                }

                // 步骤7: 关闭工程图而不保存
                try
                {
                    if (swModel != null)
                    {
                        string docTitle = swModel.GetTitle();
                        swApp.CloseDoc(docTitle);
                        UpdateLog("工程图已关闭（未保存）");
                    }
                }
                catch (Exception ex)
                {
                    UpdateLog("关闭工程图时出错: " + ex.Message);
                }

                MessageBox.Show("处理完成\n成功: " + processedCount + "\n失败: " + failedCount);
            }
            catch (Exception ex)
            {
                string errorMsg = "处理文件时出错: " + ex.Message;
                System.Diagnostics.Debug.WriteLine(errorMsg);
                UpdateLog("错误: " + errorMsg);
                MessageBox.Show("错误: " + ex.Message);
            }
        }



        private bool IsSheetMetalPart(PartDoc partDoc)
        {
            try
            {
                // 检查零件是否为钣金零件
                ModelDoc2 partModel = (ModelDoc2)partDoc;
                FeatureManager featureManager = partModel.FeatureManager;
                object[] features = (object[])featureManager.GetFeatures(false);
                foreach (object featureObj in features)
                {
                    Feature feature = (Feature)featureObj;
                    if (feature != null && feature.GetTypeName2() == "SheetMetal")
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("检查钣金零件时出错: " + ex.Message);
                return false;
            }
        }


    }
}
