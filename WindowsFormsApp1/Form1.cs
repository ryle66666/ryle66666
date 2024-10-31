using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        // 选择文件夹按钮点击事件
        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtFolderPath.Text = dialog.SelectedPath;
                }
            }
        }

        // 保存文件夹按钮点击事件
        private void btnSelectSavePath_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtSavePath.Text = dialog.SelectedPath;
                }
            }
        }

        // 处理数据按钮点击事件
        private void btnProcess_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFolderPath.Text))
            {
                MessageBox.Show("请选择文件夹!");
                return;
            }

            if (string.IsNullOrEmpty(txtSavePath.Text))
            {
                MessageBox.Show("请选择保存位置!");
                return;
            }

            if (!int.TryParse(txtLimit.Text, out int limit) || limit <= 0)
            {
                MessageBox.Show("请输入有效的数量!");
                return;
            }

            try
            {
                // 获取所有子文件夹并按时间排序（最新的在前）
                var subFolders = Directory.GetDirectories(txtFolderPath.Text)
                    .OrderByDescending(f => new DirectoryInfo(f).CreationTime)
                    .ToArray();

                if (subFolders.Length == 0)
                {
                    MessageBox.Show("选择的文件夹中没有找到子文件夹!");
                    return;
                }

                // 创建Excel
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("检测结果");

                    // 设置表头
                    worksheet.Cells[1, 1].Value = "通讯号";
                    worksheet.Cells[1, 2].Value = "设备编号";
                    worksheet.Cells[1, 3].Value = "时间";
                    worksheet.Cells[1, 4].Value = "结果";

                    // 设置测量数据表头
                    worksheet.Cells[1, 5].Value = "垂直度";
                    worksheet.Cells[1, 6].Value = "孔径长边";
                    worksheet.Cells[1, 7].Value = "孔径短边";
                    worksheet.Cells[1, 8].Value = "位置度";
                    // 错位值1-4
                    worksheet.Cells[1, 9].Value = "错位值1-测点2";
                    worksheet.Cells[1, 10].Value = "错位值2-测点5";
                    worksheet.Cells[1, 11].Value = "错位值3-测点6";
                    worksheet.Cells[1, 12].Value = "错位值4-测点7";
                    // 缺料值1-8
                    worksheet.Cells[1, 13].Value = "缺料值1-测点2";
                    worksheet.Cells[1, 14].Value = "缺料值2-测点3";
                    worksheet.Cells[1, 15].Value = "缺料值3-测点4";
                    worksheet.Cells[1, 16].Value = "缺料值4-测点5";
                    worksheet.Cells[1, 17].Value = "缺料值5-测点6";
                    worksheet.Cells[1, 18].Value = "缺料值6-测点7";
                    worksheet.Cells[1, 19].Value = "缺料值7-测点8";
                    worksheet.Cells[1, 20].Value = "缺料值8-测点9";


                    // 设置图片相关的表头
                    // 从第5列开始，每两列为一组（图片+结果）
                    for (int i = 0; i < 20; i++) // 假设最多20张图片
                    {
                        worksheet.Cells[1, 21 + i * 2].Value = $"图片{i + 1}";
                        worksheet.Cells[1, 22 + i * 2].Value = $"图片{i + 1}结果";
                    }

                    // 设置表头样式
                    using (var range = worksheet.Cells[1, 1, 1, 45]) // 根据实际列数调整
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    }

                    int row = 2;
                    int processedCount = 0; // 记录已处理的文件夹数量
                    foreach (string folder in subFolders)
                    {
                        if (processedCount >= limit)
                        {
                            break; // 达到限制数量，停止处理
                        }
                        string folderName = Path.GetFileName(folder);
                        string[] parts = folderName.Split(',');
                        if (parts.Length >= 4)
                        {
                            // 写入基本信息
                            worksheet.Cells[row, 1].Value = parts[0];  // 序号
                            worksheet.Cells[row, 2].Value = parts[1];  // 设备编号
                            worksheet.Cells[row, 3].Value = parts[2];  // 时间
                            worksheet.Cells[row, 4].Value = parts[3];  // 结果

                            if (parts[3].ToUpper() == "OK")//OK为绿色
                            {
                                worksheet.Cells[row, 4].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                            }
                            else if (parts[3].ToUpper() == "NG")//NG为红色
                            {
                                worksheet.Cells[row, 4].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                            }

                            // 处理各编号结果文件夹
                            string measurementPath = Path.Combine(folder, "各编号结果");
                            if (Directory.Exists(measurementPath))
                            {
                                ProcessMeasurements(worksheet, row, measurementPath);
                            }


                            // 处理图片
                            string resultImagePath = Path.Combine(folder, "各区域结果图");
                            if (Directory.Exists(resultImagePath))
                            {
                                // 获取OK/NG文件夹中的图片
                                ProcessImages(worksheet, row, resultImagePath);
                            }

                            row++;
                            processedCount++;
                        }
                    }

                    // 保存Excel
                    string excelPath = Path.Combine(
                        txtSavePath.Text,  // 使用选择的保存路径
                        $"检测结果_{DateTime.Now:yyyyMMddHHmmss}.xlsx"
                    );
                    package.SaveAs(new FileInfo(excelPath));

                    MessageBox.Show($"处理完成!\n文件保存在: {excelPath}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"处理失败: {ex.Message}");
            }
        }

        private void ProcessMeasurements(ExcelWorksheet worksheet, int row, string measurementPath)
        {
            var files = Directory.GetFiles(measurementPath, "*.txt");
            foreach (var file in files)
            {
                string fileName = Path.GetFileNameWithoutExtension(file);
                string[] parts = fileName.Split(',');
                if (parts.Length >= 5)
                {
                    string measurementType = parts[1];
                    
                    string result = parts[4];
                    string result2 = parts[3];

                    if (measurementType.ToLower() == "孔径")
                    {
                        // 直接从文件名获取两个值
                        string longValue = parts[2];   // 长边值
                        string shortValue = parts[3];  // 短边值

                        // 写入长边值
                        var cellLong = worksheet.Cells[row, 6];
                        cellLong.Value = $"{longValue}";
                        SetResultColor(cellLong, result);


                        // 写入短边值
                        var cellShort = worksheet.Cells[row, 7];
                        cellShort.Value = $"{shortValue}";
                        SetResultColor(cellShort, result);

                    }
                
                    else
                    {
                        // 其他测量类型的处理保持不变
                        int column = GetMeasurementColumn(measurementType, parts[0]);
                        if (column > 0)
                        {
                            var cell = worksheet.Cells[row, column];
                            cell.Value = $"{parts[2]}";
                            SetResultColor(cell, result2);
                        }
                    }
                }
            }
        }

        private int GetMeasurementColumn(string measurementType, string serialNumber)
        {
            switch (measurementType.ToLower())
            {
                case "垂直度":
                    return 5;
                case "位置度":
                    return 8;
                case "错位值":
                    switch (serialNumber)
                    {
                        case "2": return 9;
                        case "5": return 10;
                        case "6": return 11;
                        case "7": return 12;
                        default: return 0;
                    }
                case "缺料值":
                    switch (serialNumber)
                    {
                        case "2": return 13;
                        case "3": return 14;
                        case "4": return 15;
                        case "5": return 16;
                        case "6": return 17;
                        case "7": return 18;
                        case "8": return 19;
                        case "9": return 20;
                        default: return 0;
                    }
                default:
                    return 0;
            }
        }

        private void ProcessImages(ExcelWorksheet worksheet, int row, string resultImagePath)
        {
            // 处理OK文件夹中的图片
            string okPath = Path.Combine(resultImagePath, "OK");
            string ngPath = Path.Combine(resultImagePath, "NG");

            int col = 21; // 从第5列开始放图片
            int imageIndex = 1; // 用于跟踪图片编号

            // 设置行高以适应图片
            worksheet.Row(row).Height = 60; // 根据图片大小调整行高

            if (Directory.Exists(okPath))
            {
                foreach (string imagePath in Directory.GetFiles(okPath, "*.jpg"))
                {
                    // 插入图片到Excel
                    var picture = worksheet.Drawings.AddPicture(
                        $"Image_{row}_{col}",
                        new FileInfo(imagePath)
                    );
                    picture.SetPosition(row - 1, 0, col - 1, 0);
                    picture.SetSize(60, 60);

                    worksheet.Cells[row, col + 1].Value = "OK";
                    worksheet.Cells[row, col + 1].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                    col += 2; // 每个图片占两列（图片+结果）
                    imageIndex++;
                }
            }

            if (Directory.Exists(ngPath))
            {
                foreach (string imagePath in Directory.GetFiles(ngPath, "*.jpg"))
                {
                    var picture = worksheet.Drawings.AddPicture(
                        $"Image_{row}_{col}",
                        new FileInfo(imagePath)
                    );
                    picture.SetPosition(row - 1, 0, col - 1, 0);
                    picture.SetSize(60, 60);

                    worksheet.Cells[row, col + 1].Value = "NG";
                    worksheet.Cells[row, col + 1].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                    col += 2;
                    imageIndex++;
                }
            }
        }

        private void SetResultColor(ExcelRange cell, string result)
        {
            if (result.ToUpper() == "NG")
            {
                cell.Style.Font.Color.SetColor(System.Drawing.Color.Red);
            }
            if (result.ToUpper() == "OK")
            {
                cell.Style.Font.Color.SetColor(System.Drawing.Color.Green);
            }
        }

        private void SetResultColor2(ExcelRange cell, string result2)
        {
            if (result2.ToUpper() == "NG")
            {
                cell.Style.Font.Color.SetColor(System.Drawing.Color.Red);
            }
            if (result2.ToUpper() == "OK")
            {
                cell.Style.Font.Color.SetColor(System.Drawing.Color.Green);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
