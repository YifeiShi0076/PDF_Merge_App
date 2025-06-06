using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Microsoft.Win32;
using System.Diagnostics;
using PDF_Merge_App.Model;

namespace PDF_Merge_App
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		private List<FileItem> fileList = new List<FileItem>();

		public MainWindow()
		{
			InitializeComponent();
		}

		private void DragDropArea_DragOver(object sender, DragEventArgs e)
		{
			// 设置拖放效果
			e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
			e.Handled = true;
		}

		private void DragDropArea_Drop(object sender, DragEventArgs e)
		{
			if (e.Data.GetData(DataFormats.FileDrop) is string[] files)
			{
				foreach (var file in files)
				{
					if (IsValidFile(file))
					{
						var fileName = System.IO.Path.GetFileName(file);
						fileList.Add(new FileItem { FileName = fileName, FilePath = file });
						FileListBox.ItemsSource = null; // 重新绑定数据源
						FileListBox.ItemsSource = fileList;
					}
					else
					{
						MessageBox.Show($"不支持该文件鸭~: {file}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
					}
				}
			}
		}

		private static bool IsValidFile(string file)
		{
			string[] allowedExtensions = { ".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".gif" };
			string extension = System.IO.Path.GetExtension(file).ToLower();
			return allowedExtensions.Contains(extension);
		}

		private void BrowseButton_Click(object sender, RoutedEventArgs e)
		{
			// 打开保存文件对话框
			SaveFileDialog saveFileDialog = new SaveFileDialog
			{
				Filter = "PDF 文件 (*.pdf)|*.pdf",
				Title = "选择保存路径"
			};

			if (saveFileDialog.ShowDialog() == true)
			{
				OutputPathTextBox.Text = saveFileDialog.FileName;
			}
		}

		private void DeleteButton_Click(object sender, RoutedEventArgs e)
		{
			if (sender is Button button && button.Tag is string filePath)
			{
				// 从文件列表中移除对应文件
				var fileToRemove = fileList.FirstOrDefault(item => item.FilePath == filePath);
				if (fileToRemove != null)
				{
					fileList.Remove(fileToRemove);
					// 更新 ListView 数据源
					FileListBox.ItemsSource = null;
					FileListBox.ItemsSource = fileList;
				}
			}
		}

		private void MergeButton_Click(object sender, RoutedEventArgs e)
		{
			if (fileList.Count == 0)
			{
				MessageBox.Show("木有文件鸭 (╯︵╰)", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
				return;
			}

			if (string.IsNullOrWhiteSpace(OutputPathTextBox.Text))
			{
				MessageBox.Show("不设置输出路径，我可是会私吞的哦 (*･ω< ) ", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
				return;
			}

			try
			{
				// 合并文件
				MergeFiles(fileList, OutputPathTextBox.Text);

				// 弹出包含“是”和“否”的提示框，询问用户是否打开文件位置
				var result = MessageBox.Show(
					"PDF合并成功！要打开文件位置吗？ヽ(￣▽￣)ﾉ",
					"成功",
					MessageBoxButton.YesNo,
					MessageBoxImage.Information
				);

				if (result == MessageBoxResult.Yes)
				{
					OpenFileLocation(OutputPathTextBox.Text);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"被玩坏了！o(╥﹏╥)o {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void OpenFileLocation(string text)
		{
			try
			{
				// 使用 Explorer 打开文件所在位置并选中该文件
				Process.Start("explorer.exe", $"/select,\"{OutputPathTextBox.Text}\"");
			}
			catch (Exception ex)
			{
				MessageBox.Show($"无法打开文件位置 (Ｔ▽Ｔ): {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		[Obsolete]
		private void MergeFiles(List<FileItem> files, string outputPath)
		{
			List<string> file_names=files.Select(f => f.FilePath).ToList();
			using (PdfDocument outputPdf = new PdfDocument())
			{
				foreach (var file in file_names)
				{
					string extension = System.IO.Path.GetExtension(file).ToLower();
					if (extension == ".pdf")
					{
						// 合并 PDF 文件
						using (PdfDocument inputPdf = PdfReader.Open(file, PdfDocumentOpenMode.Import))
						{
							foreach (PdfPage page in inputPdf.Pages)
							{
								outputPdf.AddPage(page);
							}
						}
					}
					else
					{
						// 将图片转换为 PDF
						using (XImage image = XImage.FromFile(file))
						{
							PdfPage page = outputPdf.AddPage();
							page.Width = image.PixelWidth * 72 / image.HorizontalResolution;
							page.Height = image.PixelHeight * 72 / image.VerticalResolution;

							using (XGraphics gfx = XGraphics.FromPdfPage(page))
							{
								gfx.DrawImage(image, 0, 0, page.Width, page.Height);
							}
						}
					}
				}

				outputPdf.Save(outputPath);
			}
		}
	}
}