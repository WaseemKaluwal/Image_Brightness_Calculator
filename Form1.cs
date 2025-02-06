using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;  // EPPlus library for Excel

namespace ImageBrightnessCalculator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            InitializeCustomComponents();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private PictureBox pictureBox;
        private ComboBox pixelFormatComboBox;
        private Button selectImageButton;
        private Button calculateBrightnessButton;
        private Button reduceBrightnessButton;
        private Button savePixelDataButton;  // New button to save pixel data
        private ComboBox brightnessComboBox;
        private Label resultLabel;
        private Label originalFormatLabel;
        private Label brightnessPercentageLabel;
        private Bitmap originalImage;  // To store original image

        private void InitializeCustomComponents()
        {
            this.Text = "Image Brightness Calculator";
            this.Size = new Size(600, 600); // Increased form size for the new button

            pictureBox = new PictureBox
            {
                BorderStyle = BorderStyle.FixedSingle,
                Size = new Size(400, 200),
                Location = new Point(100, 20),
                SizeMode = PictureBoxSizeMode.Zoom
            };
            this.Controls.Add(pictureBox);

            pixelFormatComboBox = new ComboBox
            {
                Location = new Point(100, 240),
                Width = 200
            };
            pixelFormatComboBox.Items.AddRange(new string[] { "8bpp", "12bpp", "16bpp", "24bpp", "32bpp", "48bpp" });
            pixelFormatComboBox.SelectedIndex = 2; // Default to 24bpp
            this.Controls.Add(pixelFormatComboBox);

            selectImageButton = new Button
            {
                Text = "Select Image",
                Location = new Point(320, 240),
                Width = 100
            };
            selectImageButton.Click += SelectImageButton_Click;
            this.Controls.Add(selectImageButton);

            calculateBrightnessButton = new Button
            {
                Text = "Calculate Brightness",
                Location = new Point(100, 280),
                Width = 320
            };
            calculateBrightnessButton.Click += CalculateBrightnessButton_Click;
            this.Controls.Add(calculateBrightnessButton);

            reduceBrightnessButton = new Button
            {
                Text = "Reduce Brightness",
                Location = new Point(100, 320),
                Width = 320
            };
            reduceBrightnessButton.Click += ReduceBrightnessButton_Click;
            this.Controls.Add(reduceBrightnessButton);

            brightnessComboBox = new ComboBox
            {
                Location = new Point(100, 360),
                Width = 200
            };
            brightnessComboBox.Items.AddRange(new string[] { "10%", "20%", "30%", "40%", "50%", "60%", "70%", "80%", "90%" });
            brightnessComboBox.SelectedIndex = 4; // Default to 50%
            this.Controls.Add(brightnessComboBox);

            savePixelDataButton = new Button
            {
                Text = "Save Pixel Data to Excel",
                Location = new Point(100, 400),
                Width = 320
            };
            savePixelDataButton.Click += SavePixelDataButton_Click;
            this.Controls.Add(savePixelDataButton);

            resultLabel = new Label
            {
                Text = "Brightness in Pixel: ",
                Location = new Point(100, 440),
                AutoSize = true
            };
            this.Controls.Add(resultLabel);

            originalFormatLabel = new Label
            {
                Text = "Original Format: ",
                Location = new Point(100, 480),
                AutoSize = true
            };
            this.Controls.Add(originalFormatLabel);

            brightnessPercentageLabel = new Label
            {
                Text = "Brightness Percentage: ",
                Location = new Point(100, 500),
                AutoSize = true
            };
            this.Controls.Add(brightnessPercentageLabel);
        }

        private void SelectImageButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Image Files|*.bmp;*.jpg;*.jpeg;*.png;*.tiff";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    originalImage = new Bitmap(openFileDialog.FileName);  // Save the original image
                    pictureBox.Image = originalImage;
                    originalFormatLabel.Text = $"Original Format: {originalImage.PixelFormat}";
                }
            }
        }

        private void CalculateBrightnessButton_Click(object sender, EventArgs e)
        {
            if (pictureBox.Image == null)
            {
                MessageBox.Show("Please select an image first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string selectedFormat = pixelFormatComboBox.SelectedItem.ToString();
            Bitmap image = new Bitmap(pictureBox.Image);

            double totalBrightness = CalculateTotalBrightness(image, selectedFormat);
            double brightnessPercentage = CalculateBrightnessPercentage(totalBrightness, image.Width, image.Height);
            resultLabel.Text = $"Brightness: {totalBrightness}";
            brightnessPercentageLabel.Text = $"Brightness Percentage: {brightnessPercentage}%";
        }

        private double CalculateTotalBrightness(Bitmap image, string pixelFormat)
        {
            double totalBrightness = 0;

            var rect = new Rectangle(0, 0, image.Width, image.Height);
            var bmpData = image.LockBits(rect, ImageLockMode.ReadOnly, image.PixelFormat);

            try
            {
                int stride = Math.Abs(bmpData.Stride);
                int bytesPerPixel = Image.GetPixelFormatSize(image.PixelFormat) / 8;
                byte[] pixelBuffer = new byte[stride * image.Height];

                System.Runtime.InteropServices.Marshal.Copy(bmpData.Scan0, pixelBuffer, 0, pixelBuffer.Length);

                for (int y = 0; y < image.Height; y++)
                {
                    for (int x = 0; x < image.Width; x++)
                    {
                        int offset = y * stride + x * bytesPerPixel;

                        byte blue = pixelBuffer[offset];
                        byte green = bytesPerPixel > 1 ? pixelBuffer[offset + 1] : blue;
                        byte red = bytesPerPixel > 2 ? pixelBuffer[offset + 2] : blue;

                        int pixelValue = 0;

                        if (pixelFormat == "8bpp")
                        {
                            pixelValue = blue;
                        }
                        else if (pixelFormat == "12bpp")
                        {
                            pixelValue = (int)((red + green + blue) / 3 * (4095.0 / 4095.0));
                        }
                        else if (pixelFormat == "16bpp")
                        {
                            pixelValue = (int)((red + green + blue) / 3 * (65535.0 / 65535.0));
                        }
                        else if (pixelFormat == "24bpp")
                        {
                            pixelValue = (red + green + blue) / 3;
                        }
                        else if (pixelFormat == "32bpp")
                        {
                            pixelValue = (red + green + blue) / 3;
                        }
                        else if (pixelFormat == "48bpp")
                        {
                            pixelValue = (int)((red * 256 + green * 256 + blue * 256) / 3);
                        }
                        else
                        {
                            throw new NotSupportedException("Unsupported pixel format.");
                        }

                        totalBrightness += pixelValue;
                    }
                }
            }
            finally
            {
                image.UnlockBits(bmpData);
            }

            return totalBrightness;
        }


        private double CalculateBrightnessPercentage(double totalBrightness, int width, int height)
        {
            double maxBrightness = 255 * width * height;
            double brightnessPercentage = (totalBrightness / maxBrightness) * 100;
            return brightnessPercentage;
        }

        private void ReduceBrightnessButton_Click(object sender, EventArgs e)
        {
            if (pictureBox.Image == null)
            {
                MessageBox.Show("Please select an image first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Get the selected brightness reduction percentage from the ComboBox
            string selectedPercentage = brightnessComboBox.SelectedItem.ToString();
            int percentage = int.Parse(selectedPercentage.Replace("%", ""));

            Bitmap image = new Bitmap(pictureBox.Image);

            // Reduce the brightness of the image by the selected percentage
            Bitmap reducedBrightnessImage = ReduceImageBrightness(image, percentage);
            pictureBox.Image = reducedBrightnessImage;  // Display the modified image in PictureBox
        }

        private Bitmap ReduceImageBrightness(Bitmap image, int percentage)
        {
            Bitmap newImage = new Bitmap(image.Width, image.Height);
            for (int x = 0; x < image.Width; x++)
            {
                for (int y = 0; y < image.Height; y++)
                {
                    Color pixelColor = image.GetPixel(x, y);

                    int red = (int)(pixelColor.R * (1 - percentage / 100.0));
                    int green = (int)(pixelColor.G * (1 - percentage / 100.0));
                    int blue = (int)(pixelColor.B * (1 - percentage / 100.0));

                    // Ensure that the color values are within the valid range [0, 255]
                    red = Math.Max(0, Math.Min(255, red));
                    green = Math.Max(0, Math.Min(255, green));
                    blue = Math.Max(0, Math.Min(255, blue));

                    newImage.SetPixel(x, y, Color.FromArgb(red, green, blue));
                }
            }

            return newImage;
        }

        private void SavePixelDataButton_Click(object sender, EventArgs e)
        {
            if (originalImage == null || pictureBox.Image == null)
            {
                MessageBox.Show("Please select an image and apply brightness reduction first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Create a new Excel package using EPPlus
            using (var package = new ExcelPackage())
            {
                // Create a worksheet
                var worksheet = package.Workbook.Worksheets.Add("PixelData");

                // Get the reduced brightness image from the PictureBox
                Bitmap reducedImage = new Bitmap(pictureBox.Image);

                // Save both "Before" and "After" brightness data in one sheet
                SavePixelDataToWorksheet(originalImage, reducedImage, worksheet);

                // Save the file
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    FileName = "PixelData.xlsx"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    package.SaveAs(saveFileDialog.FileName);
                    MessageBox.Show("Pixel data saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }


        private void SavePixelDataToWorksheet(Bitmap beforeImage, Bitmap afterImage, ExcelWorksheet worksheet)
        {
            int row = 1; // Start from the first row
            int colBefore = 1; // Column A for "Before Brightness Reduction"
            int colAfter = 2; // Column B for "After Brightness Reduction"

            // Add headers for the columns
            worksheet.Cells[row, colBefore].Value = "Before Brightness Reduction";
            worksheet.Cells[row, colAfter].Value = "After Brightness Reduction";

            row++; // Move to the next row for data

            // Ensure the images have the same dimensions
            if (beforeImage.Width != afterImage.Width || beforeImage.Height != afterImage.Height)
            {
                MessageBox.Show("Images have different dimensions.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            for (int y = 0; y < beforeImage.Height; y++)
            {
                for (int x = 0; x < beforeImage.Width; x++)
                {
                    // Get the pixel color before and after brightness adjustment
                    Color beforePixel = beforeImage.GetPixel(x, y);
                    Color afterPixel = afterImage.GetPixel(x, y);

                    // Calculate brightness as the average of R, G, and B
                    int beforeBrightness = (beforePixel.R + beforePixel.G + beforePixel.B) / 3;
                    int afterBrightness = (afterPixel.R + afterPixel.G + afterPixel.B) / 3;

                    // Write brightness values into the worksheet
                    worksheet.Cells[row, colBefore].Value = beforeBrightness;
                    worksheet.Cells[row, colAfter].Value = afterBrightness;

                    // Move to the next row
                    row++;

                    // Check if the Excel row limit is exceeded
                    if (row > 1048576)
                    {
                        MessageBox.Show("Data exceeds Excel row limit. Consider reducing the image resolution.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }


    }
}
