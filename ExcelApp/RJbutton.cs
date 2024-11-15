using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Drawing2D;

using System.Windows.Forms;


namespace ExcelApp
{


    public class RjButton : System.Windows.Forms.Button
    {
        private int borderSize = 0;
        private int borderRadius = 40;
        private Color borderColour = Color.Red; // Default border color

        public RjButton()
        {
            this.FlatStyle = FlatStyle.Flat;
            this.FlatAppearance.BorderSize = 0;
            this.Size = new Size(150, 40);
            this.BackColor = Color.Red;
            this.ForeColor = Color.Black;
        }
        private Color gradientStartColor = Color.DeepSkyBlue;  // Default start color
        private Color gradientEndColor = Color.Aquamarine;    // Default end color

        public Color GradientStartColor
        {
            get { return gradientStartColor; }
            set { gradientStartColor = value; Invalidate(); }
        }

        public Color GradientEndColor
        {
            get { return gradientEndColor; }
            set { gradientEndColor = value; Invalidate(); }
        }

        // Public property to set the border color
        public Color BorderColor
        {
            get { return borderColour; }
            set { borderColour = value; this.Invalidate(); } // Invalidate to redraw with new border color
        }

        public int BorderSize { get => borderSize; set => borderSize = value;
        }
        public int BorderRadius { get => borderRadius; set => borderRadius = value; }
        public Color BorderColour { get => borderColour; set => borderColour = value; }

        // Method to create a rounded rectangle path
        private GraphicsPath GetFigurePath(RectangleF rect, float radius)
        {
            GraphicsPath path = new GraphicsPath();
            path.StartFigure();
            path.AddArc(rect.X, rect.Y, radius, radius, 180, 90);
            path.AddArc(rect.Width - radius, rect.Y, radius, radius, 270, 90);
            path.AddArc(rect.Width - radius, rect.Height - radius, radius, radius, 0, 90);
            path.AddArc(rect.X, rect.Height - radius, radius, radius, 90, 90);
            path.CloseFigure();

            return path;
        }

        // Override the OnPaint method to customize the button's appearance
        protected override void OnPaint(PaintEventArgs pevent)
        {
            base.OnPaint(pevent);
            pevent.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            // Define the surface area of the button
            RectangleF rectSurface = new RectangleF(0, 0, this.Width, this.Height);
            // Define the border area, slightly smaller to fit the border inside
            RectangleF rectBorder = new RectangleF(1, 1, this.Width - 2, this.Height - 2);

            // Apply the gradient fill for the button background
            using (LinearGradientBrush brush = new LinearGradientBrush(
                rectSurface, gradientStartColor, GradientEndColor, 45f)) // 45 degrees gradient
            {
                pevent.Graphics.FillRectangle(brush, rectSurface);
            }

            // Draw the border with rounded corners if the borderRadius is greater than 2
            if (borderRadius > 2)
            {
                using (GraphicsPath pathSurface = GetFigurePath(rectSurface, borderRadius))
                using (GraphicsPath pathBorder = GetFigurePath(rectBorder, borderRadius - 1F))
                using (Pen penSurface = new Pen(this.Parent.BackColor, 2)) // "Shadow" effect, using parent color
                using (Pen penBorder = new Pen(borderColour, borderSize)) // The actual border
                {
                    penBorder.Alignment = PenAlignment.Inset;
                    this.Region = new Region(pathSurface); // Set region to rounded shape
                    pevent.Graphics.DrawPath(penSurface, pathSurface); // Draw the surface path (with parent color)
                    if (borderSize >= 1)
                        pevent.Graphics.DrawPath(penBorder, pathBorder); // Draw the actual border
                }
            }
            else
            {
                // If no rounded corners, just a square shape
                this.Region = new Region(rectSurface);
                if (borderSize >= 1)
                {
                    using (Pen penBorder = new Pen(borderColour, borderSize)) // Regular border for square button
                    {
                        penBorder.Alignment = PenAlignment.Inset;
                        pevent.Graphics.DrawRectangle(penBorder, 0, 0, this.Width - 1, this.Height - 1);
                    }
                }
            }

            // Create a StringFormat to center the text
            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;  // Horizontally center the text
            stringFormat.LineAlignment = StringAlignment.Center;  // Vertically center the text

            // Draw the text centered in the button using DrawString
            pevent.Graphics.DrawString(this.Text, this.Font, new SolidBrush(this.ForeColor), rectSurface, stringFormat);
        }






        // Event handler to update the button when the parent's BackColor changes (for design-time updates)
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            this.Parent.BackColorChanged += new EventHandler(Container_BackColorChanged);
        }

        // Event to handle the parent's BackColor change
        private void Container_BackColorChanged(object sender, EventArgs e)
        {
            if (this.DesignMode)
                this.Invalidate();

        }

       

     

   
}


    //private void timer2_Tick(object sender, EventArgs e)
    //{



    //    progressBar1.Increment(1);
    //    label5.Text = progressBar1.Value.ToString() + "%";

    //}
}

