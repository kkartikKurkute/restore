using System;
using System.Drawing;
using System.Windows.Forms;

public class GradientButton : Button
{
    // Define the colors for the gradient
    public Color StartColor { get; set; } = Color.LightBlue;
    public Color EndColor { get; set; } = Color.DarkBlue;

    protected override void OnPaint(PaintEventArgs pevent)
    {
        base.OnPaint(pevent);

        // Create a linear gradient brush with the defined start and end colors
        using (LinearGradientBrush brush = new LinearGradientBrush(
            this.ClientRectangle, StartColor, EndColor, LinearGradientMode.Vertical))
        {
            // Fill the button background with the gradient
            pevent.Graphics.FillRectangle(brush, this.ClientRectangle);
        }

        // Draw the button text (centered)
        using (SolidBrush textBrush = new SolidBrush(this.ForeColor))
        {
            StringFormat stringFormat = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };
            pevent.Graphics.DrawString(this.Text, this.Font, textBrush, this.ClientRectangle, stringFormat);
        }
    }

    // Optional: Add hover effect to change the gradient
    protected override void OnMouseEnter(EventArgs e)
    {
        base.OnMouseEnter(e);
        StartColor = Color.DarkBlue;
        EndColor = Color.LightBlue;
        this.Invalidate();
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        base.OnMouseLeave(e);
        StartColor = Color.LightBlue;
        EndColor = Color.DarkBlue;
        this.Invalidate();
    }
}

public class MainForm : Form
{
    public MainForm()
    {
        // Create and add the gradient button to the form
        GradientButton gradientButton = new GradientButton
        {
            Text = "Click Me!",
            Size = new Size(200, 50),
            Location = new Point(100, 100),
            Font = new Font("Arial", 14),
            ForeColor = Color.White
        };

        this.Controls.Add(gradientButton);
    }

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}
